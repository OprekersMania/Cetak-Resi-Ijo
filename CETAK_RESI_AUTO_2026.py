import os
import math
import tkinter as tk
from tkinter import Toplevel, filedialog, messagebox, ttk
import fitz  # PyMuPDF
import win32print
import win32ui
import cv2
import numpy as np
from PIL import Image, ImageWin
from pdf2image import convert_from_path

# =========================================================
# KONFIGURASI GLOBAL
# =========================================================
POPPLER_PATH = r"C:\poppler\Library\bin"
DPI = 150
PADDING = 1
MM_MARGIN = 3
MARGIN_PX = int((MM_MARGIN / 25.4) * DPI)

# Seting ke 0 untuk tes tanpa print
TESTING = 1

# Hidden Root untuk Tkinter Dialog
root = tk.Tk()
root.withdraw()


# =========================================================
# HELPER PEMROSESAN PDF (PYMUPDF)
# =========================================================
def Geser(page, rect, delta_x=0, delta_y=0, dpi=600):
    
    pix = page.get_pixmap(clip=rect, dpi=dpi)
    # 2. Hapus area lama dengan warna putih
    # page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
    
    
    # Hapus area permanen
    page.add_redact_annot(rect)
    page.apply_redactions()

    new_rect = fitz.Rect(
        rect.x0 + delta_x,
        rect.y0 + delta_y,
        rect.x1 + delta_x,
        rect.y1 + delta_y,
    )
    
    page.insert_image(new_rect, pixmap=pix)

def gambar_garis(page, jenis, posisi, awal, akhir, tebal=1, warna=(0, 0, 0)):
    """Menggambar garis pembatas horizontal atau vertikal pada halaman PDF."""
    jenis = jenis.lower()
    if jenis in ["h", "horizontal"]:
        p1 = fitz.Point(awal, posisi)
        p2 = fitz.Point(akhir, posisi)
    elif jenis in ["v", "vertikal"]:
        p1 = fitz.Point(posisi, awal)
        p2 = fitz.Point(posisi, akhir)
    else:
        return

    page.draw_line(p1, p2, color=warna, width=tebal, stroke_opacity=1.0)

def pindah_dan_ubah_ukuran_gambar(page, rect_asal, delta_x=0, delta_y=0, lebar_baru=None, tinggi_baru=None, dpi=600):

    # Mengambil gambar di koordinat lama, menghapusnya, lalu menempelkannya kembali
    # di posisi baru dengan dimensi ukuran yang bisa disesuaikan (Resize).
    
    # Parameters:
    # - page        : Objek halaman PDF (fitz.Page)
    # - rect_asal   : Koordinat kotak gambar asli dari log (fitz.Rect)
    # - delta_x     : Jarak geser horizontal (minus = kiri, plus = kanan)
    # - delta_y     : Jarak geser vertikal (minus = atas, plus = bawah)
    # - lebar_baru  : Lebar baru dalam pt (Kosongkan jika ingin pakai lebar asli)
    # - tinggi_baru : Tinggi baru dalam pt (Kosongkan jika ingin pakai tinggi asli)
    # - dpi         : Resolusi jepretan agar gambar/barcode tetap tajam (default: 600)
    
    # 1. Snapshot area gambar asli dengan resolusi tinggi agar tidak buram
    pix = page.get_pixmap(clip=rect_asal, dpi=dpi, colorspace=fitz.csRGB)
    
    # 2. Hapus (tipex putih) area gambar yang lama agar bersih
    page.draw_rect(rect_asal, color=(1, 1, 1), fill=(1, 1, 1), width=0, stroke_opacity=0)
    
    # 3. Tentukan dimensi lebar dan tinggi akhir kotak baru
    lebar_akhir = lebar_baru if lebar_baru is not None else (rect_asal.x1 - rect_asal.x0)
    tinggi_akhir = tinggi_baru if tinggi_baru is not None else (rect_asal.y1 - rect_asal.y0)

    # 4. Hitung titik koordinat awal baru (X0, Y0) berdasarkan delta geser
    x0_baru = rect_asal.x0 + delta_x
    y0_baru = rect_asal.y0 + delta_y
    
    # 5. Bangun kotak tujuan baru (X1 dan Y1 dihitung dari dimensi ukuran baru)
    rect_tujuan_baru = fitz.Rect(
        x0_baru,
        y0_baru,
        x0_baru + lebar_akhir,
        y0_baru + tinggi_akhir
    )
    
    # 6. Tempelkan kembali gambar hasil snapshot ke kotak baru tersebut
    page.insert_image(rect_tujuan_baru, pixmap=pix)

def hapus_objek_gambar(doc, page, x0_target, y0_target):
    """Menghapus elemen gambar berdasarkan koordinat awal."""
    page.clean_contents()
    info_gambar = page.get_image_info(xrefs=True)

    id_gambar_ditemukan = None
    for img in info_gambar:
        x0, y0, _, _ = img["bbox"]
        if abs(x0 - x0_target) < 3 and abs(y0 - y0_target) < 3:
            id_gambar_ditemukan = img["xref"]
            break

    if id_gambar_ditemukan and id_gambar_ditemukan > 0:
        page.delete_image(id_gambar_ditemukan)
        doc._deleteObject(id_gambar_ditemukan)
        page.clean_contents()

def data_semua_teks(doc):
    """Mengambil seluruh data teks tiap halaman untuk analisis jenis resi."""
    data_pdf = {}
    for i, page in enumerate(doc, 1):
        data_pdf[i] = []
        for block in page.get_text("blocks"):
            x0, y0, x1, y1, text, _, block_type = block
            if text.strip() and block_type == 0:
                data_pdf[i].append(
                    {
                        "text": text.strip(),
                        "x0": round(x0, 2),
                        "y0": round(y0, 2),
                        "x1": round(x1, 2),
                        "y1": round(y1, 2),
                    }
                )
    return data_pdf

def cek_logika_area(data_halaman, x_min, x_max, y_min, y_max):
    """Mengecek keberadaan teks pada area tertentu."""
    for item in data_halaman:
        if x_min <= item["x0"] <= x_max and y_min <= item["y0"] <= y_max:
            return item["text"]
    return None

# =========================================================
# LOGIKA EDIT RESI PER EKSPEDISI
# =========================================================
def prosesJT(doc, no_halaman, daftar_teks):
    page = doc[no_halaman - 1]

    # Hapus teks pinggir & S&K
    page.add_redact_annot(fitz.Rect(7.5, 33.5, 16.5, 167.5))
    page.add_redact_annot(fitz.Rect(281.5, 30.5, 291.5, 167.5))
    page.add_redact_annot(fitz.Rect(282.0, 1.5, 292.5, 10.5))
    page.add_redact_annot(fitz.Rect(52.2, 265.0, 232.0, 273.0))

    # Geser Pengirim & Penerima
    Geser(
        page, fitz.Rect(19.0, 69.5, 218.0, 86.2), delta_x=-1.0, delta_y=-12.0
    )
    Geser(
        page, fitz.Rect(19.0, 38.5, 218.0, 74.2), delta_x=0, delta_y=-5.0
    )

    # Ambil Barcode
    kotak_barcode = fitz.Rect(22, 205, 276, 240)
    pix_barcode = page.get_pixmap(clip=kotak_barcode, dpi=600)
    page.add_redact_annot(fitz.Rect(22, 205, 276, 247))

    antrean_teks_geser = []
    for blok in daftar_teks:
        text, x0, y0, x1, y1 = (
            blok["text"],
            blok["x0"],
            blok["y0"],
            blok["x1"],
            blok["y1"],
        )

        if 15.0 <= x0 <= 280.0 and 86.15 <= y0 <= 129.0:
            page.add_redact_annot(fitz.Rect(x0, y0, x1, y1))
            antrean_teks_geser.append(
                {
                    "rect": fitz.Rect(x0, y0 - 17, x1, y1),
                    "text": text,
                    "size": 8,
                    "font": "Times-Bold",
                    "align": 0,
                }
            )
        elif 182.0 <= y0 <= 186.0:
            page.add_redact_annot(fitz.Rect(x0, y0, x1, y1))
            antrean_teks_geser.append(
                {
                    "rect": fitz.Rect(x0, y0 + 24, x1, y1 + 30),
                    "text": text,
                    "size": 14,
                    "font": "Times-Bold",
                    "align": 1,
                }
            )
        elif 245.0 <= y0 <= 249.0:
            page.add_redact_annot(fitz.Rect(x0, y0, x1, y1))
            antrean_teks_geser.append(
                {
                    "rect": fitz.Rect(x0, y0 + 10, x1, y1 + 16),
                    "text": text,
                    "size": 14,
                    "font": "Times-Bold",
                    "align": 1,
                }
            )

    if antrean_teks_geser:
        page.apply_redactions()

    for item in antrean_teks_geser:
        page.insert_textbox(
            item["rect"],
            item["text"],
            fontsize=item["size"],
            fontname=item["font"],
            align=item["align"],
        )

    page.insert_image(kotak_barcode + (0, 18, 0, 18), pixmap=pix_barcode)
    Geser(
        page, fitz.Rect(10.0, 210.0, 285.0, 290), delta_x=0, delta_y=-37.5
    )
    Geser(
        page, fitz.Rect(10.0, 129.0, 285.0, 260.0), delta_x=0, delta_y=-24.5
    )

def prosesJNE(doc, no_halaman, daftar_teks):
    page = doc[no_halaman - 1]
    hapus_objek_gambar(doc, page, x0_target=6.5, y0_target=2.9)

    page.add_redact_annot(fitz.Rect(125.5, 175.5, 1715, 184.5))
    page.add_redact_annot(fitz.Rect(11.5, 158.5, 36.5, 168.5))
    page.apply_redactions()

    Geser(
        page, fitz.Rect(8.5, 71.5, 215.5, 82.5), delta_x=0, delta_y=-15.0
    )
    Geser(
        page, fitz.Rect(8.5, 81.5, 289.5, 107.5), delta_x=0, delta_y=-15.0
    )
    Geser(
        page, fitz.Rect(4.5, 112.5, 295.5, 158.5), delta_x=0, delta_y=-17.0
    )
    Geser(
        page, fitz.Rect(50.0, 184.0, 250.0, 220.0), delta_x=0, delta_y=-40.0
    )
    Geser(
        page, fitz.Rect(8.5, 233.5, 180.5, 243.5), delta_x=0, delta_y=-50.0
    )

    gambar_garis(
        page=page,
        jenis="v",
        posisi=101.5,
        awal=97.5,
        akhir=111.5,
        tebal=1.0,
        warna=(0, 0, 0),
    )
    gambar_garis(
        page=page,
        jenis="v",
        posisi=196.5,
        awal=97.5,
        akhir=142.0,
        tebal=1.0,
        warna=(0, 0, 0),
    )

    page.draw_rect(
        fitz.Rect(6.5, 97.5, 291.5, 111.5), color=(0, 0, 0), fill=None, width=1.0
    )
    page.draw_rect(
        fitz.Rect(6.5, 142, 291.5, 180.5), color=(0, 0, 0), fill=None, width=1.0
    )
    page.draw_rect(
        fitz.Rect(6.5, 3.0, 291.5, 194.5), color=(0, 0, 0), fill=None, width=1.0
    )

def prosesAnteraja(doc, no_halaman, daftar_teks):
    page = doc[no_halaman - 1]

    target = fitz.Rect(9.5, 33.5, 285.5, 45.0)
    page.draw_rect(target, color=(1, 1, 1), fill=(1, 1, 1), width=1.0)
    page.apply_redactions()

    kotak_gambar_asli = fitz.Rect(121.96, 6.30, 173.83, 27.11)
    pindah_dan_ubah_ukuran_gambar(
        page=page,
        rect_asal=kotak_gambar_asli,
        delta_x=0,
        delta_y=0,
        lebar_baru=51.87,
        tinggi_baru=15.81,
        dpi=600,
    )

    Geser(
        page, fitz.Rect(9.5, 118.5, 106.5, 128.5), delta_x=100, delta_y=-11.5
    )
    Geser(
        page, fitz.Rect(7.5, 158.5, 216.8, 225.5), delta_x=0, delta_y=-35.0
    )
    Geser(
        page, fitz.Rect(9.5, 230.5, 125.5, 240.5), delta_x=150, delta_y=44.0
    )
    Geser(page, fitz.Rect(5.5, 43.5, 296.5, 195.0), delta_x=0, delta_y=-12.0)
    Geser(
        page, fitz.Rect(5.5, 242.5, 296.5, 330.5), delta_x=0, delta_y=-64.0
    )

def prosesGojek(doc, no_halaman, daftar_teks):
    page = doc[no_halaman - 1]

    Geser(
        page, fitz.Rect(5.0, 160.0, 297.0, 220.0), delta_x=0, delta_y=-25.0
    )
    # Konversi mm ke poin PDF (1 mm = 2.83465 pt)
    MM_TO_PT = 2.83465

    crop_top = 1 * MM_TO_PT      # 3 mm
    crop_left = 2 * MM_TO_PT     # 5 mm
    crop_right = 2 * MM_TO_PT    # 5 mm
    crop_bottom = 79 * MM_TO_PT  # 10 mm

    # Margin bingkai dari tepi dokumen ASLI (2 mm)
    border_top = 2 * MM_TO_PT      # 3 mm
    border_left = 2.5 * MM_TO_PT     # 5 mm
    border_right = 2.5 * MM_TO_PT    # 5 mm
    border_bottom = 80 * MM_TO_PT  # 10 mm

    # 1. Ambil batas dokumen fisik asli (MediaBox)
    media = page.mediabox
    rotation = page.rotation

    # 2. Gambar bingkai 2mm dari tepi dokumen ASLI (sebelum cropping)
    #    Kita masuk ke dalam sebesar 2mm dari tepi fisik MediaBox
    border_rect = fitz.Rect(
        media.x0 + border_left,
        media.y0 + border_top,
        media.x1 - border_right, 
        media.y1 - border_bottom
    )
    page.draw_rect(border_rect, color=(0, 0, 0), width=1)

    # 3. Hitung koordinat CropBox berdasarkan rotasi halaman
    if rotation == 0:
        x0 = media.x0 + crop_left
        y0 = media.y0 + crop_top
        x1 = media.x1 - crop_right
        y1 = media.y1 - crop_bottom
    elif rotation == 90:
        x0 = media.x0 + crop_bottom
        y0 = media.y0 + crop_left
        x1 = media.x1 - crop_top
        y1 = media.y1 - crop_right
    elif rotation == 180:
        x0 = media.x0 + crop_right
        y0 = media.y0 + crop_bottom
        x1 = media.x1 - crop_left
        y1 = media.y1 - crop_top
    elif rotation == 270:
        x0 = media.x0 + crop_top
        y0 = media.y0 + crop_right
        x1 = media.x1 - crop_bottom
        y1 = media.y1 - crop_left

    # 4. Potong (crop) halaman sesuai batas baru
    new_crop = fitz.Rect(x0, y0, x1, y1)
    page.set_cropbox(new_crop)


# =========================================================
# OLAH PDF
# =========================================================
def olah_pdf_input(pdf_path):
    output_path = os.path.join(
        os.path.dirname(pdf_path), f"temp_processed_{os.path.basename(pdf_path)}"
    )
    doc = fitz.open(pdf_path)
    database_teks = data_semua_teks(doc)

    for no_halaman, daftar_teks in database_teks.items():
        CEK_JT = cek_logika_area(daftar_teks, 52.2, 232.0, 265.0, 273.0)
        CEK_JNE = cek_logika_area(daftar_teks, 8.0, 47.0, 234.0, 243.0)
        CEK_ANTER = cek_logika_area(daftar_teks, 8.0, 48.0, 228.0, 242.0)
        CEK_GOJEK = cek_logika_area(daftar_teks, 15.0, 163.0,42.0, 171.0)

        if CEK_JT:
            print(f"[Halaman {no_halaman}] : RESI J&T")
            prosesJT(doc, no_halaman, daftar_teks)
        elif CEK_JNE:
            print(f"[Halaman {no_halaman}] : RESI JNE")
            prosesJNE(doc, no_halaman, daftar_teks)
        elif CEK_ANTER:
            print(f"[Halaman {no_halaman}] : RESI ANTER AJA")
            prosesAnteraja(doc, no_halaman, daftar_teks)
        elif CEK_GOJEK:
            print(f"[Halaman {no_halaman}] : RESI GOJEK")
            prosesGojek(doc, no_halaman, daftar_teks)
        else:
            print(f"[Halaman {no_halaman}] Jenis resi tidak dikenali.")

    doc.save(output_path, deflate=True)
    doc.close()
    return output_path

# =========================================================
# GUI SETTING PRINTER
# =========================================================
def select_print_option():
    printers = [p[2] for p in win32print.EnumPrinters(2)]
    default_printer = win32print.GetDefaultPrinter()

    result = {
        "printer": default_printer,
        "label_count": 2,
        "cancelled": True,
    }

    win = Toplevel()
    win.title("Print Settings")
    win.geometry("400x180")
    win.resizable(False, False)

    ttk.Label(win, text="Pilih Printer:").pack(pady=(15, 5))
    printer_combo = ttk.Combobox(
        win, values=printers, width=50, state="readonly"
    )
    printer_combo.pack()
    printer_combo.set(default_printer)

    ttk.Label(win, text="Jumlah label per Lembar Print:").pack(pady=(15, 5))
    label_combo = ttk.Combobox(
        win, values=["1", "2", "3", "4", "5"], width=10, state="readonly"
    )
    label_combo.pack()
    label_combo.set("2")

    def ok():
        result["printer"] = printer_combo.get()
        result["label_count"] = int(label_combo.get())
        result["cancelled"] = False
        win.destroy()

    def cancel():
        result["cancelled"] = True
        win.destroy()

    button_frame = ttk.Frame(win)
    button_frame.pack(pady=15)

    ttk.Button(button_frame, text="OK", width=10, command=ok).pack(
        side="left", padx=5
    )
    ttk.Button(button_frame, text="Cancel", width=10, command=cancel).pack(
        side="left", padx=5
    )

    win.protocol("WM_DELETE_WINDOW", cancel)
    win.grab_set()
    root.wait_window(win)

    return result

# =========================================================
# PRINT ENGINE
# =========================================================
def print_image(image_path, printer_name):
    print("\nPRINT TO:", printer_name)
    img = Image.open(image_path)
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)

    printable_width = hDC.GetDeviceCaps(8)  # HORZRES
    scale = printable_width / float(img.size[0])
    new_width = int(img.size[0] * scale)
    new_height = int(img.size[1] * scale)

    img = img.resize((new_width, new_height), Image.LANCZOS)
    dib = ImageWin.Dib(img)

    hDC.StartDoc(f"Print_{os.path.basename(image_path)}")
    hDC.StartPage()
    dib.draw(hDC.GetHandleOutput(), (0, 0, new_width, new_height))
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

# =========================================================
# MAIN PROGRAM
# =========================================================
def Main_Prog():
    pdf_path = filedialog.askopenfilename(
        title="Pilih File PDF", filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_path:
        print("Pengoperasian dibatalkan: File tidak dipilih.")
        return

    print("Memulai editing PDF...")
    processed_file = os.path.normpath(olah_pdf_input(pdf_path))


    if TESTING:
        settings = select_print_option()
        if settings["cancelled"]:
            print("Proses dibatalkan oleh pengguna.")
            return

        SELECTED_PRINTER = settings["printer"]
        MAX_LABEL_PER_IMAGE = settings["label_count"]

        base_name = os.path.splitext(os.path.basename(processed_file))[0]
        output_dir = os.path.dirname(processed_file)

        # Convert PDF to Image
        pages = convert_from_path(
            processed_file, dpi=DPI, poppler_path=POPPLER_PATH
        )

        cropped_images = []

        # Bounding Box Auto Crop
        for page_num, page in enumerate(pages):
            img = cv2.cvtColor(np.array(page), cv2.COLOR_RGB2BGR)
            original = img.copy()

            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            H, W = gray.shape
            _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)

            contours, _ = cv2.findContours(
                thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
            )

            if contours:
                all_pts = np.vstack(
                    [cnt for cnt in contours if cv2.contourArea(cnt) > 50]
                )
                if len(all_pts) > 0:
                    x, y, w, h = cv2.boundingRect(all_pts)
                    crop_x1 = max(x - PADDING, 0)
                    crop_y1 = max(y - PADDING, 0)
                    crop_x2 = min(x + w + PADDING, W)
                    crop_y2 = min(y + h + PADDING, H)

                    cropped_images.append(original[crop_y1:crop_y2, crop_x1:crop_x2])
                    continue

            cropped_images.append(original)

        if not cropped_images:
            messagebox.showerror("Error", "Tidak ada label berhasil dicrop.")
            return

        # Grouping & Resizing Lebar Gambar agar Seragam
        groups = [
            cropped_images[i : i + MAX_LABEL_PER_IMAGE]
            for i in range(0, len(cropped_images), MAX_LABEL_PER_IMAGE)
        ]

        for group_index, group_images in enumerate(groups):
            max_width = max(img.shape[1] for img in group_images)

            resized_group_images = []
            for img in group_images:
                h, w = img.shape[:2]
                if w != max_width:
                    new_h = int(h * (max_width / float(w)))
                    img_resized = cv2.resize(
                        img, (max_width, new_h), interpolation=cv2.INTER_CUBIC
                    )
                    resized_group_images.append(img_resized)
                else:
                    resized_group_images.append(img)

            total_height = sum(img.shape[0] for img in resized_group_images)
            if len(resized_group_images) > 1:
                total_height += MARGIN_PX * (len(resized_group_images) - 1)

            # Buat Canvas Putih
            final_image = np.full(
                (total_height, max_width, 3), 255, dtype=np.uint8
            )

            current_y = 0
            for img in resized_group_images:
                h, w = img.shape[:2]
                final_image[current_y : current_y + h, 0:max_width] = img
                current_y += h + MARGIN_PX

            # Simpan Gambar Hasil
            suffix = (
                "_CROP.png" if len(groups) == 1 else f"_CROP_{group_index + 1}.png"
            )
            out_file = os.path.join(output_dir, f"{base_name}{suffix}")
            cv2.imwrite(out_file, final_image)

            # Cetak
            print_image(out_file, SELECTED_PRINTER)

        print("\nProses Selesai!")

if __name__ == "__main__":
    Main_Prog()