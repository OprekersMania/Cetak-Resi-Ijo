import os
import tkinter as tk
from tkinter import filedialog
from collections import Counter
import pdfplumber

def kelompokkan_karakter_per_baris(daftar_karakter, toleransi_y=1.0):
    """
    Mengelompokkan karakter menjadi satu string/frame berdasarkan
    kesamaan koordinat Y (garis horizontal) dan tinggi font yang sama.
    """
    if not daftar_karakter:
        return []

    # Urutkan karakter dari atas ke bawah (top), lalu dari kiri ke kanan (x0)
    daftar_karakter.sort(key=lambda k: (k['top'], k['x0']))
    list_baris_final = []
    
    # PERBAIKAN UTAMA: Menggunakan indeks [0] untuk mengambil elemen pertama
    karakter_saat_ini = daftar_karakter[0]
    
    # Inisialisasi properti untuk blok pertama
    blok_saat_ini = {
        'text': karakter_saat_ini['text'],
        'x0': karakter_saat_ini['x0'],
        'top': karakter_saat_ini['top'],
        'x1': karakter_saat_ini['x1'],
        'bottom': karakter_saat_ini['bottom'],
        'font_size': round(karakter_saat_ini.get('size', 0), 2)
        'font_sizes_all': [round(karakter_saat_ini.get('size', 0), 2)]
    }

    for k in daftar_karakter[1:]:
        char_text = k['text']
        char_x0 = k['x0']
        char_top = k['top']
        char_x1 = k['x1']
        char_bottom = k['bottom']
        char_size = round(k.get('size', 0), 2)

        sama_baris = abs(char_top - blok_saat_ini['top']) <= toleransi_y
        sama_font = char_size == blok_saat_ini['font_size']

        if sama_baris and sama_font:
            if char_x0 - blok_saat_ini['x1'] > 1.5:
                blok_saat_ini['text'] += " "
            blok_saat_ini['text'] += char_text
            blok_saat_ini['x1'] = max(blok_saat_ini['x1'], char_x1)
            blok_saat_ini['bottom'] = max(blok_saat_ini['bottom'], char_bottom)
            blok_saat_ini['font_sizes_all'].append(char_size)
        else:
            list_baris_final.append(blok_saat_ini)
            blok_saat_ini = {
                'text': char_text, 
				'x0': char_x0, 
				'top': char_top, 
                'x1': char_x1, 
				'bottom': char_bottom, 
				'font_size': char_size,
                'font_sizes_all': [char_size]
            }
    list_baris_final.append(blok_saat_ini)
    return list_baris_final

def pilih_dan_proses_pdf():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Pilih File PDF Sumber",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_path:
        print("Pemilihan file dibatalkan.")
        return

    folder_sumber = os.path.dirname(file_path)
    nama_file_asli = os.path.splitext(os.path.basename(file_path))
    log_path = os.path.join(folder_sumber, f"log_deteksi_{nama_file_asli}.txt")

    try:
        with pdfplumber.open(file_path) as pdf, open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write("=== LOG DETEKSI FRAME TEKS (FONT & BARIS SAMA) ===\n")
            log_file.write(f"Sumber File : {file_path}\n\n")

            for index, halaman in enumerate(pdf.pages, 1):
                log_file.write(f"=== HALAMAN {index} ===\n\n")
                
                daftar_karakter = halaman.chars
                blok_teks_final = kelompokkan_karakter_per_baris(daftar_karakter)
                
                # --- TAHAP 1: Cari ukuran font mayoritas khusus di Kriteria 1 (Area Alamat) ---
                font_area_alamat = []
                for blok in blok_teks_final:
                    x0, y0 = blok['x0'], blok['top']
                    if (15.0 <= x0 <= 280.0) and (85.0 <= y0 <= 135.0):
                        if len(set(blok['font_sizes_all'])) == 1:
                            font_area_alamat.append(blok['font_size'])
                
                # Ambil nilai font terbanyak (modus)
                font_alamat_sah = None
                if font_area_alamat:
                    # Ambil angka font_size-nya saja dari hasil Counter
                    font_alamat_sah = Counter(font_area_alamat).most_common(1)[0][0]
                
                # --- TAHAP 2: Filtrasi dan Penulisan Log ---
                for blok in blok_teks_final:
                    teks = blok['text'].strip()
                    if not teks:
                        continue
                        
                    x0, y0 = blok['x0'], blok['top']
                    x1, y1 = blok['x1'], blok['bottom']
                    lebar = x1 - x0
                    tinggi = y1 - y0
                    font_size = blok['font_size']
                    
                    # Kriteria 4: Blokir teks "Syarat dan ketentuan..."
                    if (51.0 <= x0 <= 54.0) and (265.0 <= y0 <= 268.0) and (5.5 <= tinggi <= 6.5):
                        continue
                        
                    terpilih = False
                    target_kategori = ""

                    # Kriteria 1: Area Kotak Alamat (X: 19-275, Y: 87-130)
                    if (15.0 <= x0 <= 280.0) and (85.0 <= y0 <= 135.0):
                        # Hanya lolos jika ukuran font SAMA dengan mayoritas font alamat (misal 12.29 pt)
                        if font_size != font_alamat_sah:
                            continue
                        
                        terpilih = True
                        target_kategori = "Area 3 Baris Teks (87.0 - 130.0)"

                    # Kriteria 2: Kode Area (Y ~ 184.71, Tinggi ~ 18.25 pt)
                    elif (183.0 <= y0 <= 186.0) and (18.0 <= tinggi <= 18.5):
                        terpilih = True
                        target_kategori = "Target Kode Area (Y:184.71, Tinggi:18.25)"

                    # Kriteria 3: Nomor Resi (Y ~ 247.51, Tinggi ~ 18.44 pt)
                    elif (246.0 <= y0 <= 249.0) and (18.2 <= tinggi <= 18.6):
                        terpilih = True
                        target_kategori = "Target No Resi (Y:247.51, Tinggi:18.44)"

                    if terpilih:
                        log_file.write(f"Kategori Target: {target_kategori}\n")
                        log_file.write(f"Teks Blok     : '{teks}'\n")
                        log_file.write(f"Ukuran Font   : {font_size} pt\n")
                        log_file.write(f"Posisi (X, Y) : ({x0:.2f}, {y0:.2f})\n")
                        log_file.write(f"Ukuran Frame  : Lebar = {lebar:.2f} pt, Tinggi = {tinggi:.2f} pt\n")
                        log_file.write("-" * 60 + "\n")
            
            print(f"Proses selesai! Log disimpan di:\n-> {log_path}")

    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")

if __name__ == "__main__":
    pilih_dan_proses_pdf()
