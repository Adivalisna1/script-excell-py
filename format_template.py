import pandas as pd
import argparse
import os
import sys

def format_ke_template(file_input, file_output=None):
    if not os.path.exists(file_input):
        print(f"[!] Error: File '{file_input}' tidak ditemukan.")
        sys.exit(1)

    print(f"[*] Membaca file '{file_input}'...")
    try:
        df_sumber = pd.read_excel(file_input)
    except Exception as e:
        print(f"[!] Gagal membaca file Excel: {e}")
        sys.exit(1)

    # Ini adalah struktur kolom sesuai dengan foto template yang Anda kirimkan
    kolom_template = [
        'Tanggal', 'Uraian', 'Kategori', 'No Rekening', 
        'Judul Laporan', 'Akun', 'Nominal'
    ]

    # Kita buat dataframe kosong yang jumlah barisnya sama persis dengan file asli
    # Ini sangat penting agar datanya tidak hilang!
    df_hasil = pd.DataFrame(index=df_sumber.index)

    # Kamus pencarian cerdas untuk mencocokkan kolom
    kata_kunci = {
        'Tanggal': ['tanggal', 'tgl'],
        'Uraian': ['uraian', 'keterangan'],
        'Kategori': ['kategori'],
        'No Rekening': ['nomor rekening', 'no rek', 'rekening bank'],
        'Judul Laporan': ['judul', 'laporan', 'nama rekening'],
        'Akun': ['akun'],
        'Nominal': ['nominal', 'saldo', 'jumlah']
    }

    kolom_ditemukan = []
    
    # Periksa tiap kolom di template
    for col_template in kolom_template:
        ditemukan = False
        kunci_pencarian = kata_kunci.get(col_template, [col_template.lower()])
        
        # Cari di seluruh kolom file sumber
        for col_sumber in df_sumber.columns:
            # Jika salah satu kata kunci ada di nama kolom sumber (misal ada kata 'keterangan' di dalam 'KETERANGAN')
            if any(kunci in str(col_sumber).lower() for kunci in kunci_pencarian):
                df_hasil[col_template] = df_sumber[col_sumber]
                kolom_ditemukan.append(f"{col_template} (dari '{col_sumber}')")
                ditemukan = True
                break # Berhenti mencari jika sudah ketemu pasangannya
                
        if not ditemukan:
            # Jika tidak ada yang nyerempet/mirip sama sekali, biarkan kosong
            df_hasil[col_template] = ""
            
    # Pastikan urutan kolom sesuai dengan template
    df_hasil = df_hasil[kolom_template]
            
    print(f"[*] Berhasil mencocokkan {len(kolom_ditemukan)} kolom: {', '.join(kolom_ditemukan)}")

    # Otomatisasi penamaan file hasil jika tidak diinputkan di terminal
    if file_output is None:
        nama_folder_hasil = "Hasil_Template"
        if not os.path.exists(nama_folder_hasil):
            os.makedirs(nama_folder_hasil)
            
        # Ambil nama filenya saja (mengabaikan path asal)
        nama_file_asli = os.path.basename(file_input)
        nama_file, ext = os.path.splitext(nama_file_asli)
        
        # Gabungkan path folder hasil dengan nama file baru
        file_output = os.path.join(nama_folder_hasil, f"{nama_file}_Template{ext}")

    # Simpan hasil ke Excel
    print(f"[*] Menyimpan hasil cetakan ke '{file_output}'...")
    try:
        df_hasil.to_excel(file_output, index=False)
        print("[*] Berhasil! Data telah berhasil disesuaikan dengan template.")
    except Exception as e:
        print(f"[!] Gagal menyimpan file: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script untuk mencetak ulang Excel sesuai Template.")
    parser.add_argument("file_input", help="Nama/path file Excel mentahan (contoh: data_mentah.xlsx)")
    parser.add_argument("-o", "--output", help="Nama file hasil setelah jadi (opsional)")
    
    args = parser.parse_args()
    
    format_ke_template(args.file_input, args.output)
