import pandas as pd
import os
import argparse
import sys

def pilah_data_satker(file_master, folder_hasil="hasil_pilahan"):
    # 1. Pastikan folder hasil ada, jika belum buat foldernya
    if not os.path.exists(folder_hasil):
        os.makedirs(folder_hasil)
        print(f"[*] Folder '{folder_hasil}' berhasil dibuat.")

    try:
        # 2. Baca file Excel master
        print(f"[*] Membaca file {file_master}...")
        df = pd.read_excel(file_master)
        
        # Nama kolom yang ingin digunakan sebagai acuan pemisahan
        kolom_target = 'Nama Satker'

        # 3. Cek apakah kolom tersebut ada di dalam data (case-sensitive)
        if kolom_target not in df.columns:
            print(f"[!] Error: Kolom '{kolom_target}' tidak ditemukan dalam file Excel.")
            print("[!] Kolom yang tersedia:", ", ".join(df.columns))
            sys.exit(1)

        # 4. Ambil daftar unik dari kolom target
        daftar_satker = df[kolom_target].dropna().unique()
        
        print(f"[*] Ditemukan {len(daftar_satker)} kategori yang berbeda di kolom '{kolom_target}'. Memulai proses pemilahan...\n")

        # 5. Lakukan looping untuk setiap satker
        for satker in daftar_satker:
            # Filter data khusus untuk satker yang sedang diproses
            df_satker = df[df[kolom_target] == satker]
            
            # Buat nama file yang aman berdasarkan nama satker
            nama_aman = str(satker).replace("/", "_").replace("\\", "_").replace(":", "_").replace("\"", "_")
            
            # Buat nama file langsung di dalam folder hasil (tanpa folder tambahan)
            nama_file = os.path.join(folder_hasil, f"Data_{nama_aman}.xlsx")
            
            # Simpan data (isi utuh) ke Excel baru
            df_satker.to_excel(nama_file, index=False)
            print(f"  -> Berhasil menyimpan file: {nama_file}")
            
        print("\n[*] Selesai! Semua data telah berhasil dipilah dan disimpan.")

    except FileNotFoundError:
        print(f"[!] Error: File '{file_master}' tidak ditemukan. Pastikan nama file dan lokasinya benar.")
        sys.exit(1)
    except Exception as e:
        print(f"[!] Terjadi kesalahan: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # Setup argparse untuk membaca argumen dari perintah terminal
    parser = argparse.ArgumentParser(description="Script untuk memilah data Excel berdasarkan kolom 'Satker'.")
    
    # Argumen wajib: Nama file Excel master
    parser.add_argument("file_master", help="Nama/path file Excel master (contoh: data_master.xlsx)")
    
    # Argumen opsional: Folder output
    parser.add_argument("-o", "--output", default="hasil_pilahan", 
                        help="Nama folder tempat menyimpan hasil (default: hasil_pilahan)")
    
    # Parsing argumen
    args = parser.parse_args()
    
    # Menjalankan fungsi dengan data yang dilempar dari terminal
    pilah_data_satker(args.file_master, args.output)

