import pandas as pd
import argparse
import sys

def filter_kode_satker(file_master, file_output="Daftar_Kode_Satker_Unik.xlsx"):
    try:
        # 1. Baca file Excel master
        print(f"[*] Membaca file '{file_master}'...")
        df = pd.read_excel(file_master)
        
        # 2. Cari kolom yang bernama 'Kode Satker' dan 'Nama Satker'
        kolom_target_kode = 'Kode Satker'
        kolom_target_nama = 'Nama Satker'
        
        kolom_kode_asli = None
        kolom_nama_asli = None
        
        for col in df.columns:
            nama_kolom_kecil = str(col).lower().strip()
            if nama_kolom_kecil == kolom_target_kode.lower():
                kolom_kode_asli = col
            elif nama_kolom_kecil == kolom_target_nama.lower():
                kolom_nama_asli = col
                
        if not kolom_kode_asli or not kolom_nama_asli:
            print(f"[!] Error: Kolom 'Kode Satker' atau 'Nama Satker' tidak ditemukan dalam file Excel.")
            sys.exit(1)

        # 3. Ambil kedua kolom tersebut dan hapus yang kembar (duplicate)
        print("[*] Memfilter dan menghapus data yang kembar (duplikat)...")
        # Hapus baris yang kode satkernya kosong
        df_bersih = df.dropna(subset=[kolom_kode_asli]).copy()
        
        # Ambil kolom kode dan nama saja
        df_unik = df_bersih[[kolom_kode_asli, kolom_nama_asli]].drop_duplicates()
        
        # Bersihkan akhiran .0 jika terdeteksi angka desimal pada kode satker
        df_unik[kolom_kode_asli] = df_unik[kolom_kode_asli].astype(str).str.strip()
        df_unik[kolom_kode_asli] = df_unik[kolom_kode_asli].str.replace('.0', '', regex=False)
        
        # Hapus duplikat lagi (patokannya hanya Kode Satkernya saja)
        df_unik = df_unik.drop_duplicates(subset=[kolom_kode_asli])
        
        print(f"[*] Ditemukan {len(df_unik)} Kode Satker ")

        # 4. Simpan ke 1 file tunggal
        df_unik.to_excel(file_output, index=False)
        print(f"[*] Berhasil! Daftar Kode Satker telah difilter dan disimpan ke dalam 1 file: '{file_output}'")

    except FileNotFoundError:
        print(f"[!] Error: File '{file_master}' tidak ditemukan.")
        sys.exit(1)
    except Exception as e:
        print(f"[!] Terjadi kesalahan: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # Setup argparse untuk membaca argumen dari perintah terminal
    parser = argparse.ArgumentParser(description="Script untuk mengekstrak daftar Kode Satker unik dari Excel.")
    parser.add_argument("file_master", help="Nama file Excel mentahan")
    parser.add_argument("-o", "--output", default="Daftar_Kode_Satker.xlsx", help="Nama file hasil (opsional)")
    
    args = parser.parse_args()
    filter_kode_satker(args.file_master, args.output)
