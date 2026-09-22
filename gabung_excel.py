import pandas as pd
import os
import argparse
import sys

def gabung_semua_excel(folder_utama, file_output="Data_Gabungan.xlsx"):
    # 1. Pastikan folder utama ada
    if not os.path.exists(folder_utama):
        print(f"[!] Error: Folder '{folder_utama}' tidak ditemukan.")
        sys.exit(1)

    print(f"[*] Mencari file Excel di dalam folder '{folder_utama}' dan seluruh sub-foldernya...")
    
    semua_df = []
    jumlah_file = 0

    # 2. Sisir semua folder dan subfolder mencari file excel
    for root, dirs, files in os.walk(folder_utama):
        for file in files:
            # Cari file yang berakhiran .xlsx (dan hindari file temporary excel yang berawalan ~)
            if file.endswith('.xlsx') and not file.startswith('~'):
                file_path = os.path.join(root, file)
                print(f"  -> Membaca: {file_path}")
                try:
                    df = pd.read_excel(file_path)
                    semua_df.append(df)
                    jumlah_file += 1
                except Exception as e:
                    print(f"     [!] Gagal membaca {file}: {e}")

    # 3. Peringatan jika tidak ada file
    if jumlah_file == 0:
        print("[!] Tidak ada file Excel yang ditemukan untuk digabungkan.")
        return

    print(f"\n[*] Menggabungkan total {jumlah_file} file menjadi satu...")
    
    # 4. Menggabungkan semua data dari berbagai dataframe menjadi 1 dataframe utuh
    try:
        df_gabungan = pd.concat(semua_df, ignore_index=True)
    except Exception as e:
        print(f"[!] Gagal menggabungkan data: {e}")
        return

    print(f"[*] Menyimpan hasil gabungan ke file '{file_output}'...")
    
    # 5. Menyimpan dataframe gabungan ke dalam excel
    try:
        df_gabungan.to_excel(file_output, index=False)
        print(f"[*] Selesai! File gabungan berhasil disimpan sebagai '{file_output}'.")
    except Exception as e:
        print(f"[!] Gagal menyimpan file gabungan: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script untuk menggabungkan banyak file Excel di dalam folder dan subfolder.")
    
    # Parameter untuk mengatur folder input (defaultnya hasil_pilahan)
    parser.add_argument("-i", "--input", default="hasil_pilahan", 
                        help="Folder utama tempat berbagai file Excel berada (default: hasil_pilahan)")
    
    # Parameter untuk mengatur nama file output
    parser.add_argument("-o", "--output", default="Data_Gabungan.xlsx", 
                        help="Nama file hasil penggabungan (default: Data_Gabungan.xlsx)")
    
    args = parser.parse_args()
    
    gabung_semua_excel(args.input, args.output)
