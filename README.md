# Script Excel Python (Otomasi Laporan Satker)

Kumpulan script Python untuk melakukan otomasi pengolahan data Excel terkait Satuan Kerja (Satker). 

## 📋 Prasyarat / Requirements
Pastikan Anda sudah menginstal **Python** dan beberapa library (modul) yang dibutuhkan. Buka terminal (Command Prompt/PowerShell) dan jalankan perintah berikut untuk menginstal modul yang diperlukan:

```bash
pip install pandas openpyxl
```

## 🚀 Daftar Script dan Cara Penggunaan

### 1. Memilah Data Berdasarkan Satker (`pilah_satker.py`)
Script ini berfungsi untuk memecah/memilah satu file Excel mentahan yang besar menjadi banyak file Excel kecil-kecil berdasarkan kolom **Nama Satker**.

**Cara Pakai:**
```bash
python pilah_satker.py <nama_file_excel_mentah>
```
*Contoh:*
```bash
python pilah_satker.py data_master.xlsx
```
**Opsi Tambahan:**
Bisa menambahkan `-o` untuk menentukan nama folder output.
```bash
python pilah_satker.py data_master.xlsx -o folder_hasil_baru
```
*(Secara bawaan/default, hasil akan disimpan di dalam folder `hasil_pilahan`)*

---

### 2. Memformat Data Sesuai Template (`format_template.py`)
Script ini berfungsi untuk menyesuaikan format/struktur kolom file Excel agar sama persis dengan format template yang dibutuhkan (Tanggal, Uraian, Kategori, No Rekening, dsb).

**Cara Pakai:**
```bash
python format_template.py <nama_file_excel>
```
*Contoh:*
```bash
python format_template.py data_mentah.xlsx
```
**Opsi Tambahan:**
```bash
python format_template.py data_mentah.xlsx -o hasil_cetakan.xlsx
```
*(Secara bawaan/default, hasil akan disimpan otomatis di dalam folder `Hasil_Template` dengan nama file `[nama_file_asli]_Template.xlsx`)*

---

### 3. Menggabungkan Banyak File Excel (`gabung_excel.py`)
Script ini berfungsi untuk menggabungkan banyak file Excel yang ada di dalam sebuah folder (dan sub-foldernya) menjadi satu file utuh.

**Cara Pakai:**
```bash
python gabung_excel.py
```
**Opsi Tambahan:**
Bisa menambahkan `-i` untuk folder sumber (input) dan `-o` untuk nama file gabungannya (output).
```bash
python gabung_excel.py -i folder_data_pilahan -o Rekap_Semua_Data.xlsx
```
*(Secara bawaan/default, script akan mengambil file dari folder `hasil_pilahan` dan menyimpannya sebagai `Data_Gabungan.xlsx`)*

---

### 4. Mengambil Daftar Kode Satker Unik (`filter_kode_satker.py`)
Script ini berfungsi untuk mengekstrak dan memfilter daftar **Kode Satker** dan **Nama Satker** agar tidak ada data yang kembar/duplikat (unik).

**Cara Pakai:**
```bash
python filter_kode_satker.py <nama_file_excel_mentah>
```
*Contoh:*
```bash
python filter_kode_satker.py data_master.xlsx
```
**Opsi Tambahan:**
```bash
python filter_kode_satker.py data_master.xlsx -o Daftar_Unik.xlsx
```
*(Secara bawaan/default, hasil akan disimpan sebagai `Daftar_Kode_Satker.xlsx`)*
