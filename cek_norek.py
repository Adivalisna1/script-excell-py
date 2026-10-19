"""
Cek & Bandingkan Data Rekening antar 2 File Excel
=================================================
Membandingkan data rekening antara 2 file Excel dengan format tabel berbeda.
Kunci pencarian utama: NOMOR REKENING.
Setelah No Rekening ketemu di File 2, script membandingkan:
    - Kode Satker & Nama Satker
    - Nama Rekening
    - Jenis Rekening (BPG / BPN / BPP / RPL / RKK ...)

Keterangan Output:
    - Sesuai        -> No Rek cocok & No Satker, Nama Satker, Nama Rek, Jenis Rek sama
                       (Keterangan: "Sesuai (Satker: [Kode] - [Nama Satker])")
    - Beda          -> No Rek ketemu di File 2, tetapi ada data yang berbeda
                       (Keterangan merinci apa saja yang beda beserta nilai F1 vs F2)
    - Tidak Ada     -> No Rek tidak ditemukan di File 2
    - No Rek Kosong -> Baris judul satker / baris tanpa nomor rekening

Output: 1 file Excel (default: hasil_perbandingan.xlsx)

Cara pakai:
    python cek_norek.py
    python cek_norek.py file1.xlsx file2.xlsx
    python cek_norek.py file1.xlsx file2.xlsx --baris "422-1058"
"""

import argparse
import os
import re
import sys

import pandas as pd

MAX_SCAN_HEADER = 30


# --------------------------------------------------------------------------- #
# Normalisasi Nilai
# --------------------------------------------------------------------------- #
def kosong(val) -> bool:
    return val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() in ("", "nan", "None", "-")


def norm_rek(val) -> str:
    """Hanya angka/huruf, tanpa spasi/titik/strip, tanpa nol di depan."""
    if kosong(val):
        return ""
    if isinstance(val, float):
        val = f"{val:.0f}" if val.is_integer() else str(val)
    s = str(val).strip().lstrip("'")
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = re.sub(r"[^0-9A-Za-z]", "", s).lstrip("0")
    return s.upper()


def norm_nama(val) -> str:
    """Huruf besar, abaikan spasi & tanda baca."""
    return "" if kosong(val) else re.sub(r"[^0-9A-Z]", "", str(val).upper())


def norm_jenis(val) -> str:
    """Ambil huruf di awal kata: 'RPL020' -> 'RPL', 'Rpl' -> 'RPL'."""
    if kosong(val):
        return ""
    m = re.match(r"\s*([A-Za-z]+)", str(val))
    return m.group(1).upper() if m else ""


def norm_satker(val) -> str:
    if kosong(val):
        return ""
    s = str(val).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return re.sub(r"\D", "", s).lstrip("0")


# --------------------------------------------------------------------------- #
# Deteksi Kolom Otomatis
# --------------------------------------------------------------------------- #
def _low(c):
    return re.sub(r"\s+", " ", str(c)).strip().lower()


def is_kolom_rek(c):
    t = _low(c)
    if any(x in t for x in ("nama", "jenis", "pembukaan", "tanggal", "saldo")):
        return False
    return any(k in t for k in ("rek", "account", "acc no", "acct", "akun"))


def is_kolom_nama(c):
    t = _low(c)
    return "nama" in t and any(k in t for k in ("rek", "account", "akun")) and "satker" not in t


def is_kolom_jenis(c):
    t = _low(c)
    return any(k in t for k in ("jenis", "tipe", "type"))


def is_kolom_satker_kode(c):
    t = _low(c)
    return "satker" in t and any(k in t for k in ("kode", "no", "nomor")) and "nama" not in t


def is_kolom_satker_nama(c):
    t = _low(c)
    return "satker" in t and "nama" in t


def cari_baris_header(path, sheet, nama_kolom_rek=None) -> int:
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=MAX_SCAN_HEADER, dtype=str)
    for idx, row in raw.iterrows():
        for cell in row:
            if kosong(cell):
                continue
            if nama_kolom_rek:
                if _low(cell) == _low(nama_kolom_rek):
                    return idx
            elif is_kolom_rek(cell):
                return idx
    return 0


def cari_kolom(df, detektor, manual=None):
    if manual:
        match = [c for c in df.columns if _low(c) == _low(manual)]
        if match:
            return match[0]
    kandidat = [c for c in df.columns if detektor(c)]
    return kandidat[0] if kandidat else None


def baca_tabel(path, sheet, manual_rek=None):
    nama_file = os.path.basename(path)
    h = cari_baris_header(path, sheet, manual_rek)
    df = pd.read_excel(path, sheet_name=sheet, header=h, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    # Buang kolom tanpa nama yang isinya kosong total
    df = df[[c for c in df.columns if not (c.startswith("Unnamed") and df[c].isna().all())]]
    df = df.dropna(how="all")
    df.insert(0, "Baris Excel", df.index + h + 2)

    kol_rek = cari_kolom(df, is_kolom_rek, manual_rek)
    if not kol_rek:
        print(f"\nKolom No Rekening di '{nama_file}' tidak terdeteksi otomatis.")
        for i, c in enumerate(df.columns, 1):
            print(f"  {i}. {c}")
        kol_rek = df.columns[int(input("Pilih nomor kolom No Rekening: ")) - 1]

    kol_nama = cari_kolom(df, is_kolom_nama)
    kol_jenis = cari_kolom(df, is_kolom_jenis)
    kol_satker_kode = cari_kolom(df, is_kolom_satker_kode)
    kol_satker_nama = cari_kolom(df, is_kolom_satker_nama)

    df["_rek"] = df[kol_rek]
    df["_judul"] = df["_rek"].map(norm_rek) == ""

    # Nama rekening
    df["_nama"] = df[kol_nama] if kol_nama else None

    # Jenis rekening
    if kol_jenis:
        df["_jenis"] = df[kol_jenis]
    elif kol_nama:
        df["_jenis"] = df["_nama"].map(norm_jenis)
    else:
        df["_jenis"] = None

    # Kode & Nama Satker
    if kol_satker_kode and kol_satker_nama:
        df["_satker_kode"] = df[kol_satker_kode]
        df["_satker_nama"] = df[kol_satker_nama]
        info_satker = f"kolom '{kol_satker_kode}' & '{kol_satker_nama}'"
    elif kol_satker_kode:
        df["_satker_kode"] = df[kol_satker_kode]
        df["_satker_nama"] = None
        info_satker = f"kolom '{kol_satker_kode}'"
    else:
        # Ekstrak dari baris judul grup satker
        kodes, namas = [], []
        for _, row in df.iterrows():
            if row["_judul"]:
                ditemukan = False
                for v in row.drop(labels=["Baris Excel", "_rek", "_judul"], errors="ignore"):
                    if not kosong(v):
                        m = re.match(r"^\s*(\d{6,})\s*(.*)", str(v))
                        if m:
                            kodes.append(m.group(1).strip())
                            namas.append(m.group(2).strip())
                            ditemukan = True
                            break
                if not ditemukan:
                    kodes.append(None)
                    namas.append(None)
            else:
                kodes.append(None)
                namas.append(None)

        df["_satker_kode"] = pd.Series(kodes, index=df.index).ffill()
        df["_satker_nama"] = pd.Series(namas, index=df.index).ffill()
        info_satker = "diekstrak otomatis dari baris judul satker"

    info = {
        "header": h + 1,
        "rek": f"kolom '{kol_rek}'",
        "nama": f"kolom '{kol_nama}'" if kol_nama else "TIDAK ADA",
        "jenis": f"kolom '{kol_jenis}'" if kol_jenis else ("diambil dari awal Nama Rek" if kol_nama else "TIDAK ADA"),
        "satker": info_satker,
    }
    return df, info


# --------------------------------------------------------------------------- #
# Input Pilihan Baris & Path File
# --------------------------------------------------------------------------- #
def parse_baris(teks):
    """'3-30, 50, 100-120', '422-1.058', '422 sampai 1058' -> [(3,30), ...]"""
    rentang = []
    # Bersihkan titik pemisah ribuan pada angka, misal 1.058 -> 1058
    teks_clean = re.sub(r"(?<=\d)\.(?=\d)", "", teks)
    # Ganti 'sd', 'sampai', 'to' jadi '-'
    teks_clean = re.sub(r"\b(?:sd|sampai|to)\b", "-", teks_clean, flags=re.IGNORECASE)
    # Rapikan spasi di sekitar '-' (misal '422 - 1058' -> '422-1058')
    teks_clean = re.sub(r"\s*-\s*", "-", teks_clean)

    for bagian in re.split(r"[,;\s]+", teks_clean.strip()):
        if not bagian:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", bagian)
        if not m:
            raise ValueError(f"Format baris tidak dikenali: '{bagian}'. Contoh benar: 422-1058 atau 3-30, 50, 100-120")
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else a
        rentang.append((min(a, b), max(a, b)))
    return rentang


def input_rentang(df):
    awal, akhir = df["Baris Excel"].min(), df["Baris Excel"].max()
    print(f"\nData File 1 ada di baris Excel {awal} s/d {akhir}.")
    print("Contoh: 422-1058   atau   3-30, 50, 100-120   (kosongkan = semua)")
    while True:
        teks = input("Baris yang mau dicek: ").strip()
        if not teks:
            return []
        try:
            return parse_baris(teks)
        except ValueError as e:
            print(f"  [!] {e}")


def pilih_sheet(path, sheet_arg):
    if sheet_arg:
        return sheet_arg
    sheets = pd.ExcelFile(path).sheet_names
    if len(sheets) == 1:
        return sheets[0]
    print(f"\nFile '{os.path.basename(path)}' punya beberapa sheet:")
    for i, s in enumerate(sheets, 1):
        print(f"  {i}. {s}")
    inp = input("Pilih nomor sheet [1]: ").strip() or "1"
    return sheets[int(inp) - 1]


def input_path(prompt):
    while True:
        raw = input(prompt).strip().strip('"').strip("'")
        if not raw:
            continue
        if os.path.isfile(raw):
            return raw
        for ext in (".xlsx", ".xls"):
            if os.path.isfile(raw + ext):
                return raw + ext
        folder = os.path.dirname(raw) or "."
        target = os.path.basename(raw).lower()
        if os.path.isdir(folder):
            files = [f for f in os.listdir(folder) if f.lower().endswith((".xlsx", ".xls"))]
            exact = [f for f in files if f.lower() == target or os.path.splitext(f)[0].lower() == target]
            if exact:
                res = os.path.join(folder, exact[0]) if folder != "." else exact[0]
                print(f"  -> Menggunakan file: {exact[0]}")
                return res
            matched = [f for f in files if target in f.lower()]
            if len(matched) == 1:
                res = os.path.join(folder, matched[0]) if folder != "." else matched[0]
                print(f"  -> Menggunakan file: {matched[0]}")
                return res
            elif len(matched) > 1:
                print(f"  [!] Ada beberapa file mirip:")
                for idx, mf in enumerate(matched, 1):
                    print(f"      {idx}. {mf}")
                pilih = input("      Pilih nomor [1]: ").strip() or "1"
                try:
                    pilihan_f = matched[int(pilih) - 1]
                    return os.path.join(folder, pilihan_f) if folder != "." else pilihan_f
                except Exception:
                    pass
        print(f"  [!] File '{raw}' tidak ditemukan. Coba ketik lagi nama filenya.")


# --------------------------------------------------------------------------- #
# Perbandingan Data
# --------------------------------------------------------------------------- #
def bandingkan(df1, df2):
    # Buat indeks File 2 berdasarkan No Rekening yang sudah dinormalisasi
    idx2 = {}
    for _, r in df2.iterrows():
        k = norm_rek(r["_rek"])
        if k:
            idx2.setdefault(k, []).append(r)

    hasil = []
    for _, r1 in df1.iterrows():
        out = {}
        sk1 = str(r1["_satker_kode"] or "").strip()
        sn1 = str(r1["_satker_nama"] or "").strip()
        satker_f1_txt = f"{sk1} - {sn1}" if (sk1 and sn1) else (sk1 or sn1 or "-")
        k = norm_rek(r1["_rek"])

        # 1. KASUS NO REK KOSONG
        if not k:
            out["Cek No Rek"] = "Kosong"
            if sk1 and sn1:
                out["Keterangan"] = f"No Rek Kosong (Judul Satker: {satker_f1_txt})"
            elif sk1:
                out["Keterangan"] = f"No Rek Kosong (Judul Satker: {sk1})"
            else:
                out["Keterangan"] = "No Rek Kosong"
            hasil.append(out)
            continue

        # 2. KASUS NO REK TIDAK DITEMUKAN DI FILE 2
        kandidat = idx2.get(k)
        if not kandidat:
            out["Cek No Rek"] = "Tidak Ada"
            sat_info = f" (Satker: {satker_f1_txt})" if sk1 else ""
            out["Keterangan"] = f"Tidak Ada di File 2{sat_info}"
            hasil.append(out)
            continue

        # 3. KASUS NO REK KETEMU DI FILE 2 -> BANDINGKAN
        # Jika ada beberapa baris dengan No Rek sama di File 2, pilih yang paling cocok
        def hitung_skor(r2):
            skor = 0
            if norm_satker(sk1) == norm_satker(r2["_satker_kode"]):
                skor += 2
            if norm_nama(r1["_nama"]) == norm_nama(r2["_nama"]):
                skor += 2
            if norm_jenis(r1["_jenis"]) == norm_jenis(r2["_jenis"]):
                skor += 1
            return skor

        r2 = max(kandidat, key=hitung_skor)

        out["Cek No Rek"] = "Ada"
        out["Baris File 2"] = r2["Baris Excel"]
        out["Kode Satker (File 2)"] = r2["_satker_kode"]
        out["Nama Satker (File 2)"] = r2["_satker_nama"]
        out["Nama Rek (File 2)"] = r2["_nama"]
        out["Jenis Rek (File 2)"] = r2["_jenis"]

        diffs = []

        # Cek Kode Satker
        if sk1 and r2["_satker_kode"]:
            if norm_satker(sk1) != norm_satker(r2["_satker_kode"]):
                diffs.append(f"Kode Satker (F1: {sk1} vs F2: {r2['_satker_kode']})")
                out["Cek Kode Satker"] = "Beda"
            else:
                out["Cek Kode Satker"] = "Sesuai"
        else:
            out["Cek Kode Satker"] = "Kosong"

        # Cek Nama Rekening
        if r1["_nama"] and r2["_nama"]:
            if norm_nama(r1["_nama"]) != norm_nama(r2["_nama"]):
                diffs.append(f"Nama Rek (F1: {r1['_nama']} vs F2: {r2['_nama']})")
                out["Cek Nama Rek"] = "Beda"
            else:
                out["Cek Nama Rek"] = "Sesuai"
        else:
            out["Cek Nama Rek"] = "Kosong"

        # Cek Jenis Rekening
        if r1["_jenis"] and r2["_jenis"]:
            if norm_jenis(r1["_jenis"]) != norm_jenis(r2["_jenis"]):
                diffs.append(f"Jenis Rek (F1: {r1['_jenis']} vs F2: {r2['_jenis']})")
                out["Cek Jenis Rek"] = "Beda"
            else:
                out["Cek Jenis Rek"] = "Sesuai"
        else:
            out["Cek Jenis Rek"] = "Kosong"

        # Tentukan Keterangan Akhir
        if not diffs:
            # Sesuai: tampilkan no satker & nama satker
            out["Keterangan"] = f"Sesuai (Satker: {satker_f1_txt})"
        else:
            # Beda: rincikan apa saja yang beda
            out["Keterangan"] = "Beda: " + "; ".join(diffs)

        hasil.append(out)

    return pd.DataFrame(hasil, index=df1.index)


# --------------------------------------------------------------------------- #
# Simpan ke Excel dengan Format & Warna Rapi
# --------------------------------------------------------------------------- #
def simpan(df, output):
    from openpyxl.styles import Alignment, Font, PatternFill

    warna = {
        "hijau": PatternFill("solid", fgColor="C6EFCE"),
        "merah": PatternFill("solid", fgColor="FFC7CE"),
        "oranye": PatternFill("solid", fgColor="FFD8A8"),
        "kuning": PatternFill("solid", fgColor="FFEB9C"),
        "abu_judul": PatternFill("solid", fgColor="F2F2F2"),
        "header_f1": PatternFill("solid", fgColor="1F4E78"),
        "header_f2": PatternFill("solid", fgColor="2E75B6"),
        "header_ket": PatternFill("solid", fgColor="1E4620"),
    }

    def tentukan_warna_ket(v):
        if not v or not isinstance(v, str):
            return None
        if v.startswith("Sesuai"):
            return warna["hijau"]
        if v.startswith("Tidak Ada"):
            return warna["merah"]
        if v.startswith("Beda"):
            return warna["oranye"]
        if v.startswith("No Rek Kosong"):
            return warna["kuning"]
        return None

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Hasil", index=False)
        ws = writer.sheets["Hasil"]
        cols = list(df.columns)

        # Header styling
        for j, c in enumerate(cols, 1):
            cell = ws.cell(row=1, column=j)
            cell.font = Font(bold=True, color="FFFFFF")
            if c == "Keterangan":
                cell.fill = warna["header_ket"]
            elif "(File 2)" in c or c == "Baris File 2":
                cell.fill = warna["header_f2"]
            else:
                cell.fill = warna["header_f1"]
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ket_idx = cols.index("Keterangan") + 1
        status_indices = [j for j, c in enumerate(cols, 1) if c.startswith("Cek ")]

        for row in ws.iter_rows(min_row=2):
            ket_val = row[ket_idx - 1].value
            f_ket = tentukan_warna_ket(ket_val)
            if f_ket:
                row[ket_idx - 1].fill = f_ket

            # Jika baris judul satker, beri highlight lembut
            if isinstance(ket_val, str) and "Judul Satker" in ket_val:
                for c in row:
                    if c.column != ket_idx:
                        c.fill = warna["abu_judul"]
                    c.font = Font(bold=True)

            # Warna status per kolom
            for j in status_indices:
                val = row[j - 1].value
                if val in ("Ada", "Sesuai"):
                    row[j - 1].fill = warna["hijau"]
                elif val == "Tidak Ada":
                    row[j - 1].fill = warna["merah"]
                elif val == "Beda":
                    row[j - 1].fill = warna["oranye"]
                elif val in ("Kosong",):
                    row[j - 1].fill = warna["kuning"]

        # Lebar kolom otomatis
        for col_cells in ws.columns:
            lebar = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(lebar + 3, 11), 60)

        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions


# --------------------------------------------------------------------------- #
# Alur Utama & Pembandingan
# --------------------------------------------------------------------------- #
def banding_dan_simpan(df1_master, df2_master, file1, rentang=None, output=None):
    df1 = df1_master.copy()
    if rentang:
        mask = pd.Series(False, index=df1.index)
        for a, b in rentang:
            mask |= df1["Baris Excel"].between(a, b)
        df1 = df1[mask]
        teks_rentang = ", ".join(f"{a}-{b}" if a != b else str(a) for a, b in rentang)
    else:
        teks_rentang = "semua"

    if df1.empty:
        print(f"\n[!] Tidak ada data di baris {teks_rentang}.")
        return None

    print(f"\nMemproses baris File 1: {teks_rentang} ...")
    hasil = bandingkan(df1, df2_master)

    # Susun kolom output rapi (buang kolom bantu '_' dan kolom 'Unnamed')
    asli = [c for c in df1.columns if not c.startswith("_") and not c.startswith("Unnamed")]
    out = df1[asli].copy()

    # Tambahkan kolom Satker File 1 di dekat data rekening jika belum ada
    if "Kode Satker" not in out.columns:
        out.insert(out.columns.get_loc("No. Rekening") if "No. Rekening" in out.columns else 2,
                   "Kode Satker (File 1)", df1["_satker_kode"])
    if "Nama Satker" not in out.columns:
        out.insert(out.columns.get_loc("No. Rekening") if "No. Rekening" in out.columns else 3,
                   "Nama Satker (File 1)", df1["_satker_nama"])

    # Gabungkan kolom perbandingan dari File 2
    for c in hasil.columns:
        out[c] = hasil[c]

    if not output:
        output = os.path.join(os.path.dirname(os.path.abspath(file1)), "hasil_perbandingan.xlsx")

    try:
        simpan(out, output)
    except PermissionError:
        print(f"\n[!] PERINGATAN: File '{os.path.basename(output)}' sedang dibuka di Excel!")
        print("    Silakan tutup filenya di Excel terlebih dahulu, lalu jalankan lagi.")
        return None

    # Ringkasan di terminal
    print("\n=== RINGKASAN HASIL ===")
    print(f"Total baris dicek : {len(out)}")
    for k, v in out["Keterangan"].apply(lambda x: x.split(" (")[0]).value_counts().items():
        print(f"  {k:<20}: {v}")
    print(f"\nHasil perbandingan tersimpan di:\n  -> {output}")
    return output


def main():
    ap = argparse.ArgumentParser(description="Bandingkan data rekening antar 2 file Excel")
    ap.add_argument("file1", nargs="?", help="File Excel utama")
    ap.add_argument("file2", nargs="?", help="File Excel pembanding")
    ap.add_argument("--sheet1"), ap.add_argument("--sheet2")
    ap.add_argument("--rek1", help="Nama kolom no rekening file 1")
    ap.add_argument("--rek2", help="Nama kolom no rekening file 2")
    ap.add_argument("-o", "--output", help="Nama file output")
    ap.add_argument("--baris", help='Baris Excel file 1 yang dicek, contoh: "422-1058"')
    ap.add_argument("--semua", action="store_true", help="Cek semua baris tanpa ditanya")
    ap.add_argument("--no-loop", action="store_true", help="Jalankan sekali tanpa menu perulangan")
    a = ap.parse_args()

    file1 = a.file1
    file2 = a.file2
    df1_master = None
    df2_master = None

    while True:
        # Jika file belum dimuat atau user memilih ganti file
        if df1_master is None or df2_master is None:
            file1 = file1 or input_path("Path File 1 (utama)     : ")
            file2 = file2 or input_path("Path File 2 (pembanding): ")

            sheet1 = pilih_sheet(file1, a.sheet1)
            sheet2 = pilih_sheet(file2, a.sheet2)

            print("\n[+] Membaca file...")
            df1_master, info1 = baca_tabel(file1, sheet1, a.rek1)
            df2_master, info2 = baca_tabel(file2, sheet2, a.rek2)

            print(f"\nFile 1 (Utama)     : {os.path.basename(file1)} | sheet '{sheet1}'")
            print(f"   No Rekening     : {info1['rek']}")
            print(f"   Satker          : {info1['satker']}")
            print(f"\nFile 2 (Pembanding): {os.path.basename(file2)} | sheet '{sheet2}'")
            print(f"   No Rekening     : {info2['rek']}")
            print(f"   Satker          : {info2['satker']}")

        # Tentukan baris yang dicek
        if a.baris:
            rentang = parse_baris(a.baris)
            a.baris = None  # Hanya untuk run pertama
        elif a.semua:
            rentang = []
            a.semua = False
        else:
            rentang = input_rentang(df1_master)

        # Jalankan pembandingan dan simpan
        banding_dan_simpan(df1_master, df2_master, file1, rentang, a.output)

        if a.no_loop:
            break

        # Menu Interaktif Setelah Selesai
        print("\n" + "=" * 55)
        print("MENU SELANJUTNYA:")
        print("  1. Cek baris lain (tetap pakai file yang sama)")
        print("  2. Ganti file pembanding (pilih file baru)")
        print("  3. Selesai / Keluar (stop script)")
        print("=" * 55)

        pilih_lagi = False
        while True:
            opsi = input("Pilih menu [1/2/3] (default 1): ").strip().lower()
            if not opsi or opsi == "1":
                print(f"\n[+] Menggunakan file yang sama: {os.path.basename(file1)} & {os.path.basename(file2)}")
                pilih_lagi = True
                break
            elif opsi == "2":
                file1 = None
                file2 = None
                df1_master = None
                df2_master = None
                a.sheet1 = None
                a.sheet2 = None
                pilih_lagi = True
                break
            elif opsi in ("3", "stop", "exit", "keluar", "q", "n", "tidak"):
                print("\nTerima kasih! Script dihentikan.")
                return
            else:
                print("  [!] Pilihan tidak valid. Silakan ketik 1, 2, atau 3.")

        if not pilih_lagi:
            break


if __name__ == "__main__":
    main()
