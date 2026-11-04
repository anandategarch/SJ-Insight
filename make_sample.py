"""
make_sample.py — generate 3 sample workbook utk testing pipeline SJ Insight.

Membuat 3 file di folder ``data/``:
  - 42.15_Pemindahan_Barang_KIRIM.xlsx  (sheet '#42.15. Pemindahan Barang (KIRI')
  - 42.05_Rincian_Pemindahan_TERIMA.xlsx (sheet '#42.05. Rincian Pemindahan Bara')
  - DBS_KATEGORI.xlsx                   (sheet 'Sheet1')

Format mirror report ACCURATE: 3 baris judul di atas header (42.15 & 42.05),
header di baris pertama (DBS Kategori).

Anomali yang di-inject (agar semua detektor menyala):
  1. Anomali satuan        — item pakai satuan ≠ standar
  2. QTY nol/minus
  3. Desimal pada satuan utuh (KARTON)
  4. Lonjakan ekstrem       — QTY > 3.5× rata-rata peer
  5. Terlalu kecil          — QTY < 0.1× rata-rata
  6. Anomali konversi       — QTY Terkecil 2 menyimpang > 10%
  7. Konversi tidak stabil  — rasio bervariasi, tak ada nilai dominan
  8. Kirim tanpa dokumen terima (42.05)
  9. Terima lambat          — selisih hari ≥ 2

Run: python make_sample.py
"""
from __future__ import annotations

import random
from pathlib import Path

from openpyxl import Workbook

OUT_DIR = Path(__file__).resolve().parent / "data"

# ============================================================
# MASTER DATA DEMO
# ============================================================

class DemoItem:
    __slots__ = ("nama", "kode", "satuan", "satuan_terkecil", "rasio",
                 "qty_min", "qty_max", "desimal", "kategori", "jenis")

    def __init__(self, nama, kode, satuan, satuan_terkecil, rasio,
                 qty_min, qty_max, kategori, jenis, desimal=False):
        self.nama = nama
        self.kode = kode
        self.satuan = satuan
        self.satuan_terkecil = satuan_terkecil
        self.rasio = rasio
        self.qty_min = qty_min
        self.qty_max = qty_max
        self.desimal = desimal
        self.kategori = kategori
        self.jenis = jenis


ITEMS: list[DemoItem] = [
    DemoItem("BERAS CV RM 25KG",      1001, "KARUNG", "KG",  25, 4, 20,  "BAHAN BAKU",  "BOH"),
    DemoItem("GULA PASIR G1 50KG",    1002, "KARUNG", "KG",  50, 3, 15,  "BAHAN BAKU",  "BOH"),
    DemoItem("TEPUNG TERIGU CP 25KG", 1003, "KARUNG", "KG",  25, 5, 25,  "BAHAN BAKU",  "BOH"),
    DemoItem("MINYAK GORENG 1L",      1004, "KARTON", "PCS", 12, 3, 20,  "BAHAN BAKU",  "BOH"),
    DemoItem("MIE INSTAN GORENG 85G", 1005, "KARTON", "PCS", 40, 2, 15,  "BAHAN BAKU",  "BOH"),
    DemoItem("SUSU UHT 1L FULL CREAM",1006, "KARTON", "PCS", 24, 2, 12,  "BAHAN BAKU",  "BOH"),
    DemoItem("TELUR AYAM SEGAR",      1007, "KARTON", "TRAY",12, 1, 8,   "BAHAN BAKU",  "BOH"),
    DemoItem("AYAM BROILER POTONG 8", 1008, "KG",     "KG",   1, 20, 80, "BAHAN SEGAR", "BOH", desimal=True),
    DemoItem("DAGING SAPI SEGAR",      1009, "KG",     "KG",   1, 8, 40,  "BAHAN SEGAR", "BOH", desimal=True),
    DemoItem("IKAN NILA SEGAR",        1010, "KG",     "KG",   1, 8, 35,  "BAHAN SEGAR", "BOH", desimal=True),
    DemoItem("CABAI MERAH KERITING",   1012, "KG",     "KG",   1, 3, 18,  "BAHAN SEGAR", "BOH", desimal=True),
    DemoItem("BAWANG MERAH",           1013, "KG",     "KG",   1, 5, 25,  "BAHAN SEGAR", "BOH", desimal=True),
    DemoItem("ES KRISTAL 25KG",        1015, "KARUNG", "KG",  25, 4, 30,  "PENDUKUNG",   "BOH"),
    DemoItem("GAS LPG 3KG",            1016, "TABUNG", "PCS",  1, 4, 20,  "PENDUKUNG",   "BOH"),
    DemoItem("SAOS SAMBAL 320ML",      1017, "KARTON", "PCS", 24, 2, 12,  "BUMBU & SAOS","BOH"),
    DemoItem("AIR MINERAL 600ML",      1020, "KARTON", "PCS", 24, 4, 24,  "MINUMAN",     "FOH"),
    DemoItem("KOPI ROBUSTA 1KG",       1021, "PACK",   "PCS", 10, 1, 6,   "MINUMAN",     "FOH"),
    DemoItem("GELAS PLASTIK 240ML",    1023, "PACK",   "PCS", 50, 2, 10,  "KEMASAN",     "FOH"),
    DemoItem("SARUNG TANGAN HITAM",    1025, "PACK",   "PCS", 12, 1, 8,   "OPERASIONAL", "EQP"),
    DemoItem("PERMEN KOPICO 4GR",      1027, "KARTON", "PCS",200, 1, 4,   "MINUMAN",     "FOH"),
]

DC_PENGIRIM = [
    ("2001", "DC SMGMUR"),
    ("5001", "Depo"),  # → replace Depo = DCSMGMUR
]

RESTO = [
    "1001 RESTO KENCANA", "1004 RESTO MERDEKA", "1009 RESTO CAKRAWALA",
    "1012 RESTO ANGGREK", "1015 RESTO BAHAGIA", "1018 RESTO CENDANA",
    "1021 RESTO DAMAI",   "1024 RESTO ELITE",   "1027 RESTO FANTASI",
    "1030 RESTO GEMILANG", "1033 RESTO HARMONI", "1036 RESTO INDAH",
]


# ============================================================
# GENERATOR
# ============================================================

def _fmt_id_date(y, m, d, h=None, mi=None):
    s = f"{d:02d}/{m:02d}/{y}"
    if h is not None and mi is not None:
        s += f" {h:02d}:{mi:02d}"
    return s


def _tujuan_raw():
    """Gudang Tujuan raw '1009 (RESTO CAKRAWALA)'."""
    r = random.choice(RESTO)
    code = r.split(" ")[0]
    name = " ".join(r.split(" ")[1:])
    return f"{code} ({name})"


def _pengirim_raw():
    code, name = random.choice(DC_PENGIRIM)
    return f"{code} ({name})"


def _pad(n, w=4):
    return str(n).zfill(w)


def build_kirim_txns(rng, n_base=720):
    """Bangun list transaksi kirim (mirror demo.ts Waste Next)."""
    txns = []
    seq = 1

    def status_normal():
        r = rng.random()
        if r < 0.72: return "Diterima Seluruhnya"
        if r < 0.80: return "Diterima Seluruh nya"
        if r < 0.90: return "Diterima Sebagian"
        if r < 0.96: return "Dikirim"
        return "Belum Diterima"

    YEAR = 2025
    MONTH = 11
    DAYS = 30

    # Transaksi normal (baseline)
    for _ in range(n_base):
        item = rng.choice(ITEMS)
        day = rng.randint(1, DAYS)
        hour = rng.randint(6, 20)
        minute = rng.randint(0, 59)
        qty = (
            round(rng.uniform(item.qty_min, item.qty_max), 1)
            if item.desimal
            else rng.randint(item.qty_min, item.qty_max)
        )
        qty_terkecil = round(qty * item.rasio, 2)
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, max(1, day - rng.randint(0, 1)), hour, minute),
            "pengirim_raw": _pengirim_raw(),
            "tujuan_raw": _tujuan_raw(),
            "kode": item.kode,
            "nama": item.nama,
            "satuan": item.satuan,
            "qty_terkecil": qty_terkecil,
            "kuantitas": qty,
            "status": status_normal(),
        })
        seq += 1

    # Anomali 1: satuan ≠ standar (AYAM normalnya KG → beberapa PCS)
    for _ in range(9):
        day = rng.randint(1, DAYS)
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, max(1, day - 1), 8, rng.randint(0, 59)),
            "pengirim_raw": _pengirim_raw(),
            "tujuan_raw": _tujuan_raw(),
            "kode": 1008,
            "nama": "AYAM BROILER POTONG 8",
            "satuan": "PCS",  # standar KG
            "qty_terkecil": rng.randint(20, 60),
            "kuantitas": rng.randint(20, 60),
            "status": "Diterima Seluruhnya",
        })
        seq += 1

    # Anomali 2: QTY nol / minus
    for i in range(4):
        item = rng.choice(ITEMS)
        day = rng.randint(1, DAYS)
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, day, 9, 15),
            "pengirim_raw": _pengirim_raw(),
            "tujuan_raw": _tujuan_raw(),
            "kode": item.kode,
            "nama": item.nama,
            "satuan": item.satuan,
            "qty_terkecil": 0 if i < 2 else -item.rasio * 2,
            "kuantitas": 0,
            "status": "Diterima Seluruhnya",
        })
        seq += 1

    # Anomali 3: desimal pada satuan utuh (KARTON)
    for _ in range(6):
        item = rng.choice([x for x in ITEMS if not x.desimal])
        day = rng.randint(1, DAYS)
        qty = round(rng.uniform(2, 8), 5)  # 2.35 dst
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, day, 10, 30),
            "pengirim_raw": _pengirim_raw(),
            "tujuan_raw": _tujuan_raw(),
            "kode": item.kode,
            "nama": item.nama,
            "satuan": item.satuan,
            "qty_terkecil": round(qty * item.rasio, 3),
            "kuantitas": qty,
            "status": "Diterima Seluruhnya",
        })
        seq += 1

    # Anomali 4 & 5: lonjakan ekstrem / terlalu kecil (per item+tujuan)
    # GULA ke RESTO 1009: baseline ±8 karung → injeksikan 60 (lonjakan) & 0.4 (kecil)
    spike_target = RESTO[2]  # 1009 RESTO CAKRAWALA
    spike_tujuan = f"{spike_target.split(' ')[0]} ({' '.join(spike_target.split(' ')[1:])})"
    for qty in [60, 0.4, 0.3]:
        day = rng.randint(1, DAYS)
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, day, 11, 0),
            "pengirim_raw": "2001 (DC SMGMUR)",
            "tujuan_raw": spike_tujuan,
            "kode": 1002,
            "nama": "GULA PASIR G1 50KG",
            "satuan": "KARUNG",
            "qty_terkecil": round(qty * 50, 2),
            "kuantitas": qty,
            "status": "Diterima Seluruhnya",
        })
        seq += 1

    # Anomali 6: konversi menyimpang (BERAS std 25 → 30)
    for _ in range(4):
        day = rng.randint(1, DAYS)
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, day, 13, 20),
            "pengirim_raw": _pengirim_raw(),
            "tujuan_raw": _tujuan_raw(),
            "kode": 1001,
            "nama": "BERAS CV RM 25KG",
            "satuan": "KARUNG",
            "qty_terkecil": 30,  # seharusnya 25/karung → deviasi +20%
            "kuantitas": 1,
            "status": "Diterima Seluruhnya",
        })
        seq += 1

    # Anomali 7: konversi tidak stabil (MIE rasio acak 30..50)
    for _ in range(10):
        day = rng.randint(1, DAYS)
        qty = rng.randint(1, 4)
        rasio = rng.choice([30, 34, 38, 40, 44, 48, 50])
        txns.append({
            "no": f"IT/2511/{_pad(seq)}",
            "tanggal": _fmt_id_date(YEAR, MONTH, day),
            "tgl_create": _fmt_id_date(YEAR, MONTH, day, 15, 10),
            "pengirim_raw": "5001 (Depo)",
            "tujuan_raw": _tujuan_raw(),
            "kode": 1005,
            "nama": "MIE INSTAN GORENG 85G",
            "satuan": "KARTON",
            "qty_terkecil": qty * rasio,
            "kuantitas": qty,
            "status": "Diterima Seluruhnya",
        })
        seq += 1

    txns.sort(key=lambda t: t["no"])
    return txns


def write_kirim_workbook(txns, path):
    """Tulis workbook Kirim dengan 3 baris judul ACCURATE di atas header."""
    wb = Workbook()
    ws = wb.active
    ws.title = "#42.15. Pemindahan Barang (KIRI"

    # 3 baris judul ACCURATE
    ws.append(["ACCURATE Accounting System Report"])
    ws.append(["Halaman 1"])
    ws.append([f"Tercetak pada 7 November 2025 - 09:09"])

    # Header
    headers = [
        "No. Pemindahan #", "Tanggal", "Tgl/Jam Pembuatan", "Gudang Pengirim",
        "Gudang Tujuan", "Kode Barang", "Nama Barang", "Satuan",
        "Kts Diterima", "Kuantitas", "Status Pemindahan Barang",
    ]
    ws.append(headers)

    # Data
    for t in txns:
        ws.append([
            t["no"], t["tanggal"], t["tgl_create"], t["pengirim_raw"],
            t["tujuan_raw"], t["kode"], t["nama"], t["satuan"],
            t["qty_terkecil"], t["kuantitas"], t["status"],
        ])

    # Auto-width
    for c_idx, col in enumerate(ws.iter_cols(min_row=4, max_row=4), start=1):
        max_len = len(str(col[0].value or ""))
        for row in ws.iter_rows(min_row=5, max_row=min(50, ws.max_row), min_col=c_idx, max_col=c_idx):
            for cell in row:
                max_len = max(max_len, len(str(cell.value or "")))
        ws.column_dimensions[chr(64 + c_idx) if c_idx <= 26 else "A"].width = min(max_len + 2, 40)

    wb.save(path)


def build_terima_txns(rng, kirim_txns):
    """Bangun list terima dari kirim (subset dokumen yang diterima + polos)."""
    txns = []
    diterima_statuses = {"Diterima Seluruhnya", "Diterima Seluruh nya", "Diterima Sebagian"}
    tr_seq = 1
    YEAR = 2025
    MONTH = 11

    for t in kirim_txns:
        if t["status"] not in diterima_statuses:
            continue
        # Anomali 8: ±2.5% dokumen terima hilang
        if rng.random() < 0.025:
            continue

        # Anomali 9: ±8% terima lambat (2-4 hari)
        lambat = rng.randint(2, 4) if rng.random() < 0.08 else (1 if rng.random() < 0.4 else 0)
        dd, mm = t["tanggal"].split("/")[0], t["tanggal"].split("/")[1]
        tgl_terima_d = int(dd) + lambat
        tgl_terima_str = _fmt_id_date(YEAR, MONTH, tgl_terima_d)

        item = next((x for x in ITEMS if x.kode == t["kode"]), None)
        pengirim_code = t["pengirim_raw"].split(" ")[0]
        tujuan_code = t["tujuan_raw"].split(" ")[0]

        txns.append({
            "4205.xlsx": "4205.xlsx",
            "Nomor #Kirim": t["no"],
            "Nomor #Terima": f"TR/2511/{_pad(tr_seq)}",
            "Gudang #Kirim": pengirim_code,
            "Gudang #Terima": tujuan_code,
            "Status Pengiriman #": "Terima Terlambat" if lambat >= 2 else "Selesai",
            "Tanggal #Kirim": t["tanggal"],
            "Tanggal #Terima": tgl_terima_str,
            "#Tgl Kirim vs Tgl Terima": f"{lambat} Hari",
            "Kode Barang": t["kode"],
            "Nama Barang": t["nama"],
            "#Sat. Kirim": t["satuan"],
            "#Qty Kirim": t["kuantitas"],
            "Rasio Satuan": item.rasio if item else 1,
            "#Sat. Terkecil": item.satuan_terkecil if item else "PCS",
            "#Qty. Terkecil": t["qty_terkecil"],
            "#Tgl/Jam Pembuatan IT": t["tgl_create"],
            "Memiliki dokumen (Terima)": "YA",
        })
        tr_seq += 1

    return txns


def write_terima_workbook(txns, path):
    """Tulis workbook Terima dengan 3 baris judul + header + data."""
    wb = Workbook()
    ws = wb.active
    ws.title = "#42.05. Rincian Pemindahan Bara"

    # 3 baris judul
    ws.append(["ACCURATE Accounting System Report"])
    ws.append(["Halaman 1"])
    ws.append([f"Tercetak pada 7 November 2025 - 09:09"])

    headers = [
        "4205.xlsx", "Nomor #Kirim", "Nomor #Terima", "Gudang #Kirim", "Gudang #Terima",
        "Status Pengiriman #", "Tanggal #Kirim", "Tanggal #Terima", "#Tgl Kirim vs Tgl Terima",
        "Kode Barang", "Nama Barang", "#Sat. Kirim", "#Qty Kirim", "Rasio Satuan",
        "#Sat. Terkecil", "#Qty. Terkecil", "#Tgl/Jam Pembuatan IT", "Memiliki dokumen (Terima)",
    ]
    ws.append(headers)
    for t in txns:
        ws.append([
            t["4205.xlsx"], t["Nomor #Kirim"], t["Nomor #Terima"], t["Gudang #Kirim"],
            t["Gudang #Terima"], t["Status Pengiriman #"], t["Tanggal #Kirim"],
            t["Tanggal #Terima"], t["#Tgl Kirim vs Tgl Terima"], t["Kode Barang"],
            t["Nama Barang"], t["#Sat. Kirim"], t["#Qty Kirim"], t["Rasio Satuan"],
            t["#Sat. Terkecil"], t["#Qty. Terkecil"], t["#Tgl/Jam Pembuatan IT"],
            t["Memiliki dokumen (Terima)"],
        ])
    wb.save(path)


def write_kategori_workbook(path):
    """Tulis workbook DBS Kategori (sheet Sheet1, header di baris 1)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws.append(["DBS_KATEGORI.xlsx", "NAMA BARANG", "KATEGORI", "JENIS"])
    for item in ITEMS:
        ws.append(["DBS_KATEGORI.xlsx", item.nama, item.kategori, item.jenis])
    wb.save(path)


def main():
    rng = random.Random(20251104)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    kirim_txns = build_kirim_txns(rng, n_base=720)
    terima_txns = build_terima_txns(rng, kirim_txns)

    kirim_path = OUT_DIR / "42.15_Pemindahan_Barang_KIRIM.xlsx"
    terima_path = OUT_DIR / "42.05_Rincian_Pemindahan_TERIMA.xlsx"
    kategori_path = OUT_DIR / "DBS_KATEGORI.xlsx"

    write_kirim_workbook(kirim_txns, kirim_path)
    write_terima_workbook(terima_txns, terima_path)
    write_kategori_workbook(kategori_path)

    print(f"Sample workbooks ditulis ke: {OUT_DIR}")
    print(f"  Kirim    : {kirim_path.name}  ({len(kirim_txns)} baris)")
    print(f"  Terima   : {terima_path.name}  ({len(terima_txns)} baris)")
    print(f"  Kategori : {kategori_path.name}  ({len(ITEMS)} baris master)")
    print()
    print("Untuk menjalankan pipeline:")
    print(f"  python ui_app.py")
    print("  → tab Pipeline → isi ketiga path di atas → ▶ Jalankan Analisa")


if __name__ == "__main__":
    main()
