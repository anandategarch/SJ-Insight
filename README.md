# SJ Insight — Analisa Pemindahan Barang DC ↔ RESTO

Program desktop (PySide6) untuk **analisa anomali pemindahan barang** antara
gudang DC dan RESTO. Replika **1:1** dari logika Power Query M code file
`04.QUERRY SJ.xlsb` (29 query, 1.034 baris M code) — di-port ke Python
mengikuti arsitektur repo [Waste](../Waste/).

## Apa yang dianalisa?

Program mendeteksi **4 kategori anomali** pada dokumen pemindahan barang:

| Kategori | Sumber M code | Deteksi |
|---|---|---|
| **Anomali Satuan** | query `3 42 15` step A | Satuan input ≠ mode satuan standar per Nama Barang |
| **Anomali QTY** | query `3 42 15` step C | 5 rule: nol/minus · desimal pada satuan utuh · lonjakan ekstrem (>3.5× rata-rata) · terlalu kecil (<0.1×) · status non-diterima |
| **Anomali Konversi** | query `3 42 15` step L-M | `QTY Terkecil 2 = QTY Terkecil / Kuantitas` menyimpang >10% dari standar mode (per Nama Barang + Satuan) |
| **Rekonsiliasi Kirim↔Terima** | query `Merge4205&4215` | Left join 42.15 Kirim ← 42.05 Terima → deteksi dokumen hilang + `Selisih Hari` ≥ 2 |

## Struktur Proyek

```
SJ/
├── sj_anomaly/              # Pipeline + stages + loader + exporter (mirror waste_anomaly/)
│   ├── __init__.py
│   ├── config.json          # Threshold (editable, berlaku run berikutnya)
│   ├── frozen.py            # Deteksi mode frozen (PyInstaller)
│   ├── logging_setup.py     # Logger file rotasi + console
│   ├── loader.py            # Pemuat 3 sumber Excel (calamine → openpyxl fallback)
│   ├── schema.py            # Validasi kontrak header (anti degradasi senyap)
│   ├── stages.py            # ⭐ 18 stage replika M code (1.034 baris)
│   ├── pipeline.py          # Orchestrator (PIPELINE_STAGES + _gate cancel + progress)
│   ├── signals.py           # Ringkasan KPI + chart data
│   ├── exporter.py          # Multi-sheet Excel writer (WriteOnlyCell streaming)
│   ├── followup.py          # Tindak lanjut baris (centang + persist JSON)
│   └── utils.py             # load_config, clamp, fmt id-ID, severity
├── sj_ui/                   # PySide6 UI (mirror ui/)
│   ├── __init__.py
│   ├── shared.py            # Konstanta + fmt_id
│   ├── theme.py             # Palet warna + QSS global
│   ├── state.py             # AppContext (state tunggal + signal Qt)
│   ├── models.py            # PandasTableModel (QAbstractTableModel virtualized)
│   ├── workers.py           # QThread: PipelineWorker + ExportWorker
│   ├── widgets.py           # KpiCard, DataTable, WLabel
│   ├── view_pipeline.py     # Tab Pipeline (input + run + log)
│   ├── view_dashboard.py    # Tab Dashboard (8 KPI + top gudang/barang)
│   ├── view_anomali.py      # Tab Anomali (tabel baris mencurigakan + filter)
│   ├── view_merge.py        # Tab Kirim vs Terima (rekonsiliasi dokumen)
│   └── app.py               # MainWindow + main()
├── ui_app.py                # ⭐ Entry point (dipakai MULAI.bat)
├── make_sample.py           # ⭐ Generate 3 sample workbook (data demo)
├── requirements.txt         # Dependensi pinned
├── SETUP.bat                # Instalasi sekali
├── MULAI.bat                # Jalankan aplikasi (auto-generate sample)
├── data/                    # File input Excel (diisi oleh make_sample.py)
├── output/                  # Hasil ekspor Excel
└── README.md                # Dokumen ini
```

## Cara Pakai

### 1. Setup (sekali)

```bat
SETUP.bat
```

Atau manual:
```bash
pip install -r requirements.txt
```

### 2. Jalankan

```bat
MULAI.bat
```

Atau manual:
```bash
python make_sample.py    # generate 3 sample workbook ke folder data/
python ui_app.py
```

### 3. Alur di UI

1. **Tab Pipeline** → pilih 3 file Excel:
   - **42.15 Pemindahan Barang (Kirim)** — sheet `#42.15. Pemindahan Barang (KIRI`
   - **42.05 Rincian Pemindahan (Terima)** — sheet `#42.05. Rincian Pemindahan Bara`
   - **DBS Kategori** — sheet `Sheet1` (NAMA BARANG / KATEGORI / JENIS)
2. Klik **▶ Jalankan Analisa** → tunggu pipeline (13 stage, ~5 detik utk 720 baris demo)
3. **Tab Dashboard** → 8 KPI cards + tabel top gudang/barang
4. **Tab Anomali** → tabel baris dengan status ≠ NORMAL (filter gudang/barang/status)
5. **Tab Kirim vs Terima** → rekonsiliasi dokumen (highlight dokumen hilang + terima lambat)
6. Toolbar **💾 Ekspor Excel** → laporan 6 sheet (Detail · Anomali Prioritas · Ringkasan · Rekap Count · Log Run)

## Konfigurasi Threshold

Edit `sj_anomaly/config.json` — perubahan berlaku mulai run berikutnya:

```json
"thresholds": {
    "QTY_EXTREME_FACTOR":   3.5,     // Rule C.4: lonjakan ekstrem > 3.5× rata2
    "QTY_TINY_FACTOR":      0.1,     // Rule C.5: terlalu kecil < 0.1× rata2
    "MIN_TRANSAKSI":        3,       // Min transaksi utk standar rasio
    "DEVIASI_ANOMALI":      0.10,    // Rule L: |deviasi| > 10% = anomali konversi
    "DEVIASI_PERLU_CEK":    0.05,    // Rule L: |deviasi| > 5% = perlu cek
    "DOMINASI_MIN":         0.50,    // Rule L: dominasi < 50% = konversi tidak stabil
    "TERIMA_LAMBAT_HARI":   2        // Selisih hari utk flag "terima lambat"
}
```

## Sumber Spesifikasi

Program ini adalah **replika M code** dari file `04.QUERRY SJ.xlsb` (Power Query
v2.99.202.0). Mapping header asli dipertahankan agar hasil 1:1 dengan Excel:

- **Query `3 42 15`** → pipeline Kirim (13 stage: header → filter → tipe → KTG/Concet/ekstrak gudang → anomali satuan → anomali QTY → rasio konversi → status konversi)
- **Query `4 42 05 (2)`** → sheet Terima (42.05)
- **Query `21 DBS Kategori`** → master kategori
- **Query `Merge4205&4215`** → left join Kirim ← Terima + kategori (stage `merge_all`)
- **Query `4215TerimaCount` / `4215KirimCount`** → rekap duplikat Concet (stage `count_groups`)

## Dependencies

| Package | Versi | Fungsi |
|---|---|---|
| PySide6 | >=6.6 | UI framework (mirror Waste v2.6.0) |
| pandas | >=2.2,<3 | Pipeline data processing |
| numpy | >=1.26 | Numeric operations |
| openpyxl | >=3.1 | Exporter WriteOnlyCell + list_sheets |
| python-calamine | >=0.2 | Loader cepat (Rust, 7-10× openpyxl) — OPSIONAL, fallback otomatis |

## Lisensi

Internal use — replika logika internal Excel untuk tim operasional.
