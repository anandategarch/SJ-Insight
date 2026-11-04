"""sj_anomaly — replika M code "04.QUERRY SJ.xlsb" untuk analisa pemindahan
barang (Surat Jalan / Stock Transfer) DC ↔ RESTO.

Paket ini mengikuti arsitektur ``waste_anomaly`` (repo Waste):
  - ``loader``        : pemuat 3 sumber Excel (42.15 Kirim / 42.05 Terima / Kategori)
  - ``schema``        : validasi kontrak header
  - ``stages``        : stage-by-stage replika Power Query M (29 query, 1.034 baris)
  - ``pipeline``      : orchestrator (PIPELINE_STAGES + _gate cancel + progress_callback)
  - ``exporter``      : multi-sheet Excel writer (WriteOnlyCell streaming)
  - ``signals``       : ringkasan KPI per gudang / barang / breakdown anomali
  - ``utils``         : load_config, clamp, norm_text, format angka id-ID
  - ``logging_setup`` : logger file + console
  - ``followup``      : tindak lanjut baris anomali (centang + persist JSON)
"""
from __future__ import annotations

__version__ = "1.0.0"
__all__ = ["__version__"]
