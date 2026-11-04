"""pipeline.py — orchestrator replika M code '04.QUERRY SJ.xlsb'.

Mengikuti pola ``waste_anomaly.pipeline``:
  - PIPELINE_STAGES (satu sumber kebenaran jumlah & urutan stage)
  - PipelineCancelledError + _gate cancel hook antar-stage (F9)
  - progress_callback(msg, pct) — pct SELALU numeric, monotonic
  - Hasil in-memory (dict summary), TIDAK menulis file Excel dari pipeline
  - File Excel ditulis terpisah oleh ``sj_anomaly.exporter`` bila diminta

Penggunaan dasar:
    from sj_anomaly.pipeline import run_pipeline

    summary = run_pipeline(
        kirim_path="data/42.15_KIRIM.xlsx",
        terima_path="data/42.05_TERIMA.xlsx",
        kategori_path="data/DBS_KATEGORI.xlsx",
        progress_callback=lambda msg, pct: print(f"[{pct:.0f}%] {msg}"),
    )
    df = summary["df"]  # hasil in-memory
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

from . import stages
from . import signals as sig
from .loader import load_kategori, load_kirim, load_terima
from .schema import (
    SchemaValidationError,
    validate_kategori_schema,
    validate_kirim_schema,
    validate_terima_schema,
)
from .utils import load_config


# ============================================================
# SATU SUMBER KEBENARAN jumlah & urutan stage (mirror Waste)
# ============================================================
PIPELINE_STAGES: list[str] = [
    "Stage 0 — Load 3 sumber (42.15 / 42.05 / Kategori)",
    "Stage 1-3 — Header, filter baris valid, tipe data",
    "Stage 4-8 — KTG, Concet, ekstrak gudang, selisih tgl",
    "Stage A — Anomali satuan (mode per barang)",
    "Stage B — Statistik QTY per barang+satuan+tujuan",
    "Stage C — Status anomali QTY (5 rule)",
    "Stage D — Rasio konversi QTY Terkecil 2",
    "Stage E-F — Standar rasio (mode) + merge",
    "Stage G-K — Selisih, deviasi, dominasi, arah",
    "Stage L-M — Status konversi + narasi",
    "Stage Merge — Kirim × Terima × Kategori (Selisih Hari + JENIS)",
    "Stage Count — Rekap count duplikat (4215*Count)",
    "Stage Summary — Ringkasan KPI",
]
PIPELINE_N_STAGES: int = len(PIPELINE_STAGES)


class PipelineCancelledError(RuntimeError):
    """Run dibatalkan user via cancel_check — bukan error sistem."""

    def __init__(self, msg: str, log: str = ""):
        super().__init__(msg)
        self.log = log


def run_pipeline(
    kirim_path: str | os.PathLike,
    terima_path: str | os.PathLike,
    kategori_path: str | os.PathLike,
    config_path: str | os.PathLike | None = None,
    progress_callback=None,
    cancel_check=None,
) -> dict:
    """Jalankan pipeline lengkap. Return summary dict (utk UI display).

    Args:
        kirim_path:    Path ke workbook 42.15 Pemindahan Barang (Kirim).
        terima_path:   Path ke workbook 42.05 Rincian Pemindahan (Terima).
        kategori_path: Path ke workbook DBS Kategori.
        config_path:   Path ke config.json custom (opsional).
        progress_callback: callable(msg: str, pct: float) utk update UI.
        cancel_check:  callable() -> bool, opsional. Dicek antar-stage.

    Returns:
        dict summary. Kunci penting: success (bool), df (DataFrame hasil
        pipeline Kirim), df_merged (DataFrame merge Kirim×Terima×Kategori),
        df_count (rekap count), summary (KPI), log (str).
    """
    try:
        return _run_pipeline(
            kirim_path, terima_path, kategori_path,
            config_path=config_path,
            progress_callback=progress_callback,
            cancel_check=cancel_check,
        )
    except PipelineCancelledError as exc:
        return {
            "success": False,
            "cancelled": True,
            "error": str(exc),
            "log": exc.log,
        }


def _run_pipeline(
    kirim_path: str | os.PathLike,
    terima_path: str | os.PathLike,
    kategori_path: str | os.PathLike,
    config_path: str | os.PathLike | None = None,
    progress_callback=None,
    cancel_check=None,
) -> dict:
    config = load_config(config_path)
    log_lines: list[str] = []
    last_pct = 0.0

    def _log(msg: str, pct: float | None = None) -> None:
        nonlocal last_pct
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        log_lines.append(line)
        if progress_callback:
            if pct is None:
                pct = last_pct
            else:
                last_pct = pct
            progress_callback(msg, pct)

    def _gate(stage_label: str, pct: float) -> None:
        """Log awal stage + cek cancel (antara stage)."""
        _log(f"{stage_label} ... ({pct:.0f}%)", pct)
        if cancel_check is not None and cancel_check():
            raise PipelineCancelledError(
                f"Run dibatalkan oleh user sebelum {stage_label}",
                log="\n".join(log_lines),
            )

    def _cancel_gate(where: str) -> None:
        """Cek cancel antar-file fase LOAD (tanpa tulis baris stage palsu)."""
        if cancel_check is not None and cancel_check():
            raise PipelineCancelledError(
                f"Run dibatalkan oleh user saat memuat file ({where})",
                log="\n".join(log_lines),
            )

    _log("=== SJ Insight Pipeline START ===", 0)
    _log(f"Kirim    : {kirim_path}", 1)
    _log(f"Terima   : {terima_path}", 2)
    _log(f"Kategori : {kategori_path}", 3)

    # ============================================================
    # STAGE 0 — LOAD 3 SUMBER
    # ============================================================
    _gate(PIPELINE_STAGES[0], 5)
    df_kirim_raw = load_kirim(kirim_path, config, logger=_log)
    _cancel_gate("Kirim (42.15)")
    df_terima_raw = load_terima(terima_path, config, logger=_log)
    _cancel_gate("Terima (42.05)")
    df_kategori_raw = load_kategori(kategori_path, config, logger=_log)
    _cancel_gate("Kategori")
    _log(f"  → {len(df_kirim_raw)} baris kirim · {len(df_terima_raw)} baris terima · "
         f"{len(df_kategori_raw)} baris kategori", 9)

    # ---- VALIDASI SKEMA (anti degradasi senyap) ----
    try:
        validate_kirim_schema(df_kirim_raw, config, logger=_log)
        validate_terima_schema(df_terima_raw, config, logger=_log)
        validate_kategori_schema(df_kategori_raw, config, logger=_log)
    except SchemaValidationError as exc:
        _log(f"ERROR skema input: {exc}", 100)
        return {
            "success": False,
            "error": str(exc),
            "log": "\n".join(log_lines),
        }

    # ============================================================
    # STAGE 1-3 — LOAD KIRIM: header, filter, tipe
    # ============================================================
    _gate(PIPELINE_STAGES[1], 10)
    df_kirim = stages.stage_01_03_load_kirim(df_kirim_raw, config)
    dropped_junk = df_kirim.attrs.get("dropped_junk", 0)
    _log(f"  → {len(df_kirim)} baris kirim valid ({dropped_junk} baris judul/junk dibuang)", 14)

    # ============================================================
    # STAGE 4-8 — KTG, Concet, ekstrak gudang, selisih, trim
    # ============================================================
    _gate(PIPELINE_STAGES[2], 18)
    stages.stage_04_08_transform(df_kirim, config)
    n_dc = int((df_kirim["KTG [RESTO/DC]"] == "DC").sum())
    _log(f"  → {n_dc} kiriman dari DC, {len(df_kirim) - n_dc} dari RESTO", 22)

    # ============================================================
    # STAGE A — ANOMALI SATUAN
    # ============================================================
    _gate(PIPELINE_STAGES[3], 25)
    stages.stageA_anomali_satuan(df_kirim)
    n_sat = int((df_kirim["Status Anomali Satuan"] != "NORMAL").sum())
    _log(f"  → {n_sat} baris anomali satuan", 28)

    # ============================================================
    # STAGE B — STATISTIK QTY
    # ============================================================
    _gate(PIPELINE_STAGES[4], 30)
    stages.stageB_qty_stats(df_kirim, config)

    # ============================================================
    # STAGE C — STATUS ANOMALI QTY (5 rule)
    # ============================================================
    _gate(PIPELINE_STAGES[5], 34)
    stages.stageC_anomali_qty(df_kirim, config)
    n_qty = int((df_kirim["Status Anomali QTY"] != "NORMAL").sum())
    _log(f"  → {n_qty} baris anomali QTY", 38)

    # ============================================================
    # STAGE D — QTY TERKECIL 2
    # ============================================================
    _gate(PIPELINE_STAGES[6], 40)
    stages.stageD_qty_terkecil_2(df_kirim)

    # ============================================================
    # STAGE E-F — STANDAR RASIO + MERGE
    # ============================================================
    _gate(PIPELINE_STAGES[7], 45)
    stages.stageEF_standar_rasio(df_kirim)

    # ============================================================
    # STAGE G-K — SELISIH / DEVIASI / DOMINASI / RASIO / ARAH
    # ============================================================
    _gate(PIPELINE_STAGES[8], 52)
    stages.stageGK_selisih_deviasi(df_kirim)

    # ============================================================
    # STAGE L-M — STATUS ANOMALI KONVERSI + NARASI
    # ============================================================
    _gate(PIPELINE_STAGES[9], 58)
    stages.stageLM_status_konversi(df_kirim, config)
    n_konv = int((df_kirim["Status Anomali QTY Terkecil 2"] == "ANOMALI KONVERSI").sum())
    _log(f"  → {n_konv} baris anomali konversi", 62)

    # ============================================================
    # MERGE — Kirim × Terima × Kategori
    # ============================================================
    _gate(PIPELINE_STAGES[10], 66)
    df_terima = stages.load_terima_rows(df_terima_raw, config)
    df_kategori = stages.load_kategori_rows(df_kategori_raw, config)
    df_merged = stages.merge_all(df_kirim, df_terima, df_kategori, config)
    tanpa_terima = df_merged.attrs.get("tanpa_terima", 0)
    dup_join = df_merged.attrs.get("terima_dup_count", 0)
    _log(
        f"  → {len(df_merged)} baris merge · {tanpa_terima} kirim tanpa dokumen terima"
        + (f" · {dup_join} nomor kirim duplikat di 42.05" if dup_join else ""),
        72,
    )

    # ============================================================
    # COUNT — Rekap duplikat (4215*Count)
    # ============================================================
    _gate(PIPELINE_STAGES[11], 76)
    df_count = stages.count_groups(df_kirim)
    n_dup = int((df_count["Count"] > 1).sum())
    _log(f"  → {n_dup} grup Concet dengan count > 1 (potensi duplikat)", 80)

    # ============================================================
    # SUMMARY — Ringkasan KPI
    # ============================================================
    _gate(PIPELINE_STAGES[12], 84)
    summary = sig.build_summary(df_kirim, df_merged, df_count, config)
    charts = sig.build_charts_data(df_kirim, df_merged)
    _log(
        f"  → {summary['total_transaksi']} transaksi · {summary['anomali_satuan']} anom satuan · "
        f"{summary['anomali_qty']} anom QTY · {summary['anomali_konversi']} anom konversi · "
        f"{summary['kirim_tanpa_terima']} tanpa terima",
        88,
    )

    _log("=== Pipeline DONE ===", 100)

    summary.update({
        "success":       True,
        "df":            df_kirim,
        "df_merged":     df_merged,
        "df_count":      df_count,
        "df_kategori":   df_kategori,
        "df_terima":     df_terima,
        "per_gudang":    charts["per_gudang"],
        "per_barang":    charts["per_barang"],
        "anomaly_breakdown": charts["anomaly_breakdown"],
        "log":           "\n".join(log_lines),
        "timestamp":     datetime.now().isoformat(),
        "kirim_path":    str(kirim_path),
        "terima_path":   str(terima_path),
        "kategori_path": str(kategori_path),
    })
    return summary
