"""state.py — AppContext: satu sumber state + signal Qt (mirror Waste state).

Sinyal:
  - data_ready()      → dataset baru dari pipeline (df + summary)
  - filters_changed() → filter/drill-down berubah
  - status(text, kind) → pesan status bar ("running"|"done"|"error")
  - log_line(str)     → satu baris log
  - followup_changed() → daftar follow-up berubah
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QObject, Signal

from sj_anomaly.followup import (
    fingerprint_series,
    load_followup_keys,
    save_followup_keys,
)

from .shared import LOG_RUN_LINES


class AppContext(QObject):
    """State aplikasi tunggal + signal perubahan."""

    data_ready = Signal()
    filters_changed = Signal()
    followup_changed = Signal()
    status = Signal(str, str)
    log_line = Signal(str)

    def __init__(self, config: dict, parent: QObject | None = None,
                 followup_path: Path | None = None) -> None:
        super().__init__(parent)
        self.config = config
        self.df: pd.DataFrame | None = None             # df_kirim hasil pipeline
        self.df_merged: pd.DataFrame | None = None
        self.df_count: pd.DataFrame | None = None
        self.summary: dict | None = None
        self.log_lines: list[str] = []
        self.last_run_time: datetime | None = None
        self.data_version = 0

        # Filter state (sederhana — anomaly_status, gudang_tujuan, jenis)
        self.active_filters: dict[str, str | list[str] | None] = {
            "gudang": None,
            "barang": None,
            "status": None,
            "jenis": None,
        }

        # Followup (centang baris)
        self._followup_path = (
            None if followup_path is None else Path(str(followup_path))
        )
        self.followup_keys: list[str] = (
            load_followup_keys(self._followup_path)
            if self._followup_path is not None else []
        )
        self._fp_series: pd.Series | None = None
        self._fp_version: int = -1

    # ------------------------------------------------------------
    # DATA
    # ------------------------------------------------------------
    def set_dataset(self, df: pd.DataFrame, df_merged: pd.DataFrame,
                    df_count: pd.DataFrame, summary: dict) -> None:
        """Dataset baru dari pipeline: simpan + reset filter + broadcast."""
        self.df = df
        self.df_merged = df_merged
        self.df_count = df_count
        self.summary = summary
        self.last_run_time = datetime.now()
        self.data_version += 1
        self._fp_series = None  # invalidate
        self.reset_filters()
        self.data_ready.emit()
        self.filters_changed.emit()

    # ------------------------------------------------------------
    # FILTER (sederhana)
    # ------------------------------------------------------------
    def filtered_df(self) -> pd.DataFrame | None:
        """df hasil filter funnel."""
        if self.df is None:
            return None
        df = self.df
        af = self.active_filters
        if af.get("gudang"):
            df = df[df["Gudang Tujuan (Kode)"] == af["gudang"]]
        if af.get("barang"):
            df = df[df["Nama Barang"] == af["barang"]]
        if af.get("jenis") and "JENIS (Kategori)" in df.columns:
            df = df[df["JENIS (Kategori)"] == af["jenis"]]
        if af.get("status") and "Status Anomali QTY Terkecil 2" in df.columns:
            df = df[df["Status Anomali QTY Terkecil 2"] == af["status"]]
        return df

    def anomali_df(self) -> pd.DataFrame | None:
        """df terfilter, hanya baris dengan salah satu status ≠ NORMAL."""
        df = self.filtered_df()
        if df is None or df.empty:
            return df
        mask = (
            (df["Status Anomali Satuan"] != "NORMAL")
            | (df["Status Anomali QTY"] != "NORMAL")
            | (df["Status Anomali QTY Terkecil 2"] != "NORMAL")
        )
        # Tambah: kirim tanpa dokumen terima juga dianggap anomali
        if self.df_merged is not None and "Memiliki Dokumen [Terima]" in self.df_merged.columns:
            no_to_dok = dict(zip(self.df_merged["No. Pemindahan #"], self.df_merged["Memiliki Dokumen [Terima]"]))
            no_dok_series = df["No. Pemindahan #"].map(no_to_dok).fillna(True)
            mask = mask | (~no_dok_series)
        return df[mask]

    def set_filter(self, key: str, value) -> None:
        self.active_filters[key] = value
        self.filters_changed.emit()

    def reset_filters(self) -> None:
        for k in self.active_filters:
            self.active_filters[k] = None
        # Tidak broadcast — set_dataset yang emit, atau pemanggil eksplisit.

    # ------------------------------------------------------------
    # FOLLOW UP
    # ------------------------------------------------------------
    def followup_set(self) -> set[str]:
        return set(self.followup_keys)

    def toggle_followup(self, fp: str, checked: bool) -> bool:
        """Tambah/hapus SATU baris. Return True bila berubah."""
        fset = self.followup_set()
        if checked == (fp in fset):
            return False
        if checked:
            self.followup_keys.append(fp)
        else:
            self.followup_keys = [k for k in self.followup_keys if k != fp]
        self._persist_followup()
        self.followup_changed.emit()
        return True

    def clear_followup(self) -> int:
        n = len(self.followup_keys)
        if n == 0:
            return 0
        self.followup_keys = []
        self._persist_followup()
        self.followup_changed.emit()
        return n

    def _persist_followup(self) -> None:
        if self._followup_path is None:
            return
        if not save_followup_keys(self.followup_keys, self._followup_path):
            self.set_status("⚠ Gagal menyimpan daftar follow-up (disk?)", "error")

    def fingerprints(self) -> pd.Series | None:
        """Sidik jari per baris dataset PENUH (memo per data_version)."""
        if self.df is None:
            return None
        if self._fp_series is None or self._fp_version != self.data_version:
            self._fp_series = fingerprint_series(self.df)
            self._fp_version = self.data_version
        return self._fp_series

    # ------------------------------------------------------------
    # LOG / STATUS
    # ------------------------------------------------------------
    def log(self, msg: str) -> None:
        self.log_lines.append(str(msg))
        if len(self.log_lines) > LOG_RUN_LINES:
            del self.log_lines[: len(self.log_lines) - LOG_RUN_LINES]
        self.log_line.emit(str(msg))

    def set_status(self, text: str, kind: str = "running") -> None:
        self.status.emit(text, kind)
