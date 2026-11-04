"""workers.py — pekerja background (QThread) untuk pipeline & export.

Mengikuti pola Waste ui/workers:
  - PipelineWorker : run_pipeline + progress + cancel antar-stage
  - ExportWorker   : export_to_excel di background
"""
from __future__ import annotations

import traceback
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QThread, Signal

from sj_anomaly.exporter import export_to_excel
from sj_anomaly.pipeline import run_pipeline
from sj_anomaly.utils import load_config


class PipelineWorker(QThread):
    """Jalankan pipeline SJ Insight di background thread.

    Cancel efektif antar-stage (cancel_check).
    """

    progress = Signal(str, float)
    done = Signal(dict)
    failed = Signal(str, str)  # (pesan singkat, traceback)

    def __init__(
        self,
        kirim_path: str,
        terima_path: str,
        kategori_path: str,
        config_path: str | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._kirim_path = kirim_path
        self._terima_path = terima_path
        self._kategori_path = kategori_path
        self._config_path = config_path
        self._cancel = False

    def cancel(self) -> None:
        self._cancel = True

    def run(self) -> None:
        try:
            summary = run_pipeline(
                kirim_path=self._kirim_path,
                terima_path=self._terima_path,
                kategori_path=self._kategori_path,
                config_path=self._config_path,
                progress_callback=lambda msg, pct: self.progress.emit(str(msg), float(pct)),
                cancel_check=lambda: self._cancel,
            )
        except Exception as e:
            self.failed.emit(str(e), traceback.format_exc())
            return
        if summary.get("cancelled") or self._cancel:
            self.failed.emit("Run dibatalkan pengguna", "")
            return
        self.done.emit(summary)


class ExportWorker(QThread):
    """Tulis laporan Excel multi-sheet di background."""

    done = Signal(dict)
    failed = Signal(str)

    def __init__(
        self,
        df_merged: pd.DataFrame,
        summary: dict,
        path: str,
        config: dict,
        log_lines: list[str],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._df_merged = df_merged
        self._summary = summary
        self._path = path
        self._config = config
        self._log_lines = log_lines

    def run(self) -> None:
        try:
            result = export_to_excel(
                self._df_merged, self._path, self._config, self._summary, self._log_lines,
            )
        except Exception as e:
            self.failed.emit(str(e))
            return
        if isinstance(result, dict):
            self.done.emit(result)
        else:
            self.failed.emit(str(result))
