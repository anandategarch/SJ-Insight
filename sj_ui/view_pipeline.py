"""view_pipeline.py — tab Pipeline: input + run + log (mirror Waste view_pipeline).

Komponen:
  - 3 input file: Kirim / Terima / Kategori (browse + Pakai Sample)
  - Tombol Run / Cancel
  - Progress bar + persen
  - Daftar stage (PIPELINE_STAGES) dengan status
  - Log proses (QPlainTextEdit read-only, max 2000 baris)
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from sj_anomaly.pipeline import PIPELINE_STAGES

from .shared import APP_VERSION
from .state import AppContext
from .workers import PipelineWorker


class PipelineView(QWidget):
    """Tab Pipeline: input 3 sumber + run + log."""

    run_requested = Signal(str, str, str)  # kirim, terima, kategori
    cancel_requested = Signal()

    def __init__(self, ctx: AppContext, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("pipelineView")
        self.ctx = ctx
        self._worker: PipelineWorker | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # ---- Header ----
        title = QLabel("Pipeline Analisa — Replika M code '04.QUERRY SJ.xlsb'")
        title.setStyleSheet("font-size: 16px; font-weight: 700; color: #1c1917;")
        layout.addWidget(title)
        sub = QLabel(
            f"v{APP_VERSION} · 13 stage replika Power Query · 3 sumber Excel (42.15 / 42.05 / Kategori)"
        )
        sub.setObjectName("muted")
        sub.setStyleSheet("font-size: 12px;")
        layout.addWidget(sub)

        # ---- Input cards ----
        input_row = QHBoxLayout()
        input_row.setSpacing(8)
        self._kirim_edit = self._build_input_card(
            "42.15 Pemindahan Barang (Kirim)",
            "Sheet '#42.15. Pemindahan Barang (KIRI' — report ACCURATE sisi kirim",
            "kirim",
            input_row,
        )
        self._terima_edit = self._build_input_card(
            "42.05 Rincian Pemindahan (Terima)",
            "Sheet '#42.05. Rincian Pemindahan Bara' — bukti penerimaan",
            "terima",
            input_row,
        )
        self._kategori_edit = self._build_input_card(
            "DBS Kategori",
            "Sheet 'Sheet1' — master kategori & jenis barang (NAMA BARANG / KATEGORI / JENIS)",
            "kategori",
            input_row,
        )
        layout.addLayout(input_row)

        # ---- Actions ----
        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        self._run_btn = QPushButton("▶ Jalankan Analisa")
        self._run_btn.setObjectName("accent")
        self._run_btn.setCursor(Qt.PointingHandCursor)
        self._run_btn.clicked.connect(self._on_run)
        action_row.addWidget(self._run_btn)

        self._cancel_btn = QPushButton("⏹ Batalkan Run")
        self._cancel_btn.setCursor(Qt.PointingHandCursor)
        self._cancel_btn.clicked.connect(self._on_cancel)
        self._cancel_btn.setEnabled(False)
        action_row.addWidget(self._cancel_btn)

        action_row.addStretch()
        self._status_label = QLabel("")
        self._status_label.setObjectName("muted")
        action_row.addWidget(self._status_label)
        layout.addLayout(action_row)

        # ---- Progress ----
        prog_row = QHBoxLayout()
        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setTextVisible(True)
        prog_row.addWidget(self._progress, 1)
        self._pct_label = QLabel("0%")
        self._pct_label.setMinimumWidth(40)
        prog_row.addWidget(self._pct_label)
        layout.addLayout(prog_row)

        # ---- Stage list ----
        stage_label = QLabel("Daftar Stage:")
        stage_label.setStyleSheet("font-weight: 600; color: #1c1917;")
        layout.addWidget(stage_label)
        self._stage_labels: list[QLabel] = []
        stage_container = QFrame()
        stage_container.setObjectName("card")
        sl = QVBoxLayout(stage_container)
        sl.setContentsMargins(10, 8, 10, 8)
        sl.setSpacing(2)
        for stage in PIPELINE_STAGES:
            lbl = QLabel(stage)
            lbl.setObjectName("muted")
            lbl.setStyleSheet("font-size: 12px;")
            self._stage_labels.append(lbl)
            sl.addWidget(lbl)
        layout.addWidget(stage_container)

        # ---- Log ----
        log_label = QLabel("Log Run:")
        log_label.setStyleSheet("font-weight: 600; color: #1c1917;")
        layout.addWidget(log_label)
        self._log = QPlainTextEdit()
        self._log.setObjectName("log")
        self._log.setReadOnly(True)
        self._log.setMaximumBlockCount(2000)
        layout.addWidget(self._log, 1)

        # Connect signals
        self.ctx.log_line.connect(self._on_log_line)
        self.ctx.status.connect(self._on_status)

    # ------------------------------------------------------------
    # Build helpers
    # ------------------------------------------------------------
    def _build_input_card(
        self,
        title: str,
        hint: str,
        kind: str,
        layout: QHBoxLayout,
    ) -> QLineEdit:
        """Bangun kartu input file dengan tombol Browse + Sample."""
        card = QFrame()
        card.setObjectName("card")
        v = QVBoxLayout(card)
        v.setContentsMargins(12, 10, 12, 10)
        v.setSpacing(4)

        t = QLabel(title)
        t.setObjectName("cardTitle")
        v.addWidget(t)
        h = QLabel(hint)
        h.setObjectName("hint")
        h.setWordWrap(True)
        v.addWidget(h)

        row = QHBoxLayout()
        edit = QLineEdit()
        edit.setPlaceholderText(f"Pilih file .xlsx utk {title}")
        row.addWidget(edit, 1)
        browse = QPushButton("Telusuri…")
        browse.clicked.connect(lambda: self._browse(edit))
        row.addWidget(browse)
        v.addLayout(row)

        layout.addWidget(card, 1)
        return edit

    def _browse(self, edit: QLineEdit) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih file Excel", "", "Excel Workbook (*.xlsx *.xlsm);;Semua file (*)"
        )
        if path:
            edit.setText(path)

    # ------------------------------------------------------------
    # Run / Cancel
    # ------------------------------------------------------------
    def _on_run(self) -> None:
        kirim = self._kirim_edit.text().strip()
        terima = self._terima_edit.text().strip()
        kategori = self._kategori_edit.text().strip()
        if not (kirim and terima and kategori):
            self._on_status("Lengkapi ketiga file input terlebih dahulu.", "error")
            return
        if not (Path(kirim).exists() and Path(terima).exists() and Path(kategori).exists()):
            self._on_status("Salah satu file tidak ditemukan — cek path.", "error")
            return

        # Reset UI
        self._log.clear()
        self._progress.setValue(0)
        self._pct_label.setText("0%")
        for lbl in self._stage_labels:
            lbl.setStyleSheet("font-size: 12px; color: #57534e;")

        # Disable run, enable cancel
        self._run_btn.setEnabled(False)
        self._cancel_btn.setEnabled(True)

        # Start worker
        self._worker = PipelineWorker(kirim, terima, kategori, parent=self)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_cancel(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._on_status("Membatalkan run…", "running")

    # ------------------------------------------------------------
    # Callbacks
    # ------------------------------------------------------------
    def _on_progress(self, msg: str, pct: float) -> None:
        # Highlight stage terkini
        for lbl in self._stage_labels:
            stage_text = lbl.text()
            if stage_text in msg and "..." in msg:
                lbl.setStyleSheet("font-size: 12px; color: #0f766e; font-weight: 600;")
            elif stage_text in msg:
                lbl.setStyleSheet("font-size: 12px; color: #16a34a;")
        self._progress.setValue(int(pct))
        self._pct_label.setText(f"{pct:.0f}%")

    def _on_done(self, summary: dict) -> None:
        if not summary.get("success"):
            self._on_failed(summary.get("error", "Pipeline gagal tanpa pesan"), "")
            return
        self._progress.setValue(100)
        self._pct_label.setText("100%")
        self._run_btn.setEnabled(True)
        self._cancel_btn.setEnabled(False)
        self._on_status(
            f"Selesai · {summary.get('total_transaksi', 0)} transaksi · "
            f"{summary.get('anomali_satuan', 0) + summary.get('anomali_qty', 0) + summary.get('anomali_konversi', 0)} anomali",
            "done",
        )
        # Set dataset ke context → trigger data_ready (tab lain render)
        df = summary.pop("df")
        df_merged = summary.pop("df_merged")
        df_count = summary.pop("df_count")
        summary.pop("df_kategori", None)
        summary.pop("df_terima", None)
        summary["df_count"] = df_count  # kembalikan utk exporter
        self.ctx.set_dataset(df, df_merged, df_count, summary)

    def _on_failed(self, msg: str, tb: str) -> None:
        self._run_btn.setEnabled(True)
        self._cancel_btn.setEnabled(False)
        self._on_status(f"Gagal: {msg}", "error")
        if tb:
            self._log.appendPlainText(f"\n=== TRACEBACK ===\n{tb}")

    def _on_log_line(self, msg: str) -> None:
        self._log.appendPlainText(msg)

    def _on_status(self, text: str, kind: str) -> None:
        colors = {
            "running": "#0f766e",
            "done": "#16a34a",
            "error": "#dc2626",
        }
        self._status_label.setText(text)
        self._status_label.setStyleSheet(f"color: {colors.get(kind, '#57534e')}; font-size: 12px;")
