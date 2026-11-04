"""app.py — MainWindow + main() (mirror Waste app.py, v2.6.0 rebuild).

Struktur:
  - Toolbar: Ekspor Excel + Bantuan
  - QTabWidget 4 tab: Dashboard · Anomali · Kirim vs Terima · Pipeline
  - Status bar: pesan status + hitungan + waktu run
"""
from __future__ import annotations

import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QAction, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QStatusBar,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QDialog,
)

from sj_anomaly.followup import FP_COL  # noqa: F401  (constant contract)
from sj_anomaly.logging_setup import setup_logging
from sj_anomaly.utils import load_config

from .shared import APP_DIR, APP_TITLE, APP_VERSION, LOG_RUN_LINES
from .state import AppContext
from .theme import QSS
from .view_anomali import AnomaliView
from .view_dashboard import DashboardView
from .view_merge import MergeView
from .view_pipeline import PipelineView
from .workers import ExportWorker


def _build_help_html() -> str:
    """Bangun HTML Bantuan (F1)."""
    return f"""
<h2>Panduan Singkat — SJ Insight v{APP_VERSION}</h2>
<p><b>Alur kerja:</b> tab <b>Pipeline</b> → centang <b>📂 Mode Folder Otomatis</b>
(atau pilih 3 file manual) → <b>▶ Jalankan Analisa</b> → buka tab
<b>Dashboard</b> untuk KPI / <b>Anomali</b> untuk tabel baris mencurigakan /
<b>Kirim vs Terima</b> untuk rekonsiliasi dokumen.</p>
<p><b>Shortcut:</b> <b>Ctrl+1..Ctrl+4</b> pindah tab · <b>Ctrl+E</b> ekspor Excel
· <b>F1</b> bantuan ini.</p>
<h3>Tab Pipeline (tempat semua dimulai)</h3>
<ul>
<li><b>📂 Mode Folder Otomatis</b> — baca SEMUA file .xlsx dari folder yang
sama persis dengan Power Query M code di <code>04.QUERRY SJ.xlsb</code>:
<code>C:\\KERJA\\1. RESTO\\9.NEW QUERY\\3. 42.15</code> (Kirim),
<code>\\4. 42.05</code> (Terima), <code>\\21. DBS Kategori</code>.
Mirror dari <code>Folder.Files("...")</code> + invoke transform. Tombol
<b>🔄 Isi Otomatis dari Config</b> mengisi 3 field dengan path dari
<code>config.json</code>.</li>
<li><b>Mode File Manual</b> (default) — pilih 3 file .xlsx individually via
<b>Telusuri…</b>. Bisa juga pilih <b>folder</b> (semua .xlsx di folder dibaca).</li>
<li>42.15 Kirim — report ACCURATE sisi pengiriman (sheet '#42.15. Pemindahan
Barang (KIRI'). 42.05 Terima — bukti penerimaan (sheet '#42.05. Rincian
Pemindahan Bara'). DBS Kategori — master kategori (sheet 'Sheet1').</li>
<li><b>Auto-detect header row</b> — program otomatis mencari baris header
berdasarkan keyword ('Nama Barang', 'Satuan', 'Nomor #Kirim', dst.). Ini
menangani report ACCURATE dengan jumlah baris judul yang bervariasi (2/3/4)
dan header merged-cell (kolom genap terbaca 'nan' / 'ColumnN' otomatis
dilewati).</li>
<li><b>▶ Jalankan Analisa ↔ ⏹ Batalkan Run</b> — satu tombol dua-mode;
pembatalan efektif di gerbang antar-stage (F9).</li>
<li>Selama run: <b>progres</b> (persen) + <b>log proses</b> per stage terlihat.</li>
<li>Setelah selesai: tab lain berisi hasil analisa.</li>
</ul>
<h3>Tab Dashboard</h3>
<ul>
<li>8 kartu KPI: Total Transaksi · Anomali Satuan · Anomali QTY · Anomali
Konversi · Perlu Cek · Kirim tanpa Terima · Terima Lambat · Grup Duplikat.</li>
<li>Tabel <b>Top 12 Gudang Tujuan</b> & <b>Top 12 Barang</b> by volume/anomali.</li>
<li>Periode + sumber file di bagian atas.</li>
</ul>
<h3>Tab Anomali</h3>
<ul>
<li>Tabel baris dengan salah satu status ≠ NORMAL (Anomali Satuan / QTY /
Konversi / dokumen terima hilang).</li>
<li><b>Filter</b>: Gudang / Barang / Status Konversi. Klik header kolom untuk sort.</li>
<li>Sel berwarna = severity (merah=anomali, oranye=perlu cek, kuning=perhatian,
hijau=normal).</li>
<li>Tooltip sel = teks penuh (kolom panjang tidak pernah terpotong).</li>
</ul>
<h3>Tab Kirim vs Terima</h3>
<ul>
<li>Left join Kirim ← Terima on No. Pemindahan # = Nomor #Kirim.</li>
<li><b>Selisih Hari</b> = Tanggal #Terima − Tanggal kirim (≥ 2 hari = terima lambat).</li>
<li>Tombol <b>⚑ Hanya Kirim tanpa Terima</b> & <b>⚑ Hanya Terima Lambat</b> — toggle filter.</li>
</ul>
<h3>Ekspor Excel</h3>
<p>Tombol <b>💾 Ekspor Excel</b> di toolbar (aktif setelah pipeline selesai) —
laporan 6 sheet: Detail · Anomali Prioritas · Ringkasan · Rekap Count · Log Run
(+ extra bila ada).</p>
<h3>💡 Sumber Spesifikasi</h3>
<p>Program ini adalah replika Power Query M code dari file
<b>04.QUERRY SJ.xlsb</b> (29 query, 1.034 baris M code). Semua mapping header,
aturan anomali (satuan/QTY/konversi), dan join kirim-terima mengikuti M code
tersebut persis. Path folder sumber di <code>config.json</code> sama dengan
<code>Folder.Files("C:\\KERJA\\1. RESTO\\9.NEW QUERY\\...")</code> di M code.
Threshold bisa di-override via <code>config.json</code>.</p>
"""


class MainWindow(QMainWindow):
    """Jendela utama SJ Insight."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1280, 820)
        self.setMinimumSize(960, 640)

        # ---- Config & context ----
        config = load_config()
        followup_path = APP_DIR / "followup.json"
        self.ctx = AppContext(config, parent=self, followup_path=followup_path)

        # ---- Toolbar ----
        self._build_toolbar()

        # ---- Tabs ----
        self._tabs = QTabWidget()
        self._tabs.setDocumentMode(True)
        self._dashboard = DashboardView(self.ctx)
        self._anomali = AnomaliView(self.ctx)
        self._merge = MergeView(self.ctx)
        self._pipeline = PipelineView(self.ctx)
        self._tabs.addTab(self._pipeline, "⚙ Pipeline")
        self._tabs.addTab(self._dashboard, "📊 Dashboard")
        self._tabs.addTab(self._anomali, "🚨 Anomali")
        self._tabs.addTab(self._merge, "↔ Kirim vs Terima")
        self.setCentralWidget(self._tabs)

        # Default tab: Pipeline (tempat mulai)
        self._tabs.setCurrentIndex(0)

        # ---- Status bar ----
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._status_text = QLabel("Siap. Pilih 3 file di tab Pipeline untuk mulai.")
        self._status_text.setStyleSheet("color: #57534e;")
        sb.addWidget(self._status_text, 1)
        self._status_count = QLabel("")
        sb.addPermanentWidget(self._status_count)
        self._status_time = QLabel("")
        sb.addPermanentWidget(self._status_time)

        # Connect signals
        self.ctx.status.connect(self._on_status)
        self.ctx.data_ready.connect(self._on_data_ready)

        # Export worker slot
        self._export_worker: ExportWorker | None = None

        # Shortcuts
        self._setup_shortcuts()

    # ------------------------------------------------------------
    # Toolbar
    # ------------------------------------------------------------
    def _build_toolbar(self) -> None:
        tb = self.addToolBar("Utama")
        tb.setMovable(False)
        tb.setIconSize(QSize(18, 18))

        act_export = QAction("💾 Ekspor Excel", self)
        act_export.setShortcut(QKeySequence("Ctrl+E"))
        act_export.triggered.connect(self._on_export)
        act_export.setEnabled(False)
        self._act_export = act_export
        tb.addAction(act_export)

        tb.addSeparator()

        act_help = QAction("❓ Bantuan", self)
        act_help.setShortcut(QKeySequence("F1"))
        act_help.triggered.connect(self._on_help)
        tb.addAction(act_help)

    def _setup_shortcuts(self) -> None:
        """Ctrl+1..4 pindah tab."""
        for i in range(4):
            sc = QKeySequence(f"Ctrl+{i + 1}")
            act = QAction(self)
            act.setShortcut(sc)
            act.triggered.connect(lambda _=False, idx=i: self._tabs.setCurrentIndex(idx))
            self.addAction(act)

    # ------------------------------------------------------------
    # Callbacks
    # ------------------------------------------------------------
    def _on_status(self, text: str, kind: str) -> None:
        colors = {"running": "#0f766e", "done": "#16a34a", "error": "#dc2626"}
        self._status_text.setText(text)
        self._status_text.setStyleSheet(f"color: {colors.get(kind, '#57534e')};")

    def _on_data_ready(self) -> None:
        """Aktifkan ekspor & update status bar."""
        s = self.ctx.summary
        if not s:
            return
        self._act_export.setEnabled(True)
        anom = (
            s.get("anomali_satuan", 0)
            + s.get("anomali_qty", 0)
            + s.get("anomali_konversi", 0)
            + s.get("kirim_tanpa_terima", 0)
        )
        self._status_count.setText(
            f"{s.get('total_transaksi', 0)} transaksi · {anom} anomali"
        )
        if self.ctx.last_run_time:
            self._status_time.setText(
                self.ctx.last_run_time.strftime("%d %b %Y %H:%M:%S")
            )
        # Pindah ke Dashboard otomatis setelah run selesai
        self._tabs.setCurrentIndex(1)

    # ------------------------------------------------------------
    # Export
    # ------------------------------------------------------------
    def _on_export(self) -> None:
        if self.ctx.df_merged is None or self.ctx.df_merged.empty:
            QMessageBox.warning(self, "Ekspor", "Belum ada data untuk diekspor — jalankan pipeline dulu.")
            return
        default_name = f"SJ_Analisa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Simpan Laporan Excel", default_name,
            "Excel Workbook (*.xlsx);;Semua file (*)",
        )
        if not path:
            return

        # Jalankan ExportWorker di background
        self._export_worker = ExportWorker(
            df_merged=self.ctx.df_merged,
            summary={**(self.ctx.summary or {}), "df_count": self.ctx.df_count},
            path=path,
            config=self.ctx.config,
            log_lines=self.ctx.log_lines[-LOG_RUN_LINES:],
            parent=self,
        )
        self._export_worker.done.connect(self._on_export_done)
        self._export_worker.failed.connect(self._on_export_failed)
        self._act_export.setEnabled(False)
        self._status_text.setText("Mengekspor Excel…")
        self._export_worker.start()

    def _on_export_done(self, result: dict) -> None:
        self._act_export.setEnabled(True)
        if result.get("success"):
            QMessageBox.information(
                self, "Ekspor Selesai",
                f"Laporan tersimpan: {result.get('output_path')}\n\n"
                f"Total baris: {result.get('total_rows', 0)}\n"
                f"Anomali: {result.get('anomali_count', 0)}\n"
                f"Sheet: 6 (Detail · Anomali Prioritas · Ringkasan · Rekap Count · Log Run)",
            )
            self._status_text.setText(f"Ekspor selesai: {result.get('output_path')}")
        else:
            QMessageBox.warning(self, "Ekspor Gagal", result.get("error", "Kesalahan tak diketahui."))

    def _on_export_failed(self, msg: str) -> None:
        self._act_export.setEnabled(True)
        QMessageBox.warning(self, "Ekspor Gagal", msg)

    # ------------------------------------------------------------
    # Help dialog
    # ------------------------------------------------------------
    def _on_help(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Bantuan — SJ Insight")
        dlg.resize(720, 640)
        v = QVBoxLayout(dlg)
        v.setContentsMargins(0, 0, 0, 0)
        browser = QTextBrowser()
        browser.setHtml(_build_help_html())
        browser.setOpenExternalLinks(True)
        v.addWidget(browser)
        dlg.exec()


def _qt_excepthook(exc_type, exc_value, exc_tb):
    """Global exception hook — log + tampilkan ke user."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    logging.getLogger("sj_anomaly").error(
        "Uncaught exception", exc_info=(exc_type, exc_value, exc_tb)
    )
    traceback.print_exception(exc_type, exc_value, exc_tb)


def main(argv: list[str] | None = None) -> int:
    """Entry point aplikasi."""
    argv = argv if argv is not None else sys.argv
    setup_logging()
    sys.excepthook = _qt_excepthook

    app = QApplication(argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("SJ Insight")
    app.setStyleSheet(QSS)

    # Coba pasang ikon (opsional)
    icon_path = Path(__file__).resolve().parent.parent / "assets" / "app_icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
