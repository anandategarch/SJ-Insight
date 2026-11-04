"""theme.py — palet warna + QSS global untuk UI PySide6 (mirror Waste theme).

Tema SJ Insight: teal sebagai aksen (sesuai logo), slate sebagai base,
dengan token severity untuk pewarnaan baris tabel anomali.
"""
from __future__ import annotations

C_BG = "#f5f5f4"            # stone-100 — latar utama
C_CARD = "#ffffff"           # kartu
C_HEADER = "#1c1917"        # stone-900 — teks judul
C_BORDER = "#e7e5e4"        # stone-200 — border kartu
C_TEXT_MUTED = "#57534e"   # stone-600 — teks sekunder
C_ACCENT = "#0f766e"        # teal-700 — tombol aksi utama
C_ACCENT_HOVER = "#115e59"  # teal-800

C_TAB_IDLE_FG = "#44403c"   # stone-700 — teks tab non-aktif (WCAG AA)

# Status colors (mirror Waste) — (fg, bg tint)
STATUS_COLORS: dict[str, tuple[str, str]] = {
    "anomali":    ("#dc2626", "#fef2f2"),
    "perlu-cek":  ("#ea580c", "#fff7ed"),
    "perhatian":  ("#b45309", "#fffbeb"),
    "normal":     ("#16a34a", "#f0fdf4"),
}

SELECT_BG = "#ccfbf1"        # teal-100 — seleksi baris
SELECT_FG = "#134e4a"

UI_FONT_FAMILY = '"Segoe UI", "Tahoma", "DejaVu Sans", sans-serif'
UI_FONT_PX = 13
UI_HEADER_WEIGHT = 600

QSS = f"""
* {{
    font-family: {UI_FONT_FAMILY};
    font-size: {UI_FONT_PX}px;
}}
QMainWindow {{ background: {C_BG}; }}
QFrame#card {{
    background: {C_CARD};
    border: 1px solid {C_BORDER};
    border-radius: 8px;
}}
QFrame#kpiframe {{
    background: {C_CARD};
    border: 1px solid {C_BORDER};
    border-radius: 8px;
}}
QLabel#cardTitle {{
    color: {C_HEADER};
    font-weight: 600;
    background: transparent;
    border: none;
}}
QLabel#hint {{ color: {C_TEXT_MUTED}; background: transparent; border: none; font-size: 12px; }}
QLabel#muted {{ color: {C_TEXT_MUTED}; background: transparent; border: none; }}

QPushButton {{
    background: #ffffff;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 6px 14px;
}}
QPushButton:hover {{ border-color: #a8a29e; background: #fafaf9; }}
QPushButton:focus {{ border-color: {C_ACCENT}; }}
QPushButton:disabled {{ color: #a8a29e; background: #f5f5f4; }}

QPushButton#accent {{
    background: {C_ACCENT};
    color: #ffffff;
    border: 1px solid {C_ACCENT};
    font-weight: 600;
}}
QPushButton#accent:hover {{ background: {C_ACCENT_HOVER}; }}
QPushButton#accent:focus {{ border-color: #5eead4; }}
QPushButton#accent:disabled {{ background: #a8a29e; border-color: #a8a29e; color: #f5f5f4; }}

QPushButton#danger {{ background: #ffffff; color: #dc2626; border: 1px solid #fecaca; }}
QPushButton#danger:hover {{ background: #fef2f2; }}
QPushButton#danger:focus {{ border-color: #dc2626; }}

QPushButton#ghost {{ background: transparent; color: {C_ACCENT_HOVER}; border: 1px dashed {C_BORDER}; }}
QPushButton#ghost:hover {{ background: #f0fdfa; border-color: {C_ACCENT}; border-style: solid; }}
QPushButton#ghost:focus {{ border-color: {C_ACCENT}; }}

QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    background: #ffffff;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 4px 8px;
    min-height: 22px;
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{ border-color: {C_ACCENT}; }}
QComboBox QAbstractItemView {{
    background: #ffffff;
    border: 1px solid {C_BORDER};
    selection-background-color: {SELECT_BG};
    selection-color: {SELECT_FG};
}}

QTableView {{
    background: #ffffff;
    border: 1px solid {C_BORDER};
    gridline-color: #f5f5f4;
    selection-background-color: {SELECT_BG};
    selection-color: {SELECT_FG};
    alternate-background-color: #fafcfb;
}}
QHeaderView::section {{
    background: #fafaf9;
    color: {C_HEADER};
    border: none;
    border-right: 1px solid {C_BORDER};
    border-bottom: 1px solid {C_BORDER};
    padding: 6px 8px;
    font-weight: {UI_HEADER_WEIGHT};
}}
QTableView::item {{ padding: 3px 8px; }}

QTabWidget::pane {{ border: 1px solid {C_BORDER}; border-radius: 6px; background: {C_BG}; top: -1px; }}
QTabBar::tab {{
    background: #e7e5e4;
    color: {C_TAB_IDLE_FG};
    padding: 7px 12px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
}}
QTabBar::tab:selected {{ background: {C_CARD}; color: {C_HEADER}; font-weight: 600; }}

QProgressBar {{
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    background: #ffffff;
    text-align: center;
    min-height: 18px;
}}
QProgressBar::chunk {{ background: {C_ACCENT}; border-radius: 5px; }}

QPlainTextEdit#log {{
    background: #0c0a09;
    color: #d6d3d1;
    font-family: "Consolas", "DejaVu Sans Mono", monospace;
    font-size: 12px;
    border: 1px solid {C_BORDER};
    border-radius: 6px;
}}
QTextBrowser {{ background: #ffffff; border: 1px solid {C_BORDER}; border-radius: 6px; }}

QToolBar {{ background: {C_HEADER}; border: none; padding: 4px 8px; spacing: 8px; }}
QToolBar QToolButton {{
    background: transparent;
    color: #f5f5f4;
    border: none;
    border-radius: 6px;
    padding: 6px 12px;
    font-size: 13px;
}}
QToolBar QToolButton:hover {{ background: #292524; }}
QToolBar QToolButton:pressed {{ background: #44403c; }}
QToolBar::separator {{ background: #44403c; width: 1px; margin: 6px 4px; }}

QStatusBar {{ background: {C_CARD}; border-top: 1px solid {C_BORDER}; color: {C_TEXT_MUTED}; }}
QSplitter::handle {{ background: {C_BORDER}; }}
QSplitter::handle:horizontal {{ width: 2px; }}
QSplitter::handle:vertical {{ height: 2px; }}
QScrollArea {{ border: none; background: transparent; }}

QGroupBox {{
    background: {C_CARD};
    border: 1px solid {C_BORDER};
    border-radius: 8px;
    margin-top: 12px;
    font-weight: 600;
    color: {C_HEADER};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    top: 0px;
    padding: 0 4px;
}}
"""
