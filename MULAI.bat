@echo off
REM MULAI.bat — jalankan SJ Insight (v1.1.0 — auto-aktifkan virtualenv)
REM
REM Asumsi: SETUP.bat sudah dijalankan sekali (virtualenv .venv ada).

setlocal
cd /d "%~dp0"

REM ---- Aktifkan virtualenv bila ada ----
if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
) else (
    echo WARNING: Virtualenv .venv tidak ditemukan.
    echo Jalankan SETUP.bat dulu untuk instalasi.
    pause
    exit /b 1
)

REM ---- Generate sample data bila belum ada ----
if not exist "data\42.15_Pemindahan_Barang_KIRIM.xlsx" (
    echo Generate sample data...
    python make_sample.py
    echo.
)

echo === SJ Insight — Memulai aplikasi ===
python ui_app.py
if errorlevel 1 (
    echo.
    echo Aplikasi berhenti dengan error. Cek logs\sj_insight.log
    pause
)
