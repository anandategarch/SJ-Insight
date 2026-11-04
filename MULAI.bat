@echo off
REM MULAI.bat — jalankan SJ Insight (mirror Waste MULAI.bat)
REM
REM Asumsi: SETUP.bat sudah dijalankan sekali (dependensi terpasang).

setlocal
cd /d "%~dp0"

REM Generate sample data bila belum ada
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
