@echo off
REM SETUP.bat — instalasi sekali untuk SJ Insight (mirror Waste SETUP.bat)
REM
REM Menjalankan: pip install -r requirements.txt
REM Setelah itu jalankan MULAI.bat untuk membuka aplikasi.

setlocal
cd /d "%~dp0"

echo === SJ Insight - Setup ===
echo.
echo Memeriksa Python...
python --version 2>nul
if errorlevel 1 (
    echo ERROR: Python tidak ditemukan di PATH.
    echo Install Python 3.10+ dari https://python.org lalu ulangi.
    pause
    exit /b 1
)

echo.
echo Menginstal dependensi (pandas, openpyxl, PySide6, python-calamine)...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Gagal menginstal dependensi. Coba jalankan manual:
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

echo.
echo === Setup selesai! ===
echo Jalankan MULAI.bat untuk membuka aplikasi.
pause
