@echo off
REM SETUP.bat — instalasi sekali untuk SJ Insight (v1.1.0 — robust Windows Long Path)
REM
REM Strategi:
REM   1. Buat virtualenv di folder pendek (.venv di root repo) supaya path
REM      PySide6 (yang nested-nya panjang) tetap di bawah 260 char limit Windows.
REM   2. Pasang dependensi ke virtualenv itu (isolasi, tidak sentuh global).
REM   3. MULAI.bat akan auto-aktifkan virtualenv.
REM
REM Alternatif bila gagal: enable Windows Long Paths via PowerShell admin:
REM   New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" `
REM     -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force

setlocal
cd /d "%~dp0"

echo === SJ Insight - Setup ===
echo.

REM ---- Cek Python ----
echo Memeriksa Python...
python --version 2>nul
if errorlevel 1 (
    echo ERROR: Python tidak ditemukan di PATH.
    echo Install Python 3.10+ dari https://python.org lalu ulangi.
    pause
    exit /b 1
)

REM ---- Buat virtualenv di path pendek ----
if exist ".venv\Scripts\python.exe" (
    echo.
    echo Virtualenv sudah ada di .venv — skip pembuatan.
) else (
    echo.
    echo Membuat virtualenv di .venv ^(path pendek — anti Windows Long Path issue^)...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Gagal membuat virtualenv. Coba manual:
        echo   python -m venv .venv
        pause
        exit /b 1
    )
)

REM ---- Aktifkan virtualenv ----
call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo ERROR: Gagal mengaktifkan virtualenv.
    pause
    exit /b 1
)

REM ---- Upgrade pip ----
echo.
echo Upgrade pip...
python -m pip install --upgrade pip

REM ---- Install dependensi ----
echo.
echo Menginstal dependensi ke virtualenv...
echo   ^(pandas, openpyxl, PySide6, python-calamine^)
echo.
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: Gagal menginstal dependensi.
    echo ========================================
    echo.
    echo Kemungkinan penyebab paling umum: Windows Long Path BELUM di-enable.
    echo PySide6 punya nested path ^>260 char ^(mis. PySide6\qml\Qt\labs\...\rcc\*.obj^).
    echo.
    echo SOLUSI 1 ^(recommended, sekali buat selamanya^):
    echo   Buka PowerShell sebagai Administrator, lalu jalankan:
    echo   New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
    echo   Lalu tutup cmd ini, buka baru, jalankan SETUP.bat lagi.
    echo.
    echo SOLUSI 2 ^(tanpa admin^):
    echo   Pindahkan folder SJ-Insight ke root pendek ^(mis. C:\SJ^) lalu jalankan SETUP.bat.
    echo.
    echo SOLUSI 3 ^(manual^):
    echo   pip install -r requirements.txt --no-cache-dir
    echo.
    pause
    exit /b 1
)

echo.
echo === Setup selesai! ===
echo Virtualenv aktif di: .venv\
echo.
echo Jalankan MULAI.bat untuk membuka aplikasi.
pause
