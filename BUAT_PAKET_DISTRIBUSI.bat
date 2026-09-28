@echo off
setlocal enabledelayedexpansion
title Pembuat Paket Distribusi Absensi Cleaner
cd /d "%~dp0"

echo ================================================================
echo       PEMBUAT PAKET DISTRIBUSI (WINDOWS ^& MAC)
echo ================================================================
echo.
echo Skrip ini akan otomatis membuat 2 paket bersih siap bagikan:
echo   1. HASIL_PAKET\Paket_Windows.zip (Tanpa perlu Python)
echo   2. HASIL_PAKET\Paket_Mac.zip     (Bersih khusus Mac)
echo.
echo ================================================================
pause

REM Buat folder penampung paket hasil
if not exist "HASIL_PAKET" mkdir "HASIL_PAKET"

REM ================================================================
REM  BAGIAN 1: MEMBUAT PAKET WINDOWS (.EXE)
REM ================================================================
echo.
echo [1/4] Mempersiapkan dependensi dan PyInstaller...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/4] Melakukan build executable Windows (PyInstaller)...
pyinstaller --noconfirm --onedir --console ^
  --name "AbsensiCleaner" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Gagal melakukan build PyInstaller Windows.
    pause
    exit /b %ERRORLEVEL%
)

REM Buat launcher praktis di dalam folder build Windows
(
echo @echo off
echo title Absensi Cleaner
echo cd /d "%%~dp0"
echo echo Membuka Absensi Cleaner...
echo timeout /t 1 /nobreak ^>nul
echo start "" "AbsensiCleaner.exe"
) > "dist\AbsensiCleaner\BUKA_APLIKASI.bat"

REM Buat file petunjuk Windows
(
echo ==================================================
echo            PANDUAN ABSENSI CLEANER ^(WINDOWS^)
echo ==================================================
echo.
echo CARA MEMBUKA:
echo 1. Pastikan folder ini sudah diekstrak ^(bukan dibuka dari zip^).
echo 2. Klik ganda file "BUKA_APLIKASI.bat" atau "AbsensiCleaner.exe".
echo 3. Browser akan otomatis terbuka.
echo 4. Komputer Anda TIDAK PERLU memiliki Python terpasang.
echo.
echo CATATAN WINDOWS DEFENDER:
echo Jika muncul peringatan biru "Windows protected your PC":
echo   -^> Klik "More info"
echo   -^> Klik tombol "Run anyway"
echo ==================================================
) > "dist\AbsensiCleaner\CARA_PAKAI_WINDOWS.txt"

echo.
echo [3/4] Mengompresi Paket Windows ke HASIL_PAKET\Paket_Windows.zip...
powershell -NoProfile -Command "Compress-Archive -Path 'dist\AbsensiCleaner\*' -DestinationPath 'HASIL_PAKET\Paket_Windows.zip' -Force"

REM ================================================================
REM  BAGIAN 2: MEMBUAT PAKET BERSIH KHUSUS MAC
REM ================================================================
echo.
echo [4/4] Menyiapkan Paket Bersih khusus macOS...

set "MAC_TEMP=HASIL_PAKET\AbsensiCleaner_Mac"
if exist "%MAC_TEMP%" rmdir /s /q "%MAC_TEMP%"
mkdir "%MAC_TEMP%"

copy "JALANKAN_MAC.command" "%MAC_TEMP%\" >nul
copy "jalankan.py" "%MAC_TEMP%\" >nul
copy "app.py" "%MAC_TEMP%\" >nul
copy "requirements.txt" "%MAC_TEMP%\" >nul
xcopy "templates" "%MAC_TEMP%\templates\" /s /e /i /q >nul
xcopy "static" "%MAC_TEMP%\static\" /s /e /i /q >nul

(
echo ==================================================
echo              PANDUAN ABSENSI CLEANER ^(MAC^)
echo ==================================================
echo.
echo CARA MEMBUKA:
echo 1. Pastikan file zip sudah diekstrak.
echo 2. Klik ganda file "JALANKAN_MAC.command" dari Finder.
echo 3. Browser Safari/Chrome akan otomatis terbuka.
echo.
echo PERSYARATAN:
echo - Komputer Mac memerlukan Python 3 terpasang.
echo   Jika belum ada, unduh dari: https://www.python.org/downloads/
echo.
echo JIKA MUNCUL PERINGATAN IZIN DI MAC:
echo Buka Terminal di folder ini, lalu ketik perintah:
echo   chmod +x JALANKAN_MAC.command
echo ==================================================
) > "%MAC_TEMP%\CARA_PAKAI_MAC.txt"

powershell -NoProfile -Command "Compress-Archive -Path '%MAC_TEMP%' -DestinationPath 'HASIL_PAKET\Paket_Mac.zip' -Force"
rmdir /s /q "%MAC_TEMP%"

echo.
echo ================================================================
echo                         SELESAI!
echo ================================================================
echo File paket yang siap Anda bagikan berada di folder HASIL_PAKET:
echo.
echo   1. HASIL_PAKET\Paket_Windows.zip  --^> Kirim ke teman Windows
echo   2. HASIL_PAKET\Paket_Mac.zip      --^> Kirim ke teman Mac
echo ================================================================
echo.
explorer HASIL_PAKET
pause
