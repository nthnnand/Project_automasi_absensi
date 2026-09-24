@echo off
title Build Absensi Cleaner ke EXE
echo ================================================
echo   MEMBUAT FILE EXE ABSENSI CLEANER
echo ================================================
echo.

cd /d "%~dp0"

echo [1/3] Menginstall library pendukung & PyInstaller...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/3] Memproses build PyInstaller...
pyinstaller --noconfirm --onedir --console ^
  --name "AbsensiCleaner" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  app.py

echo.
echo [3/3] Menyiapkan shortcut peluncur...
(
echo @echo off
echo title Absensi Cleaner
echo echo Membuka Absensi Cleaner...
echo timeout /t 2 /nobreak ^>nul
echo start "" "http://localhost:5000"
echo "%%~dp0AbsensiCleaner.exe"
) > "dist\AbsensiCleaner\BUKA_APLIKASI.bat"

echo.
echo ================================================
echo   SELESAI!
echo   Folder aplikasi siap dibagikan berada di:
echo   dist\AbsensiCleaner
echo ================================================
pause
