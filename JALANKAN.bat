@echo off
title Absensi Cleaner
echo.
echo  ================================================
echo   ABSENSI CLEANER - Web App
echo  ================================================
echo.

REM Matikan server lama jika masih berjalan
echo  Memeriksa proses lama...
taskkill /F /IM python.exe /T >nul 2>&1
timeout /t 1 /nobreak >nul

echo  Memulai server versi terbaru...
echo  Silakan buka browser dan akses:
echo.
echo      http://localhost:5000
echo.
echo  Tekan Ctrl+C untuk menghentikan server.
echo  ================================================
REM Buka browser otomatis
start "" "http://localhost:5000"

REM Jalankan python secara dinamis (mencari venv lokal atau python sistem)

if exist "%~dp0venv\Scripts\python.exe" (
    "%~dp0venv\Scripts\python.exe" "%~dp0jalankan.py"
) else (
    python "%~dp0jalankan.py"
)
pause


