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
echo.
start "" "http://localhost:5000"
C:\Users\Henri\AppData\Local\Programs\Python\Python313\python.exe "%~dp0app.py"
pause

