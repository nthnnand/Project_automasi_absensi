@echo off
title Absensi Cleaner
cd /d "%~dp0"
echo Membuka Absensi Cleaner...
timeout /t 1 /nobreak >nul
start "" "AbsensiCleaner.exe"
