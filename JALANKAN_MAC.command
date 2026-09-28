#!/bin/bash
# ================================================================
# Launcher Absensi Cleaner untuk macOS
# Cukup klik ganda (double-click) file ini dari Finder untuk membuka aplikasi!
# ================================================================

cd "$(dirname "$0")"

echo "================================================"
echo "    ABSENSI CLEANER - macOS Launcher"
echo "================================================"
echo ""

# 1. Cari binary Python 3
if command -v python3 &>/dev/null; then
    PY_BIN="python3"
elif command -v python &>/dev/null; then
    PY_BIN="python"
else
    echo "[ERROR] Python 3 tidak ditemukan di Mac ini!"
    echo "Silakan unduh dan pasang Python 3 dari:"
    echo "https://www.python.org/downloads/"
    echo ""
    read -p "Tekan Enter untuk keluar..."
    exit 1
fi

# 2. Cek apakah ada virtual environment lokal
if [ -f "venv/bin/python" ]; then
    PY_BIN="venv/bin/python"
fi

# 3. Jalankan peluncur universal
$PY_BIN jalankan.py

# Jaga jendela terminal tetap terbuka jika aplikasi berhenti
echo ""
read -p "Aplikasi berhenti. Tekan Enter untuk menutup jendela ini..."
