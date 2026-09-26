#!/bin/bash
# ================================================================
# Skrip Pembuat File Executable Absensi Cleaner untuk macOS
# ================================================================

cd "$(dirname "$0")"

echo "================================================"
echo "   MEMBUAT FILE STANDALONE ABSENSI CLEANER (MAC)"
echo "================================================"
echo ""

# 1. Cari binary Python 3
if command -v python3 &>/dev/null; then
    PY_BIN="python3"
elif command -v python &>/dev/null; then
    PY_BIN="python"
else
    echo "[ERROR] Python 3 tidak ditemukan!"
    read -p "Tekan Enter untuk keluar..."
    exit 1
fi

echo "[1/3] Menginstall library pendukung & PyInstaller..."
$PY_BIN -m pip install --upgrade pip
$PY_BIN -m pip install -r requirements.txt
$PY_BIN -m pip install pyinstaller

echo ""
echo "[2/3] Memproses build PyInstaller untuk macOS..."
# Catatan penting: di macOS, pemisah argumen --add-data menggunakan titik dua (:) bukan titik koma (;)
$PY_BIN -m PyInstaller --noconfirm --onedir --console \
  --name "AbsensiCleaner" \
  --add-data "templates:templates" \
  --add-data "static:static" \
  app.py

echo ""
echo "[3/3] Menyiapkan shortcut peluncur Mac..."
cat << 'EOF' > "dist/AbsensiCleaner/BUKA_APLIKASI_MAC.command"
#!/bin/bash
cd "$(dirname "$0")"
echo "Membuka Absensi Cleaner..."
sleep 2 && open "http://localhost:5000" &
./AbsensiCleaner
EOF
chmod +x "dist/AbsensiCleaner/BUKA_APLIKASI_MAC.command"
chmod +x "dist/AbsensiCleaner/AbsensiCleaner" 2>/dev/null

echo ""
echo "================================================"
echo "   SELESAI!"
echo "   Folder aplikasi Mac siap dibagikan berada di:"
echo "   dist/AbsensiCleaner"
echo "================================================"
echo ""
read -p "Tekan Enter untuk keluar..."
