# -*- coding: utf-8 -*-
"""
Peluncur Universal Absensi Cleaner (Cross-Platform: Windows & macOS)
===================================================================
Dapat dijalankan di Windows maupun Mac dengan perintah:
    python jalankan.py  (atau python3 jalankan.py di Mac)
"""

import sys
import os
import time
import socket
import subprocess
import webbrowser
import threading

def check_dependencies():
    """Memeriksa apakah library yang dibutuhkan sudah terinstal."""
    required = {
        "flask": "Flask",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "xlrd": "xlrd"
    }
    missing = []
    for module_name, pip_name in required.items():
        try:
            __import__(module_name)
        except ImportError:
            missing.append(pip_name)
    
    if missing:
        print(f"[INFO] Modul belum terpasang: {', '.join(missing)}")
        print("[INFO] Menginstal modul yang diperlukan otomatis...")
        try:
            req_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")
            if os.path.exists(req_file):
                subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file])
            else:
                subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
            print("[OK] Instalasi dependensi selesai!\n")
        except Exception as e:
            print(f"[PERINGATAN] Gagal menginstal otomatis: {e}")
            print(f"Silakan jalankan manual: pip install {' '.join(missing)}\n")

def find_available_port(default_port=5000):
    """Mencari port yang tidak bentrok (berguna jika port 5000 dipakai AirPlay Mac)."""
    candidate_ports = [default_port, 5050, 8000, 8080]
    for p in candidate_ports:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            if s.connect_ex(('127.0.0.1', p)) != 0:
                return p
    return default_port

def open_browser_delayed(url, delay=1.5):
    """Membuka browser bawaan sistem (Chrome/Safari/Edge/dll) secara otomatis."""
    def _worker():
        time.sleep(delay)
        webbrowser.open(url)
    threading.Thread(target=_worker, daemon=True).start()

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    print("=" * 56)
    print("       ABSENSI CLEANER — PELUNCUR APLIKASI")
    print("=" * 56)
    print(f"Sistem Operasi : {sys.platform.upper()}")
    print(f"Python Binary  : {sys.executable}")
    print("-" * 56)

    # 1. Cek & install dependensi jika belum ada
    check_dependencies()

    # 2. Tentukan port yang bebas
    port = find_available_port(5000)
    url = f"http://localhost:{port}"

    print(f"\n[OK] Memulai server...")
    print(f"--> URL Aplikasi : {url}")
    if port != 5000:
        print(f"--> (Catatan: Port 5000 terpakai, otomatis beralih ke port {port})")
    print("--> Tekan Ctrl+C untuk menghentikan server kapan saja.\n")

    # 3. Buka browser otomatis
    open_browser_delayed(url)

    # 4. Set environment variable PORT dan jalankan app.py
    os.environ["PORT"] = str(port)
    import app
    # host 0.0.0.0 agar bisa diakses oleh komputer/laptop lain di jaringan kantor yang sama
    app.app.run(host="0.0.0.0", port=port, debug=False)

if __name__ == "__main__":
    main()
