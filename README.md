# Attendance Automation System

Sistem otomasi pengolahan data absensi berbasis lokal/offline menggunakan Python, Flask, Pandas, OpenPyXL, dan SQLite.

Sistem dibuat untuk mengotomatisasi proses pengolahan data absensi hasil export dari mesin absensi dan memasukkannya ke dalam template Excel resmi yang sudah digunakan.

---

# 1. Latar Belakang

Proses absensi saat ini masih dilakukan secara manual.

Mesin absensi menghasilkan file laporan absensi yang harus diekspor atau diunduh terlebih dahulu. Setelah itu, data absensi harus diperiksa dan dicocokkan dengan data pegawai, kemudian dimasukkan secara manual ke dalam template Excel resmi.

Proses tersebut memiliki beberapa kendala:

- Membutuhkan waktu yang cukup lama.
- Harus mencari nama pegawai satu per satu.
- Risiko salah memasukkan data pegawai.
- Risiko salah memasukkan waktu masuk atau pulang.
- Risiko duplikasi data.
- Risiko kesalahan copy-paste.
- Risiko kesalahan manusia dalam pengolahan data.
- Sulit melakukan pembaruan data secara berkala.
- Data raw dapat sulit ditelusuri jika proses dilakukan secara manual.

Oleh karena itu dibuat sistem otomasi yang dapat membaca file absensi mentah, membersihkan dan mengolah data, mengecek pegawai pada Master Data, menentukan apakah pegawai termasuk PNS atau P3K, kemudian memasukkan hasilnya ke template Excel resmi.

---

# 2. Tujuan Sistem

Tujuan sistem:

1. Mengurangi proses copy-paste manual.
2. Mengotomatisasi pengolahan data absensi.
3. Mengurangi human error.
4. Mengurangi risiko duplikasi data.
5. Mengotomatisasi pemilihan waktu masuk.
6. Mengotomatisasi pemilihan waktu pulang.
7. Memvalidasi jam masuk pegawai.
8. Mengecek pegawai berdasarkan Master Data.
9. Memisahkan data PNS dan P3K.
10. Mempertahankan template Excel resmi.
11. Mengisi sheet `Input Data Absen` secara otomatis.
12. Memungkinkan pembaruan Draft setiap minggu.
13. Menghasilkan Draft PNS dan Draft P3K/PPPK.
14. Memungkinkan Draft diekspor kembali menjadi file Excel.

---

# 3. Konsep Utama Sistem

Sistem TIDAK membuat template Excel baru dari nol.

Sistem menggunakan template Excel resmi yang sudah tersedia.

Template resmi digunakan sebagai dasar Draft.

Alur utama:

Raw Attendance
        |
        v
Cleaning & Validation
        |
        v
Attendance Processing
        |
        v
Master Data Checking
        |
        v
PNS / P3K Classification
        |
        v
Official Excel Template
        |
        v
Sheet "Input Data Absen"
        |
        v
Draft PNS / Draft P3K
        |
        v
Weekly Update
        |
        v
Export Excel

---

# 4. Scope Sistem

## 4.1 Termasuk Dalam Scope

Sistem menangani:

- Membuat project bulanan.
- Upload Master Data.
- Upload Template PNS.
- Upload Template P3K/PPPK.
- Upload Raw Attendance.
- Membaca file Excel.
- Membersihkan data.
- Menghapus exact duplicate.
- Memproses C/Masuk.
- Memproses C/Keluar.
- Validasi jam masuk.
- Mengecek nama pada Master Data.
- Mengecek sheet PNS.
- Mengecek sheet P3K.
- Menentukan kategori pegawai.
- Menentukan Draft tujuan.
- Mencari pegawai pada Draft.
- Menambahkan pegawai jika belum ada.
- Mapping tanggal absensi ke template.
- Mengisi data ke sheet `Input Data Absen`.
- Menampilkan warning/error.
- Preview hasil.
- Menyimpan Draft.
- Melakukan update Draft mingguan.
- Export Draft ke Excel.

---

## 4.2 Tidak Termasuk Dalam Scope

Sistem tidak menangani:

- Integrasi langsung dengan mesin fingerprint.
- Integrasi langsung dengan mesin face attendance.
- Mengambil data langsung dari perangkat absensi.
- Membuat template resmi baru dari nol.
- Mengubah raw attendance asli.
- Menghapus raw attendance asli.
- Mengubah Master Data secara otomatis.
- Mengarang data absensi.
- Mengubah struktur resmi template tanpa kebutuhan.
- Menentukan aturan Keterangan yang belum diberikan secara resmi.

---

# 5. Teknologi

## Backend

- Python
- Flask

## Data Processing

- Pandas
- OpenPyXL

## Database

- SQLite

## Frontend

- HTML
- CSS
- JavaScript

## Environment

- Localhost
- Offline / Local Network

---

# 6. Input Sistem

Sistem memiliki tiga jenis input utama:

1. Master Data
2. Raw Attendance
3. Official Excel Template

---

# 7. Master Data

Master Data berupa satu file Excel yang memiliki dua sheet utama:

```text
Pegawai JATIM.xlsx

├── PNS
└── P3K
