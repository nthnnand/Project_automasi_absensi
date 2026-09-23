# -*- coding: utf-8 -*-
"""
Absensi Cleaner v2 — Flask Web App
====================================
Otomatis: baca fingerprint/WFH → cleaning → cocokkan nama → isi template PNS/PPPK

Fixes:
  Case 1 — Prioritas Status Baru (Kolom F): F > E, Lembur Masuk/Keluar dikenali
  Case 2 — Preserve Data Validation dropdown (keep_vba untuk .xlsm)
  Case 3 — Fuzzy matching ketat: threshold 80, deteksi multi-kandidat ambiguitas
  Case 4 — Warna merah pada sel jam kosong (data tidak tersedia)
  Case 5 — WFH data processing dari Google Forms Excel + merge dengan WFO
"""

import os, io, re, zipfile, copy, shutil, difflib
import pandas as pd
from datetime import time as dt_time
from flask import Flask, render_template, request, send_file, jsonify, session
import xlrd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "absensi_cleaner_v2_2026"
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Auto-match threshold: skor >= ini → rekomendasi sangat cocok
AUTO_MATCH_THRESHOLD = 80

# Pilihan dropdown untuk kolom K (Keterangan) di sheet Input Data Absen
DROPDOWN_KETERANGAN_FORMULA = '"H,WFH,S,DL,DK,TK,I,IT"'


# ══════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════

def norm_id(raw):
    """Normalize fingerprint ID (remove .0 from Excel float)"""
    s = str(raw).strip()
    try:
        if "." in s and float(s) == int(float(s)):
            return str(int(float(s)))
    except Exception:
        pass
    return s


# Singkatan gelar akademik Indonesia yang sering muncul di data WFH (Google Forms)
# Difilter agar tidak mengganggu skor matching nama
GELAR_WORDS = {
    # Sarjana (S1)
    "SH", "SE", "SI", "SM", "SA", "SP", "ST",
    "PD", "AG", "IP", "HI", "HUM", "AB",
    "KOM", "PSI", "MAT", "SOS", "KED", "KEP",
    "STAT", "IKOM", "IAN", "STP", "SST",
    # Master (S2)
    "MM", "MH", "MA", "MT", "MSI", "MSC",
    # Diploma
    "MD", "AK",
    # Doktor
    "DR",
    # Gelar umum
    "TR", "STR",
    # Gelar profesi
    "AP", "SAP",
}


def norm_name(name):
    """Normalize name for matching: uppercase, strip punctuation, collapse spaces."""
    s = str(name).upper().strip()
    s = re.sub(r'[.,\-\'\"()]+', ' ', s)   # remove punctuation
    s = re.sub(r'\s+', ' ', s).strip()      # collapse spaces
    return s


def name_words(name):
    """
    Get list of words from normalized name (min 2 chars).
    Filter out common academic title abbreviations (gelar)
    agar tidak mengganggu skor fuzzy matching.
    """
    return [
        w for w in norm_name(name).split()
        if len(w) >= 2 and w not in GELAR_WORDS
    ]


def match_score(fp_name, tpl_name):
    """
    Score how well fingerprint name matches template name.
    Returns (score 0-100, match_type string).

    Levels:
      1. Exact match          → 100
      2. Word subset/superset → 60–90
      3. Partial word overlap → 40–80 (min 3 chars to support common Indonesian names)
      4. Fuzzy word matching  → 35–75 (misal: mustikah vs mustika)
      5. Character similarity → 30–88 (typo seperti Rijono vs Riyono)
    """
    n1 = norm_name(fp_name)
    n2 = norm_name(tpl_name)

    # 1. Exact match
    if n1 == n2:
        return 100, "exact"

    w1 = name_words(fp_name)
    w2 = name_words(tpl_name)

    if not w1 or not w2:
        return 0, "none"

    sw1, sw2 = set(w1), set(w2)

    # 2. All words of shorter name found in longer name
    if sw1.issubset(sw2):
        pct = len(sw1) / max(len(sw2), 1) * 90
        return max(int(pct), 60), "subset"
    if sw2.issubset(sw1):
        pct = len(sw2) / max(len(sw1), 1) * 90
        return max(int(pct), 60), "superset"

    # 3. Overlap: shared words / total words
    shared = sw1 & sw2
    if shared:
        # Require at least one word with 3+ characters to avoid false positives
        long_shared = [w for w in shared if len(w) >= 3]
        if long_shared:
            overlap = len(shared) / max(len(sw1 | sw2), 1)
            return max(int(overlap * 80), 40), "partial"

    # 4. Fuzzy word-level matching (e.g. mustikah vs mustika)
    matched_words = 0
    for word1 in sw1:
        if len(word1) < 3:
            continue
        for word2 in sw2:
            if len(word2) < 3:
                continue
            r = difflib.SequenceMatcher(None, word1, word2).ratio()
            if r >= 0.80:
                matched_words += 1
                break
    if matched_words > 0:
        overlap = matched_words / max(len(sw1), len(sw2))
        return max(int(overlap * 75), 35), "fuzzy_words"

    # 5. Character-level similarity (untuk typo: Rijono vs Riyono, Mustikah vs Mustika)
    ratio = difflib.SequenceMatcher(None, n1, n2).ratio()
    if ratio >= 0.78:
        return max(int(ratio * 88), 68), "typo"
    elif ratio >= 0.50:
        return max(int(ratio * 60), 30), "similarity"

    return 0, "none"


def find_best_match(tpl_name, fp_lookup):
    """
    Find best matching fingerprint ID for a template name.
    fp_lookup = {fp_id: fp_name}

    Returns (fp_id, score, match_type, needs_review).
      needs_review=False → satu kandidat terpercaya >= AUTO_MATCH_THRESHOLD (65), auto-match
      needs_review=True  → ambiguitas atau skor di bawah threshold (tapi ada kandidat > 0),
                           ditampilkan sebagai 'Perlu Review' dengan tombol Masukkan / Abaikan
    """
    candidates = []  # list of (score, mtype, fp_id)

    for fp_id, fp_name in fp_lookup.items():
        score, mtype = match_score(fp_name, tpl_name)
        if score > 0:
            candidates.append((score, mtype, fp_id))

    if not candidates:
        return None, 0, "none", False

    # Urutkan descending berdasarkan skor
    candidates.sort(key=lambda x: -x[0])
    best_score, best_type, best_id = candidates[0]

    # Kandidat dengan skor >= threshold
    high_candidates = [c for c in candidates if c[0] >= AUTO_MATCH_THRESHOLD]

    if len(high_candidates) == 1:
        # Satu kandidat terpercaya >= threshold → auto-match
        return best_id, best_score, best_type, False
    elif len(high_candidates) > 1:
        # Ambiguitas: beberapa kandidat skornya tinggi → perlu review
        return best_id, best_score, best_type, True
    else:
        # Ada kandidat terbaik (skor < threshold) → Perlu Review (tidak langsung diabaikan!)
        return best_id, best_score, best_type, True


# ══════════════════════════════════════════════════════════════
#  FINGERPRINT DATA: READ & CLEAN
# ══════════════════════════════════════════════════════════════

def read_fingerprint(filepath):
    """
    Read fingerprint XLS/XLSX file. Returns list of dicts.
    Membaca kolom F (index 5) sebagai 'status_baru' (Case 1).
    """
    ext = os.path.splitext(filepath)[1].lower()
    rows = []

    if ext == ".xls":
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        for i in range(1, ws.nrows):
            # Baca status_baru dari kolom F (index 5) jika ada
            status_baru = ""
            if ws.ncols > 5:
                v = ws.cell(i, 5).value
                status_baru = str(v).strip() if v else ""

            rows.append({
                "no_id":       norm_id(str(ws.cell(i, 0).value)),
                "nama":        str(ws.cell(i, 2).value).strip(),
                "waktu":       str(ws.cell(i, 3).value).strip(),
                "status":      str(ws.cell(i, 4).value).strip(),
                "status_baru": status_baru,
            })
    else:
        wb = load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Baca status_baru dari kolom F (index 5) jika kolom tersedia
            status_baru = ""
            if len(row) > 5 and row[5] is not None:
                status_baru = str(row[5]).strip()

            rows.append({
                "no_id":       norm_id(str(row[0]) if row[0] is not None else ""),
                "nama":        str(row[2]).strip() if row[2] else "",
                "waktu":       str(row[3]).strip() if row[3] else "",
                "status":      str(row[4]).strip() if row[4] else "",
                "status_baru": status_baru,
            })
        wb.close()

    return rows


def extract_datetime(waktu_str):
    """Extract (day_number, time_str) from 'dd/mm/yyyy HH:MM'."""
    try:
        parts = waktu_str.strip().split(" ")
        date_parts = parts[0].split("/")
        day = int(date_parts[0])
        time_parts = parts[1].split(":")
        time_str = f"{int(time_parts[0]):02d}:{int(time_parts[1]):02d}"
        return day, time_str
    except Exception:
        return None, None


def time_earlier(t1, t2):
    """True if t1 is earlier than t2 (both 'HH:MM' strings)."""
    try:
        h1, m1 = map(int, t1.split(":"))
        h2, m2 = map(int, t2.split(":"))
        return h1 * 60 + m1 < h2 * 60 + m2
    except Exception:
        return False


def clean_fingerprint(rows, start_day=1, end_day=31):
    """
    Clean fingerprint data. Returns:
      fp_names:  {fp_id: canonical_name}
      fp_data:   {fp_id: {day: {"masuk": "HH:MM", "keluar": "HH:MM"}}}

    Case 1 — Prioritas Status Baru:
      Kolom F (status_baru) diutamakan jika terisi.
      Jika kolom F kosong, gunakan Kolom E (status).
      "Lembur Masuk" dihitung sebagai Masuk.
      "Lembur Keluar" dihitung sebagai Keluar.
    """
    # Status yang dikenali sebagai masuk / keluar (case-insensitive)
    STATUS_MASUK_LOWER  = {"c/masuk",  "lembur masuk"}
    STATUS_KELUAR_LOWER = {"c/keluar", "lembur keluar"}

    fp_names = {}
    fp_data = {}

    for r in rows:
        no_id       = r["no_id"]
        nama        = r["nama"]
        waktu       = r["waktu"]
        status      = r["status"]
        status_baru = r.get("status_baru", "")

        # 1. Skip baris yang tidak lengkap
        if not no_id or not nama or not waktu:
            continue

        # Case 1: Resolusi status efektif — Kolom F (Status Baru) prioritas utama
        if status_baru and status_baru.lower() not in ("", "none", "-"):
            effective_status = status_baru
        else:
            effective_status = status

        # Tentukan masuk/keluar dari status efektif (case-insensitive)
        eff_lower = effective_status.lower().strip()
        if eff_lower in STATUS_MASUK_LOWER:
            key = "masuk"
        elif eff_lower in STATUS_KELUAR_LOWER:
            key = "keluar"
        else:
            continue  # status tidak dikenali, skip

        # 2–3. Ekstrak tanggal & waktu
        day, time_str = extract_datetime(waktu)
        if day is None or time_str is None:
            continue

        # Filter tanggal sesuai rentang yang dipilih
        if day < start_day or day > end_day:
            continue

        # Simpan nama kanonik per ID
        if no_id not in fp_names:
            fp_names[no_id] = nama

        # Init nested dict
        if no_id not in fp_data:
            fp_data[no_id] = {}
        if day not in fp_data[no_id]:
            fp_data[no_id][day] = {}

        # 4–7. Dedup + simpan waktu paling awal
        if key not in fp_data[no_id][day]:
            fp_data[no_id][day][key] = time_str
        else:
            if time_earlier(time_str, fp_data[no_id][day][key]):
                fp_data[no_id][day][key] = time_str

    return fp_names, fp_data


# ══════════════════════════════════════════════════════════════
#  WFH DATA: READ & CLEAN (Google Forms Excel)
# ══════════════════════════════════════════════════════════════

def read_wfh_data(filepath, start_day=1, end_day=31):
    """
    Baca dan proses data WFH dari Excel (Google Forms export).

    Kolom yang diharapkan (case-insensitive):
    - Timestamp       → dipecah menjadi tanggal (day int) + jam (HH:MM)
    - NAMA LENGKAP    → identitas pegawai
    - KEHADIRAN       → 'CHECK IN' atau 'CHECK OUT'

    Returns: (wfh_names, wfh_data) format identik clean_fingerprint():
      wfh_names = {wfh_id: nama}
      wfh_data  = {wfh_id: {day: {"masuk": "HH:MM", "keluar": "HH:MM"}}}
    """
    df = pd.read_excel(filepath)

    # Normalisasi nama kolom: lowercase, strip whitespace
    df.columns = df.columns.str.strip().str.lower()

    # Cari kolom yang cocok (flexible matching)
    col_timestamp = None
    col_nama = None
    col_kehadiran = None

    for col in df.columns:
        cl = col.lower().strip()
        if 'timestamp' in cl or 'waktu' in cl:
            col_timestamp = col
        elif 'nama' in cl and ('lengkap' in cl or col == 'nama'):
            col_nama = col
        elif 'kehadiran' in cl or 'status' in cl or 'check' in cl:
            col_kehadiran = col

    # Fallback: jika tidak ditemukan, coba pakai kolom pertama, kedua, ketiga
    if col_timestamp is None and len(df.columns) >= 1:
        col_timestamp = df.columns[0]
    if col_nama is None and len(df.columns) >= 2:
        col_nama = df.columns[1]
    if col_kehadiran is None and len(df.columns) >= 3:
        col_kehadiran = df.columns[2]

    if col_timestamp is None or col_nama is None or col_kehadiran is None:
        raise ValueError(
            "File WFH tidak memiliki kolom yang diperlukan "
            "(Timestamp, NAMA LENGKAP, KEHADIRAN)"
        )

    # Drop baris kosong
    df = df.dropna(subset=[col_timestamp, col_nama, col_kehadiran])

    # Parse timestamp
    df[col_timestamp] = pd.to_datetime(df[col_timestamp], dayfirst=True, errors='coerce')
    df = df.dropna(subset=[col_timestamp])

    # Ekstrak day (int) dan jam (HH:MM)
    df['_day'] = df[col_timestamp].dt.day
    df['_time'] = df[col_timestamp].dt.strftime('%H:%M')

    # Filter tanggal sesuai rentang yang dipilih
    df = df[(df['_day'] >= start_day) & (df['_day'] <= end_day)]

    # Normalisasi nama dan kehadiran
    df['_nama'] = df[col_nama].astype(str).str.strip()
    df['_kehadiran'] = df[col_kehadiran].astype(str).str.strip().str.upper()

    wfh_names = {}  # {wfh_id: canonical_name}
    wfh_data = {}   # {wfh_id: {day: {"masuk": "HH:MM", "keluar": "HH:MM"}}}

    # Group by nama + day
    for (nama, day), group in df.groupby(['_nama', '_day']):
        if not nama or nama.lower() in ('nan', 'none', ''):
            continue

        # Generate WFH ID unik berdasarkan nama (prefix WFH_ agar tidak bentrok fp_id)
        wfh_id = "WFH_" + norm_name(nama)

        if wfh_id not in wfh_names:
            wfh_names[wfh_id] = nama

        if wfh_id not in wfh_data:
            wfh_data[wfh_id] = {}

        day_int = int(day)
        if day_int not in wfh_data[wfh_id]:
            wfh_data[wfh_id][day_int] = {}

        # Filter CHECK IN → jam paling awal
        checkins = group[group['_kehadiran'] == 'CHECK IN']
        if not checkins.empty:
            earliest = checkins['_time'].min()
            if 'masuk' not in wfh_data[wfh_id][day_int]:
                wfh_data[wfh_id][day_int]['masuk'] = earliest
            elif time_earlier(earliest, wfh_data[wfh_id][day_int]['masuk']):
                wfh_data[wfh_id][day_int]['masuk'] = earliest

        # Filter CHECK OUT → jam paling akhir
        checkouts = group[group['_kehadiran'] == 'CHECK OUT']
        if not checkouts.empty:
            latest = checkouts['_time'].max()
            if 'keluar' not in wfh_data[wfh_id][day_int]:
                wfh_data[wfh_id][day_int]['keluar'] = latest
            elif not time_earlier(latest, wfh_data[wfh_id][day_int]['keluar']):
                wfh_data[wfh_id][day_int]['keluar'] = latest

    return wfh_names, wfh_data


def merge_attendance_data(fp_names, fp_data, wfh_names, wfh_data):
    """
    Gabungkan data WFO (fingerprint) dan WFH menjadi satu dict.

    Strategi: cross-reference nama WFH ke fp_id fingerprint yang sesuai.
    - Jika nama WFH cocok dengan nama fingerprint → gabungkan hari WFH ke fp_id tsb
    - Jika tidak cocok (pegawai WFH-only) → tambahkan sebagai WFH_NAMA (fallback)

    Jika satu orang tercatat WFO dan WFH di hari yang sama,
    data WFO (fingerprint) diutamakan — tidak di-overwrite oleh WFH.

    Returns: (merged_names, merged_data)
    """
    merged_names = dict(fp_names)    # mulai dari copy WFO
    merged_data = {}

    # Deep-copy fp_data
    for fp_id, days in fp_data.items():
        merged_data[fp_id] = {}
        for day, times in days.items():
            merged_data[fp_id][day] = dict(times)

    # Cross-reference: cari fp_id yang cocok untuk setiap pegawai WFH
    # Bangun lookup {wfh_id: matched_fp_id atau None}
    wfh_to_fp = {}
    for wfh_id, wfh_name in wfh_names.items():
        best_fp_id = None
        best_score = 0
        for fp_id, fp_name in fp_names.items():
            score, _ = match_score(wfh_name, fp_name)
            if score > best_score:
                best_score = score
                best_fp_id = fp_id
        # Gunakan threshold yang sama (AUTO_MATCH_THRESHOLD) untuk cross-reference
        if best_score >= 0 and best_fp_id:
            wfh_to_fp[wfh_id] = best_fp_id
        else:
            wfh_to_fp[wfh_id] = None  # WFH-only, tidak ada pasangan fingerprint

    # Merge WFH data berdasarkan cross-reference
    for wfh_id, days in wfh_data.items():
        target_id = wfh_to_fp.get(wfh_id)

        if target_id:
            # Cocok dengan fp_id → gabungkan hari WFH ke entry fingerprint yang sama
            if target_id not in merged_data:
                merged_data[target_id] = {}
            for day, times in days.items():
                if day not in merged_data[target_id]:
                    # Hari ini hanya ada WFH → tambahkan
                    merged_data[target_id][day] = dict(times)
                # else: hari sudah ada data WFO → WFO diutamakan, skip WFH
        else:
            # WFH-only (tidak ada fingerprint) → tambahkan sebagai WFH_NAMA
            merged_names[wfh_id] = wfh_names[wfh_id]
            if wfh_id not in merged_data:
                merged_data[wfh_id] = {}
            for day, times in days.items():
                if day not in merged_data[wfh_id]:
                    merged_data[wfh_id][day] = dict(times)

    return merged_names, merged_data


# ══════════════════════════════════════════════════════════════
#  TEMPLATE: READ EMPLOYEE LIST
# ══════════════════════════════════════════════════════════════

def read_template_employees(filepath):
    """
    Read employee names from Input Data Absen sheet.
    Returns list of {row: int, name: str} (row = Excel row number).
    Dilengkapi fallback ke sheet 'Database' jika cache formula di Input Data Absen kosong.
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)

    ws = wb['Input Data Absen']
    employees = []
    for r in range(6, ws.max_row + 1):
        no = ws.cell(r, 1).value
        nama = ws.cell(r, 2).value
        if no is not None and nama and str(nama).strip() and not str(nama).startswith("="):
            employees.append({
                "row": r,
                "no": no,
                "name": str(nama).strip(),
            })

    # Fallback ke sheet 'Database' jika formula belum dievaluasi cache-nya
    if not employees and 'Database' in wb.sheetnames:
        ws_db = wb['Database']
        for r_db in range(5, ws_db.max_row + 1):
            no = ws_db.cell(r_db, 1).value
            nama = ws_db.cell(r_db, 2).value
            if no is not None and nama and str(nama).strip():
                employees.append({
                    "row": r_db + 1,
                    "no": no,
                    "name": str(nama).strip(),
                })

    wb.close()
    return employees


# ══════════════════════════════════════════════════════════════
#  TEMPLATE: FILL WITH CLEANED DATA
# ══════════════════════════════════════════════════════════════

def time_str_to_dt(time_str):
    """Convert 'HH:MM' to datetime.time for Excel."""
    try:
        h, m = map(int, time_str.split(":"))
        return dt_time(h, m, 0)
    except Exception:
        return None


def is_empty_time_cell(value):
    """
    Cek apakah sel jam dianggap kosong/tidak terisi.
    Menangkap: None, 0, 0.0, string kosong, time(0,0,0) = 00:00
    """
    if value is None:
        return True
    if value == 0 or value == 0.0:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    if isinstance(value, dt_time) and value == dt_time(0, 0, 0):
        return True
    return False


def apply_keterangan_dropdown(ws, employees=None):
    """
    Terapkan Data Validation (dropdown list) pada kolom K (Keterangan) untuk seluruh hari 1 s/d 31.
    Pilihan dropdown: H, WFH, S, DL, DK, TK, I, IT.
    """
    if employees:
        min_row = min(e["row"] for e in employees)
        max_row = max(e["row"] for e in employees)
    else:
        emp_rows = []
        for r in range(6, ws.max_row + 1):
            no = ws.cell(r, 1).value
            nama = ws.cell(r, 2).value
            if no is not None and nama and str(nama).strip():
                emp_rows.append(r)
        if emp_rows:
            min_row, max_row = min(emp_rows), max(emp_rows)
        else:
            min_row, max_row = 6, max(ws.max_row, 31)

    dv = DataValidation(
        type="list",
        formula1=DROPDOWN_KETERANGAN_FORMULA,
        allow_blank=True
    )
    ws.add_data_validation(dv)
    for day in range(1, 32):
        col_k = 3 * day + 2
        col_letter = get_column_letter(col_k)
        dv.add(f"{col_letter}{min_row}:{col_letter}{max_row}")
    return dv


def fill_template(src_path, out_path, employees, fp_names, fp_data, wfh_names, wfh_data, decisions=None, start_day=1, end_day=31):
    """
    Copy template, fill K/M/P columns in 'Input Data Absen'.
    - Kolom K (Keterangan): diisi 'H' untuk WFO, 'WFH' untuk WFH sesuai sumber data yang dipilih.
    - Kolom M (Masuk) & P (Pulang): diisi jam absen.
    - Keputusan manual penuh: data baru masuk jika dipilih secara eksplisit oleh user.
    - Hanya tanggal dalam rentang start_day s/d end_day yang diproses.
    """
    shutil.copy2(src_path, out_path)
    ext = os.path.splitext(src_path)[1].lower()
    wb = load_workbook(out_path, keep_vba=(ext == ".xlsm"))
    ws = wb['Input Data Absen']
    red_fill = PatternFill(start_color="FFC0392B", end_color="FFC0392B", fill_type="solid")

    decisions = decisions or {}
    employee_rows = set()
    filled_cells = set()

    for emp in employees:
        tpl_row = emp["row"]
        tpl_name = emp["name"]
        employee_rows.add(tpl_row)

        user_choice = (
            decisions.get(str(tpl_row)) or 
            decisions.get(tpl_row) or 
            decisions.get(tpl_name) or
            decisions.get(norm_name(tpl_name))
        )

        use_wfo = False
        use_wfh = False

        if isinstance(user_choice, dict):
            use_wfo = bool(user_choice.get("wfo"))
            use_wfh = bool(user_choice.get("wfh"))
        elif isinstance(user_choice, str):
            c_low = user_choice.lower().strip()
            if c_low in ("wfo", "fingerprint"):
                use_wfo = True
            elif c_low in ("wfh",):
                use_wfh = True
            elif c_low in ("both", "include", "masukkan"):
                use_wfo = True
                use_wfh = True
            elif c_low in ("ignore", "abaikan"):
                use_wfo = False
                use_wfh = False

        # ATURAN UTAMA: Tidak ada auto-fill! Jika user tidak memilih / abaikan, jangan isi data!
        if not (use_wfo or use_wfh):
            continue

        # 1. Tulis data WFO jika user memilih WFO
        if use_wfo and fp_names:
            f_id, _, _, _ = find_best_match(tpl_name, fp_names)
            if f_id and f_id in fp_data:
                for day, times in fp_data[f_id].items():
                    if start_day <= day <= end_day:
                        col_k = 3 * day + 2
                        col_m = 3 * day + 3
                        col_p = 3 * day + 4
                        wfo_has_data = False
                        if "masuk" in times:
                            t = time_str_to_dt(times["masuk"])
                            if t:
                                ws.cell(row=tpl_row, column=col_m, value=t)
                                filled_cells.add((tpl_row, col_m))
                                wfo_has_data = True
                        if "keluar" in times:
                            t = time_str_to_dt(times["keluar"])
                            if t:
                                ws.cell(row=tpl_row, column=col_p, value=t)
                                filled_cells.add((tpl_row, col_p))
                                wfo_has_data = True
                        if wfo_has_data:
                            ws.cell(row=tpl_row, column=col_k, value="H")
                            filled_cells.add((tpl_row, col_k))

        # 2. Tulis data WFH jika user memilih WFH
        if use_wfh and wfh_names:
            w_id, _, _, _ = find_best_match(tpl_name, wfh_names)
            if w_id and w_id in wfh_data:
                for day, times in wfh_data[w_id].items():
                    if start_day <= day <= end_day:
                        col_k = 3 * day + 2
                        col_m = 3 * day + 3
                        col_p = 3 * day + 4
                        wfh_has_data = False
                        # Jika hanya WFH yang dipilih, atau sel belum diisi oleh WFO
                        if (tpl_row, col_m) not in filled_cells or not use_wfo:
                            if "masuk" in times:
                                t = time_str_to_dt(times["masuk"])
                                if t:
                                    ws.cell(row=tpl_row, column=col_m, value=t)
                                    filled_cells.add((tpl_row, col_m))
                                    wfh_has_data = True
                        if (tpl_row, col_p) not in filled_cells or not use_wfo:
                            if "keluar" in times:
                                t = time_str_to_dt(times["keluar"])
                                if t:
                                    ws.cell(row=tpl_row, column=col_p, value=t)
                                    filled_cells.add((tpl_row, col_p))
                                    wfh_has_data = True
                        if wfh_has_data and ((tpl_row, col_k) not in filled_cells or not use_wfo):
                            ws.cell(row=tpl_row, column=col_k, value="WFH")
                            filled_cells.add((tpl_row, col_k))

    # Pass kedua: Beri warna merah HANYA untuk hari dalam rentang [start_day, end_day]
    # pada sel jam yang kosong / belum terisi
    for tpl_row in employee_rows:
        for day in range(start_day, end_day + 1):
            if 1 <= day <= 31:
                col_m = 3 * day + 3
                col_p = 3 * day + 4
                for col in (col_m, col_p):
                    cell = ws.cell(row=tpl_row, column=col)
                    if is_empty_time_cell(cell.value):
                        cell.fill = red_fill

    # Pass ketiga: Terapkan Data Validation dropdown pada kolom K (Keterangan) untuk seluruh hari (1 s/d 31)
    apply_keterangan_dropdown(ws, employees)

    wb.save(out_path)
    wb.close()
    return len(filled_cells)


# ══════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/debug-matching", methods=["POST"])
def debug_matching():
    """
    Debug endpoint: tampilkan detail parsing fingerprint/WFH + skor matching.
    Berguna untuk verifikasi Case 1 (Status Baru) dan Case 3 (fuzzy match).
    """
    fp_path   = session.get("fp_path")
    wfh_path  = session.get("wfh_path")
    pns_path  = session.get("tpl_pns_path")
    pppk_path = session.get("tpl_pppk_path")

    has_fp  = fp_path and os.path.exists(fp_path)
    has_wfh = wfh_path and os.path.exists(wfh_path)

    if not has_fp and not has_wfh:
        return jsonify({"error": "Upload minimal 1 file absensi (Fingerprint atau WFH)"}), 400

    try:
        # Parse data WFO (fingerprint) jika ada
        fp_names, fp_data = {}, {}
        raw_sample = []
        if has_fp:
            rows = read_fingerprint(fp_path)
            fp_names, fp_data = clean_fingerprint(rows)

            # Sample raw rows (max 30) untuk cek Case 1
            for r in rows[:50]:
                if r.get("status_baru"):
                    raw_sample.append({
                        "nama":        r["nama"],
                        "waktu":       r["waktu"],
                        "status_e":    r["status"],
                        "status_baru": r["status_baru"],
                    })

        # Parse data WFH jika ada
        wfh_names, wfh_data = {}, {}
        if has_wfh:
            wfh_names, wfh_data = read_wfh_data(wfh_path)

        # Merge WFO + WFH
        if has_fp and has_wfh:
            merged_names, merged_data = merge_attendance_data(
                fp_names, fp_data, wfh_names, wfh_data
            )
        elif has_fp:
            merged_names, merged_data = fp_names, fp_data
        else:
            merged_names, merged_data = wfh_names, wfh_data

        # Detail matching untuk setiap template
        debug_matches = {}
        for tpl_type, tpl_path in [("pns", pns_path), ("pppk", pppk_path)]:
            if not tpl_path or not os.path.exists(tpl_path):
                continue
            employees = read_template_employees(tpl_path)
            matches = []
            for emp in employees:
                fp_id, score, mtype, needs_review = find_best_match(emp["name"], merged_names)
                matches.append({
                    "tpl_name":    emp["name"],
                    "fp_name":     merged_names.get(fp_id, "") if fp_id else "",
                    "score":       score,
                    "match_type":  mtype,
                    "needs_review": needs_review,
                    "will_fill":   bool(fp_id and score >= AUTO_MATCH_THRESHOLD and not needs_review),
                })
            debug_matches[tpl_type] = matches

        return jsonify({
            "threshold":          AUTO_MATCH_THRESHOLD,
            "fp_employees":       len(fp_names),
            "wfh_employees":      len(wfh_names),
            "merged_employees":   len(merged_names),
            "fp_names":           fp_names,
            "wfh_names":          wfh_names,
            "status_baru_sample": raw_sample,
            "matching":           debug_matches,
        })
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/upload-fingerprint", methods=["POST"])
def upload_fingerprint():
    """Upload fingerprint file, parse & clean."""
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Nama file kosong"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xls", ".xlsx", ".xlsm"):
        return jsonify({"error": "Format harus .xls atau .xlsx"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, "fingerprint" + ext)
    f.save(save_path)
    session["fp_path"] = save_path
    session["fp_filename"] = f.filename

    try:
        rows = read_fingerprint(save_path)
        fp_names, fp_data = clean_fingerprint(rows)

        # Stats: hitung total record bersih
        total_records = sum(
            len(times)
            for days in fp_data.values()
            for times in days.values()
        )

        session["fp_employee_count"] = len(fp_names)

        return jsonify({
            "success": True,
            "filename": f.filename,
            "raw_rows": len(rows),
            "employees": len(fp_names),
            "clean_records": total_records,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/upload-wfh", methods=["POST"])
def upload_wfh():
    """Upload dan parse file WFH Excel (Google Forms export)."""
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Nama file kosong"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xls", ".xlsx", ".xlsm"):
        return jsonify({"error": "Format harus .xls atau .xlsx"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, "wfh_data" + ext)
    f.save(save_path)
    session["wfh_path"] = save_path
    session["wfh_filename"] = f.filename

    try:
        wfh_names, wfh_data = read_wfh_data(save_path)

        # Stats
        total_records = sum(
            len(times)
            for days in wfh_data.values()
            for times in days.values()
        )

        return jsonify({
            "success": True,
            "filename": f.filename,
            "employees": len(wfh_names),
            "clean_records": total_records,
        })
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/upload-template", methods=["POST"])
def upload_template():
    """Upload PNS or PPPK template file."""
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400

    tpl_type = request.form.get("type", "").upper()  # "PNS" or "PPPK"
    if tpl_type not in ("PNS", "PPPK"):
        return jsonify({"error": "Tipe template harus PNS atau PPPK"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Nama file kosong"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm"):
        return jsonify({"error": "Template harus format .xlsx"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, f"template_{tpl_type.lower()}{ext}")
    f.save(save_path)
    session[f"tpl_{tpl_type.lower()}_path"] = save_path
    session[f"tpl_{tpl_type.lower()}_filename"] = f.filename

    try:
        employees = read_template_employees(save_path)
        return jsonify({
            "success": True,
            "type": tpl_type,
            "filename": f.filename,
            "employees": len(employees),
            "names": [e["name"] for e in employees],
        })
    except Exception as e:
        return jsonify({"error": f"Error membaca template: {str(e)}"}), 500


@app.route("/preview", methods=["POST"])
def preview():
    """Preview matching results before processing."""
    fp_path   = session.get("fp_path")
    wfh_path  = session.get("wfh_path")
    pns_path  = session.get("tpl_pns_path")
    pppk_path = session.get("tpl_pppk_path")

    has_fp  = fp_path and os.path.exists(fp_path)
    has_wfh = wfh_path and os.path.exists(wfh_path)

    if not has_fp and not has_wfh:
        return jsonify({"error": "Upload minimal 1 file absensi (Fingerprint atau WFH)"}), 400
    if not pns_path and not pppk_path:
        return jsonify({"error": "Minimal upload 1 template (PNS atau PPPK)"}), 400

    data = request.get_json(silent=True) or {}
    try:
        start_day = int(data.get("start_day") or session.get("start_day") or 1)
        end_day = int(data.get("end_day") or session.get("end_day") or 31)
    except Exception:
        start_day, end_day = 1, 31

    if start_day > end_day:
        start_day, end_day = end_day, start_day

    session["start_day"] = start_day
    session["end_day"] = end_day

    try:
        # 1. Parse data WFO (fingerprint) jika ada dengan filter rentang tanggal
        fp_names, fp_data = {}, {}
        if has_fp:
            rows = read_fingerprint(fp_path)
            fp_names, fp_data = clean_fingerprint(rows, start_day=start_day, end_day=end_day)

        # 2. Parse data WFH jika ada dengan filter rentang tanggal
        wfh_names, wfh_data = {}, {}
        if has_wfh:
            wfh_names, wfh_data = read_wfh_data(wfh_path, start_day=start_day, end_day=end_day)

        # 3. Unified data untuk unmatched check
        if has_fp and has_wfh:
            merged_names, merged_data = merge_attendance_data(
                fp_names, fp_data, wfh_names, wfh_data
            )
        elif has_fp:
            merged_names, merged_data = fp_names, fp_data
        else:
            merged_names, merged_data = wfh_names, wfh_data

        result = {
            "pns": None,
            "pppk": None,
            "unmatched_fp": [],
            "has_fp": bool(has_fp and fp_names),
            "has_wfh": bool(has_wfh and wfh_names),
            "start_day": start_day,
            "end_day": end_day,
        }
        matched_fp_ids = set()

        def get_rec(score):
            if score >= AUTO_MATCH_THRESHOLD:
                return "high", f"Sangat Cocok ({score}%)"
            elif score >= 40:
                return "mid", f"Perlu Diperiksa ({score}%)"
            elif score > 0:
                return "low", f"Kecocokan Rendah ({score}%)"
            return "none", "Tidak Ditemukan"

        for tpl_type, tpl_path in [("pns", pns_path), ("pppk", pppk_path)]:
            if not tpl_path or not os.path.exists(tpl_path):
                continue

            employees = read_template_employees(tpl_path)
            matches_wfo = []
            matches_wfh = []

            for emp in employees:
                # 1. Matching terhadap Fingerprint (WFO)
                if has_fp and fp_names:
                    f_id, f_score, f_mtype, _ = find_best_match(emp["name"], fp_names)
                    f_days = sorted([d for d in fp_data.get(f_id, {}).keys() if start_day <= d <= end_day]) if (f_id and f_id in fp_data) else []
                else:
                    f_id, f_score, f_mtype = None, 0, "none"
                    f_days = []

                # 2. Matching terhadap WFH
                if has_wfh and wfh_names:
                    w_id, w_score, w_mtype, _ = find_best_match(emp["name"], wfh_names)
                    w_days = sorted([d for d in wfh_data.get(w_id, {}).keys() if start_day <= d <= end_day]) if (w_id and w_id in wfh_data) else []
                else:
                    w_id, w_score, w_mtype = None, 0, "none"
                    w_days = []

                # Tanggal yang bentrok / muncul di kedua sumber
                conflict_days = sorted(list(set(f_days) & set(w_days)))
                has_conflict = len(conflict_days) > 0

                f_rec_level, f_rec_label = get_rec(f_score)
                w_rec_level, w_rec_label = get_rec(w_score)

                if f_id:
                    matched_fp_ids.add(f_id)
                if w_id:
                    matched_fp_ids.add(w_id)

                matches_wfo.append({
                    "tpl_no":        emp["no"],
                    "tpl_row":       emp["row"],
                    "tpl_name":      emp["name"],
                    "fp_id":         f_id,
                    "matched_name":  fp_names.get(f_id, "") if f_id else "",
                    "source":        "WFO",
                    "score":         f_score,
                    "match_type":    f_mtype,
                    "rec_level":     f_rec_level,
                    "rec_label":     f_rec_label,
                    "days":          f_days,
                    "days_count":    len(f_days),
                    "has_conflict":  has_conflict,
                    "conflict_days": conflict_days,
                    "has_wfh":       len(w_days) > 0,
                    "wfh_days":      w_days,
                })

                matches_wfh.append({
                    "tpl_no":        emp["no"],
                    "tpl_row":       emp["row"],
                    "tpl_name":      emp["name"],
                    "fp_id":         w_id,
                    "matched_name":  wfh_names.get(w_id, "") if w_id else "",
                    "source":        "WFH",
                    "score":         w_score,
                    "match_type":    w_mtype,
                    "rec_level":     w_rec_level,
                    "rec_label":     w_rec_label,
                    "days":          w_days,
                    "days_count":    len(w_days),
                    "has_conflict":  has_conflict,
                    "conflict_days": conflict_days,
                    "has_wfo":       len(f_days) > 0,
                    "wfo_days":      f_days,
                })

            result[tpl_type] = {
                "total":        len(employees),
                "wfo":          matches_wfo,
                "wfh":          matches_wfh,
                "matches_fp":   matches_wfo,
                "matches_wfh":  matches_wfh,
            }

        # Employee dari sumber data yang tidak tercocokkan ke mana pun
        for src_id, src_name in merged_names.items():
            if src_id not in matched_fp_ids:
                result["unmatched_fp"].append({
                    "fp_id":   src_id,
                    "fp_name": src_name,
                })

        result["saved_decisions"] = session.get("decisions", {})
        return jsonify(result)
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/save-decisions", methods=["POST"])
def save_decisions():
    """Simpan keputusan user ke session."""
    data = request.get_json(silent=True) or {}
    decisions = data.get("decisions", {})
    session["decisions"] = decisions
    return jsonify({"success": True})


@app.route("/process", methods=["POST"])
def process():
    """Process cleaning + fill templates → download ZIP."""
    fp_path   = session.get("fp_path")
    wfh_path  = session.get("wfh_path")
    pns_path  = session.get("tpl_pns_path")
    pppk_path = session.get("tpl_pppk_path")

    has_fp  = fp_path and os.path.exists(fp_path)
    has_wfh = wfh_path and os.path.exists(wfh_path)

    if not has_fp and not has_wfh:
        return jsonify({"error": "Upload minimal 1 file absensi (Fingerprint atau WFH)"}), 400
    if not pns_path and not pppk_path:
        return jsonify({"error": "Minimal upload 1 template"}), 400

    data = request.get_json(silent=True) or {}
    decisions = data.get("decisions") or session.get("decisions", {})
    try:
        start_day = int(data.get("start_day") or session.get("start_day") or 1)
        end_day = int(data.get("end_day") or session.get("end_day") or 31)
    except Exception:
        start_day, end_day = 1, 31

    if start_day > end_day:
        start_day, end_day = end_day, start_day

    try:
        fp_names, fp_data = {}, {}
        if has_fp:
            rows = read_fingerprint(fp_path)
            fp_names, fp_data = clean_fingerprint(rows, start_day=start_day, end_day=end_day)

        wfh_names, wfh_data = {}, {}
        if has_wfh:
            wfh_names, wfh_data = read_wfh_data(wfh_path, start_day=start_day, end_day=end_day)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for tpl_type, tpl_path in [("PNS", pns_path), ("PPPK", pppk_path)]:
                if not tpl_path or not os.path.exists(tpl_path):
                    continue

                employees = read_template_employees(tpl_path)
                tpl_decisions = {}
                if isinstance(decisions, dict):
                    tpl_decisions = decisions.get(tpl_type) or decisions.get(tpl_type.lower()) or {}

                out_path = os.path.join(UPLOAD_FOLDER, f"{tpl_type}_FILLED.xlsx")
                filled = fill_template(
                    tpl_path, out_path, employees,
                    fp_names, fp_data, wfh_names, wfh_data,
                    decisions=tpl_decisions,
                    start_day=start_day, end_day=end_day
                )

                orig_name = session.get(f"tpl_{tpl_type.lower()}_filename", f"{tpl_type}.xlsx")
                base = os.path.splitext(orig_name)[0]
                zip_name = f"{base}_TERISI.xlsx"
                zf.write(out_path, zip_name)

                if os.path.exists(out_path):
                    os.remove(out_path)

        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name="Absensi_Bersih.zip",
        )
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    print("\n" + "=" * 54)
    print("  ABSENSI CLEANER v2 — Web App")
    print("  Buka browser: http://localhost:5000")
    print("=" * 54 + "\n")
    app.run(debug=False, port=5000)
