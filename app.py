# -*- coding: utf-8 -*-
"""
Absensi Cleaner v2 — Flask Web App
====================================
Otomatis: baca fingerprint → cleaning → cocokkan nama → isi template PNS/PPPK

Fixes:
  Case 1 — Prioritas Status Baru (Kolom F): F > E, Lembur Masuk/Keluar dikenali
  Case 2 — Preserve Data Validation dropdown (keep_vba untuk .xlsm)
  Case 3 — Fuzzy matching ketat: threshold 80, deteksi multi-kandidat ambiguitas
  Case 4 — Warna merah pada sel jam kosong (data tidak tersedia)
"""

import os, io, re, zipfile, copy, shutil, difflib
from datetime import time as dt_time
from flask import Flask, render_template, request, send_file, jsonify, session
import xlrd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = "absensi_cleaner_v2_2026"
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Auto-match threshold: skor >= ini → langsung diisi ke template
# 65 dipilih agar kasus typo seperti "ria mustikah" vs "RIA MUSTIKA" tetap ter-match
AUTO_MATCH_THRESHOLD = 65


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


def norm_name(name):
    """Normalize name for matching: uppercase, strip punctuation, collapse spaces."""
    s = str(name).upper().strip()
    s = re.sub(r'[.,\-\'\"()]+', ' ', s)   # remove punctuation
    s = re.sub(r'\s+', ' ', s).strip()      # collapse spaces
    return s


def name_words(name):
    """Get list of words from normalized name (min 2 chars)."""
    return [w for w in norm_name(name).split() if len(w) >= 2]


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


def clean_fingerprint(rows):
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
#  TEMPLATE: READ EMPLOYEE LIST
# ══════════════════════════════════════════════════════════════

def read_template_employees(filepath):
    """
    Read employee names from Input Data Absen sheet.
    Returns list of {row: int, name: str} (row = Excel row number).
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)

    ws = wb['Input Data Absen']
    employees = []
    for r in range(6, ws.max_row + 1):
        no = ws.cell(r, 1).value
        nama = ws.cell(r, 2).value
        if no is not None and nama and str(nama).strip():
            employees.append({
                "row": r,
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


def fill_template(src_path, out_path, matching, fp_data, decisions=None):
    """
    Copy template, fill M/P columns in 'Input Data Absen'.
    matching  = [{tpl_row, tpl_name, fp_id, score, match_type, needs_review}, ...]
    fp_data   = {fp_id: {day: {masuk: 'HH:MM', keluar: 'HH:MM'}}}
    decisions = {tpl_row_str_or_name: 'include' | 'ignore', ...}

    Case 2 — keep_vba=True untuk .xlsm: preserves Data Validation (dropdown kolom K)
    Case 3 — Auto-fill jika score >= AUTO_MATCH_THRESHOLD dan tidak needs_review,
             ATAU jika user memilih 'include'/'masukkan' untuk item review
    Case 4 — Sel jam yang kosong / 00:00 diberi warna merah (data tidak tersedia)
    """
    # Salin file asli agar tidak berubah
    shutil.copy2(src_path, out_path)

    # Case 2: keep_vba=True untuk .xlsm agar Data Validation (dropdown) tidak hilang
    ext = os.path.splitext(src_path)[1].lower()
    wb = load_workbook(out_path, keep_vba=(ext == ".xlsm"))
    ws = wb['Input Data Absen']

    # Case 4: Warna merah untuk sel kosong / 00:00
    red_fill = PatternFill(start_color="FFC0392B", end_color="FFC0392B", fill_type="solid")

    filled_count = 0

    # Kumpulkan semua baris karyawan (untuk pass merah nanti)
    employee_rows = set()

    # Kumpulkan semua hari yang ada data di fingerprint
    valid_days = set()
    for days_data in fp_data.values():
        valid_days.update(d for d in days_data.keys() if 1 <= d <= 31)

    decisions = decisions or {}

    for m in matching:
        fp_id        = m.get("fp_id")
        score        = m.get("score", 0)
        needs_review = m.get("needs_review", False)
        tpl_row      = m["tpl_row"]
        tpl_name     = m.get("tpl_name", "")

        employee_rows.add(tpl_row)

        if not fp_id or fp_id not in fp_data:
            # Tidak ada ID / data fingerprint untuk pegawai ini
            continue

        # Cek keputusan user jika ada (bisa berdasarkan row string/int atau nama)
        user_choice = (
            decisions.get(str(tpl_row)) or 
            decisions.get(tpl_row) or 
            decisions.get(tpl_name) or
            decisions.get(norm_name(tpl_name))
        )

        should_fill = False
        if user_choice in ("include", "masukkan"):
            # User eksplisit memilih 'Masukkan'
            should_fill = True
        elif user_choice in ("ignore", "abaikan"):
            # User eksplisit memilih 'Abaikan'
            should_fill = False
        elif score >= AUTO_MATCH_THRESHOLD and not needs_review:
            # Skor >= 65 dan tidak ambigu: auto-fill tetap otomatis diproses
            should_fill = True
        else:
            # Skor di bawah threshold / perlu review dan belum/tidak dipilih masukkan
            should_fill = False

        if not should_fill:
            # Baris ini tidak diisi (akan ditandai merah di pass berikutnya)
            continue

        days_data = fp_data[fp_id]

        for day, times in days_data.items():
            if day < 1 or day > 31:
                continue

            # Mapping: tanggal d → kolom Masuk = 3*d+3, kolom Pulang = 3*d+4
            col_m = 3 * day + 3
            col_p = 3 * day + 4

            if "masuk" in times:
                t = time_str_to_dt(times["masuk"])
                if t:
                    ws.cell(row=tpl_row, column=col_m, value=t)
                    filled_count += 1

            if "keluar" in times:
                t = time_str_to_dt(times["keluar"])
                if t:
                    ws.cell(row=tpl_row, column=col_p, value=t)
                    filled_count += 1

    # Case 4: Pass kedua — tandai sel jam yang kosong/00:00 dengan warna merah
    # Hanya untuk hari yang ada di data fingerprint (valid_days)
    # Menangkap: None, 0, "", time(0,0,0) = semua kondisi "tidak terisi"
    for tpl_row in employee_rows:
        for day in valid_days:
            col_m = 3 * day + 3
            col_p = 3 * day + 4
            for col in (col_m, col_p):
                cell = ws.cell(row=tpl_row, column=col)
                if is_empty_time_cell(cell.value):
                    cell.fill = red_fill

    wb.save(out_path)
    wb.close()
    return filled_count


# ══════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/debug-matching", methods=["POST"])
def debug_matching():
    """
    Debug endpoint: tampilkan detail parsing fingerprint + skor matching.
    Berguna untuk verifikasi Case 1 (Status Baru) dan Case 3 (fuzzy match).
    """
    fp_path  = session.get("fp_path")
    pns_path = session.get("tpl_pns_path")
    pppk_path = session.get("tpl_pppk_path")

    if not fp_path or not os.path.exists(fp_path):
        return jsonify({"error": "File fingerprint belum diupload"}), 400

    try:
        rows = read_fingerprint(fp_path)
        fp_names, fp_data = clean_fingerprint(rows)

        # Sample raw rows (max 30) untuk cek Case 1
        raw_sample = []
        for r in rows[:50]:
            if r.get("status_baru"):
                raw_sample.append({
                    "nama":        r["nama"],
                    "waktu":       r["waktu"],
                    "status_e":    r["status"],
                    "status_baru": r["status_baru"],
                })

        # Detail matching untuk setiap template
        debug_matches = {}
        for tpl_type, tpl_path in [("pns", pns_path), ("pppk", pppk_path)]:
            if not tpl_path or not os.path.exists(tpl_path):
                continue
            employees = read_template_employees(tpl_path)
            matches = []
            for emp in employees:
                fp_id, score, mtype, needs_review = find_best_match(emp["name"], fp_names)
                matches.append({
                    "tpl_name":    emp["name"],
                    "fp_name":     fp_names.get(fp_id, "") if fp_id else "",
                    "score":       score,
                    "match_type":  mtype,
                    "needs_review": needs_review,
                    "will_fill":   bool(fp_id and score >= AUTO_MATCH_THRESHOLD and not needs_review),
                })
            debug_matches[tpl_type] = matches

        return jsonify({
            "threshold":      AUTO_MATCH_THRESHOLD,
            "fp_employees":   len(fp_names),
            "fp_names":       fp_names,
            "status_baru_sample": raw_sample,  # Baris dengan Status Baru terisi (Case 1)
            "matching":       debug_matches,
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
    fp_path  = session.get("fp_path")
    pns_path = session.get("tpl_pns_path")
    pppk_path = session.get("tpl_pppk_path")

    if not fp_path or not os.path.exists(fp_path):
        return jsonify({"error": "File fingerprint belum diupload"}), 400
    if not pns_path and not pppk_path:
        return jsonify({"error": "Minimal upload 1 template (PNS atau PPPK)"}), 400

    try:
        # Clean fingerprint
        rows = read_fingerprint(fp_path)
        fp_names, fp_data = clean_fingerprint(rows)

        result = {"pns": None, "pppk": None, "unmatched_fp": []}
        matched_fp_ids = set()

        for tpl_type, tpl_path in [("pns", pns_path), ("pppk", pppk_path)]:
            if not tpl_path or not os.path.exists(tpl_path):
                continue

            employees = read_template_employees(tpl_path)
            matches = []

            for emp in employees:
                # 4-tuple: fp_id, score, match_type, needs_review
                fp_id, score, mtype, needs_review = find_best_match(emp["name"], fp_names)

                match_info = {
                    "tpl_no":       emp["no"],
                    "tpl_row":      emp["row"],
                    "tpl_name":     emp["name"],
                    "fp_id":        fp_id,
                    "fp_name":      fp_names.get(fp_id, "") if fp_id else "",
                    "score":        score,
                    "match_type":   mtype,
                    "needs_review": needs_review,
                    "days_with_data": 0,
                }
                if fp_id and fp_id in fp_data:
                    match_info["days_with_data"] = len(fp_data[fp_id])
                    matched_fp_ids.add(fp_id)
                matches.append(match_info)

            # Hitung matched: hanya yang score >= threshold DAN tidak needs_review (Auto-Fill)
            matched = sum(
                1 for m in matches
                if m["fp_id"] and m["score"] >= AUTO_MATCH_THRESHOLD and not m["needs_review"]
            )
            need_review_count = sum(
                1 for m in matches
                if m["fp_id"] and m["needs_review"]
            )
            result[tpl_type] = {
                "total":        len(employees),
                "matched":      matched,
                "need_review":  need_review_count,
                "unmatched":    len(employees) - matched - need_review_count,
                "matches":      matches,
            }

        # Fingerprint employee yang tidak tercocokkan ke mana pun
        for fp_id, fp_name in fp_names.items():
            if fp_id not in matched_fp_ids:
                result["unmatched_fp"].append({
                    "fp_id":   fp_id,
                    "fp_name": fp_name,
                })

        # Sertakan keputusan yang tersimpan di session jika ada
        result["saved_decisions"] = session.get("decisions", {})

        return jsonify(result)
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/save-decisions", methods=["POST"])
def save_decisions():
    """Simpan keputusan user ('include' / 'ignore') untuk item 'Perlu Review' ke session."""
    data = request.get_json(silent=True) or {}
    decisions = data.get("decisions", {})
    session["decisions"] = decisions
    return jsonify({"success": True})


@app.route("/process", methods=["POST"])
def process():
    """Process cleaning + fill templates → download ZIP."""
    fp_path   = session.get("fp_path")
    pns_path  = session.get("tpl_pns_path")
    pppk_path = session.get("tpl_pppk_path")

    if not fp_path or not os.path.exists(fp_path):
        return jsonify({"error": "File fingerprint belum diupload"}), 400
    if not pns_path and not pppk_path:
        return jsonify({"error": "Minimal upload 1 template"}), 400

    # Ambil keputusan user dari request payload atau session
    data = request.get_json(silent=True) or {}
    decisions = data.get("decisions") or session.get("decisions", {})

    try:
        # 1. Clean fingerprint
        rows = read_fingerprint(fp_path)
        fp_names, fp_data = clean_fingerprint(rows)

        # 2. Proses setiap template
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:

            for tpl_type, tpl_path in [("PNS", pns_path), ("PPPK", pppk_path)]:
                if not tpl_path or not os.path.exists(tpl_path):
                    continue

                employees = read_template_employees(tpl_path)

                # Bangun matching list dengan 4-tuple (termasuk needs_review)
                matching = []
                for emp in employees:
                    fp_id, score, mtype, needs_review = find_best_match(emp["name"], fp_names)
                    matching.append({
                        "tpl_row":      emp["row"],
                        "tpl_name":     emp["name"],
                        "fp_id":        fp_id,
                        "score":        score,
                        "match_type":   mtype,
                        "needs_review": needs_review,
                    })

                # Ambil keputusan spesifik untuk template ini (PNS / PPPK)
                tpl_decisions = {}
                if isinstance(decisions, dict):
                    tpl_decisions = decisions.get(tpl_type) or decisions.get(tpl_type.lower()) or {}

                # Isi template (buat salinan)
                out_path = os.path.join(UPLOAD_FOLDER, f"{tpl_type}_FILLED.xlsx")
                filled = fill_template(tpl_path, out_path, matching, fp_data, decisions=tpl_decisions)

                # Tambahkan ke ZIP
                orig_name = session.get(f"tpl_{tpl_type.lower()}_filename", f"{tpl_type}.xlsx")
                base = os.path.splitext(orig_name)[0]
                zip_name = f"{base}_TERISI.xlsx"
                zf.write(out_path, zip_name)

                # Bersihkan file sementara
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
