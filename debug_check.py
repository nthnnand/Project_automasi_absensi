# -*- coding: utf-8 -*-
"""
Debug script: cek kenapa data tidak masuk ke template.
Jalankan setelah upload file via web UI.
"""
import os, glob, json
from app import (
    read_fingerprint, clean_fingerprint, read_wfh_data,
    merge_attendance_data, read_template_employees,
    find_best_match, norm_name, AUTO_MATCH_THRESHOLD
)

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")

print("=" * 60)
print("  DEBUG: Cek kenapa data tidak masuk")
print("=" * 60)

# 1. Cek file apa saja yang ada di uploads/
print("\n📁 File di uploads/:")
files = glob.glob(os.path.join(UPLOAD_DIR, "*"))
if not files:
    print("  ❌ KOSONG! Belum ada file yang di-upload.")
    exit()
for f in files:
    size = os.path.getsize(f)
    print(f"  • {os.path.basename(f)} ({size:,} bytes)")

# 2. Coba baca fingerprint
fp_files = [f for f in files if os.path.basename(f).startswith("fingerprint")]
fp_names, fp_data = {}, {}
if fp_files:
    fp_path = fp_files[0]
    print(f"\n📊 Parsing WFO (fingerprint): {os.path.basename(fp_path)}")
    try:
        rows = read_fingerprint(fp_path)
        print(f"  Raw rows: {len(rows)}")
        if rows:
            print(f"  Sample row 0: {rows[0]}")
            # Cek format waktu
            for r in rows[:5]:
                print(f"    ID={r['no_id']} | Nama={r['nama']} | Waktu={r['waktu']} | Status={r['status']} | StatusBaru={r.get('status_baru','')}")
        
        fp_names, fp_data = clean_fingerprint(rows)
        print(f"  ✅ Employees (cleaned): {len(fp_names)}")
        
        if not fp_names:
            print("  ⚠️  TIDAK ADA employee setelah cleaning!")
            print("     Kemungkinan: format waktu bukan 'dd/mm/yyyy HH:MM'")
            print("     Atau status bukan 'C/Masuk' / 'C/Keluar' / 'Lembur Masuk' / 'Lembur Keluar'")
            # Cek status unik
            statuses = set()
            statuses_baru = set()
            for r in rows:
                statuses.add(r['status'])
                if r.get('status_baru'):
                    statuses_baru.add(r['status_baru'])
            print(f"     Status unik (Kolom E): {statuses}")
            print(f"     Status Baru unik (Kolom F): {statuses_baru}")
        else:
            print(f"  Nama karyawan WFO: {list(fp_names.values())[:10]}...")
            # Cek hari apa saja yang ada data
            all_days = set()
            for fid, days in fp_data.items():
                all_days.update(days.keys())
            print(f"  Hari dengan data: {sorted(all_days)}")
            # Sample data
            sample_id = list(fp_data.keys())[0]
            print(f"  Sample data ID '{sample_id}': {fp_data[sample_id]}")
    except Exception as e:
        print(f"  ❌ ERROR: {e}")
        import traceback; traceback.print_exc()
else:
    print("\n📊 WFO (fingerprint): tidak di-upload")

# 3. Coba baca WFH
wfh_files = [f for f in files if os.path.basename(f).startswith("wfh")]
wfh_names, wfh_data = {}, {}
if wfh_files:
    wfh_path = wfh_files[0]
    print(f"\n🏠 Parsing WFH: {os.path.basename(wfh_path)}")
    try:
        wfh_names, wfh_data = read_wfh_data(wfh_path)
        print(f"  ✅ Employees: {len(wfh_names)}")
        if wfh_names:
            print(f"  Nama karyawan WFH: {list(wfh_names.values())[:10]}...")
            all_days = set()
            for wid, days in wfh_data.items():
                all_days.update(days.keys())
            print(f"  Hari dengan data: {sorted(all_days)}")
    except Exception as e:
        print(f"  ❌ ERROR: {e}")
        import traceback; traceback.print_exc()
else:
    print("\n🏠 WFH: tidak di-upload")

# 4. Merge
if fp_names or wfh_names:
    if fp_names and wfh_names:
        merged_names, merged_data = merge_attendance_data(fp_names, fp_data, wfh_names, wfh_data)
    elif fp_names:
        merged_names, merged_data = fp_names, fp_data
    else:
        merged_names, merged_data = wfh_names, wfh_data
    
    print(f"\n🔗 Merged: {len(merged_names)} employees total")
else:
    print("\n❌ Tidak ada data absensi untuk diproses!")
    exit()

# 5. Cek template dan matching
for tpl_prefix, tpl_label in [("template_pns", "PNS"), ("template_pppk", "PPPK")]:
    tpl_files = [f for f in files if os.path.basename(f).startswith(tpl_prefix)]
    if not tpl_files:
        print(f"\n📋 Template {tpl_label}: tidak di-upload")
        continue
    
    tpl_path = tpl_files[0]
    print(f"\n📋 Template {tpl_label}: {os.path.basename(tpl_path)}")
    try:
        employees = read_template_employees(tpl_path)
        print(f"  Karyawan di template: {len(employees)}")
        if employees:
            print(f"  Sample: {[e['name'] for e in employees[:5]]}")
        
        # Cek matching
        auto_count = 0
        review_count = 0
        no_match_count = 0
        
        print(f"\n  {'Template Name':<30} {'Matched To':<30} {'Score':>5} {'Type':<12} {'Status'}")
        print(f"  {'-'*30} {'-'*30} {'-'*5} {'-'*12} {'-'*15}")
        
        for emp in employees:
            fp_id, score, mtype, needs_review = find_best_match(emp["name"], merged_names)
            fp_name = merged_names.get(fp_id, "-") if fp_id else "-"
            
            if fp_id and score >= AUTO_MATCH_THRESHOLD and not needs_review:
                status = "✅ Auto-Fill"
                auto_count += 1
            elif fp_id and needs_review:
                status = "⚠️  Perlu Review"
                review_count += 1
            else:
                status = "❌ Tidak cocok"
                no_match_count += 1
            
            has_data = "📊" if (fp_id and fp_id in merged_data) else "  "
            print(f"  {emp['name']:<30} {fp_name:<30} {score:>5} {mtype:<12} {status} {has_data}")
        
        print(f"\n  Summary: ✅ Auto={auto_count} | ⚠️  Review={review_count} | ❌ No Match={no_match_count}")
        
        if auto_count == 0 and review_count == 0:
            print("  🚨 TIDAK ADA YANG MATCH! Ini sebabnya template kosong.")
            print("     Kemungkinan penyebab:")
            print("     1. Nama di fingerprint SANGAT BERBEDA dari nama di template")
            print("     2. Data fingerprint kosong setelah cleaning (cek status)")
            print("     3. Format file tidak sesuai")
    except Exception as e:
        print(f"  ❌ ERROR: {e}")
        import traceback; traceback.print_exc()

print("\n" + "=" * 60)
print("  DEBUG SELESAI")
print("=" * 60)
