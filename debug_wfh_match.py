"""Debug: Check WFH name matching against templates"""
import sys
sys.path.insert(0, '.')
from app import *

# Read WFH data
wfh_names, wfh_data = read_wfh_data('uploads/wfh_data.xlsx')
print(f'WFH employees parsed: {len(wfh_names)}')

# Read fingerprint data (if exists)
fp_names, fp_data = {}, {}
fp_path = 'uploads/fingerprint.xls'
import os
for ext in ['.xls', '.xlsx', '.xlsm']:
    p = f'uploads/fingerprint{ext}'
    if os.path.exists(p):
        fp_path = p
        break

if os.path.exists(fp_path):
    rows = read_fingerprint(fp_path)
    fp_names, fp_data = clean_fingerprint(rows)
    print(f'Fingerprint employees parsed: {len(fp_names)}')
else:
    print('No fingerprint file found')

# Read template employees
pns_emps, pppk_emps = [], []
if os.path.exists('uploads/template_pns.xlsx'):
    pns_emps = read_template_employees('uploads/template_pns.xlsx')
if os.path.exists('uploads/template_pppk.xlsx'):
    pppk_emps = read_template_employees('uploads/template_pppk.xlsx')
print(f'PNS template employees: {len(pns_emps)}')
print(f'PPPK template employees: {len(pppk_emps)}')

all_tpl = pns_emps + pppk_emps
all_tpl_names = {e['name'] for e in all_tpl}

# Step 1: Check WFH -> Fingerprint cross-reference (merge step)
print('\n' + '='*80)
print('STEP 1: WFH -> Fingerprint Cross-Reference (merge_attendance_data)')
print('='*80)
for wfh_id, wfh_name in sorted(wfh_names.items(), key=lambda x: x[1]):
    best_score = 0
    best_fp_name = ''
    best_fp_id = ''
    for fp_id, fp_name in fp_names.items():
        score, _ = match_score(wfh_name, fp_name)
        if score > best_score:
            best_score = score
            best_fp_name = fp_name
            best_fp_id = fp_id
    if best_score >= AUTO_MATCH_THRESHOLD:
        print(f'  [MERGED ] score={best_score:3d} WFH:"{wfh_name}" -> FP:"{best_fp_name}" (id={best_fp_id})')
    elif best_score > 0:
        print(f'  [WFH-ONLY] score={best_score:3d} WFH:"{wfh_name}" -> best FP:"{best_fp_name}" (BELOW THRESHOLD)')
    else:
        print(f'  [WFH-ONLY] score=  0 WFH:"{wfh_name}" -> no FP match at all')

# Step 2: After merge, check template -> merged_names matching
print('\n' + '='*80)
print('STEP 2: Template -> Merged Names Matching')
print('='*80)

if fp_names:
    merged_names, merged_data = merge_attendance_data(fp_names, fp_data, wfh_names, wfh_data)
else:
    merged_names, merged_data = wfh_names, wfh_data

print(f'Merged employees: {len(merged_names)}')

unmatched_wfh = []
for emp in all_tpl:
    fp_id, score, mtype, needs_review = find_best_match(emp['name'], merged_names)
    matched_name = merged_names.get(fp_id, '') if fp_id else ''
    
    # Check if this matched a WFH entry
    is_wfh = fp_id and fp_id.startswith('WFH_') if fp_id else False
    
    if is_wfh or (not fp_id and emp['name'] in [wfh_names[wid] for wid in wfh_names]):
        status = 'AUTO' if (score >= AUTO_MATCH_THRESHOLD and not needs_review) else 'REVIEW' if needs_review else 'NO MATCH'
        print(f'  [{status:8s}] score={score:3d} ({mtype:12s}) TPL:"{emp["name"]}" -> "{matched_name}" (id={fp_id})')
        if status != 'AUTO':
            unmatched_wfh.append((emp['name'], matched_name, score, mtype, status))

print(f'\n{"="*80}')
print(f'WFH-RELATED ISSUES: {len(unmatched_wfh)} template names with WFH problems')
print(f'{"="*80}')
for tpl_name, matched, score, mtype, status in unmatched_wfh:
    print(f'  [{status:8s}] score={score:3d} TPL:"{tpl_name}" -> "{matched}"')
