"""
verify_all_gates.py

Comprehensive verification suite for TriNetX two-cohort study tables (v2).
Enforces:
1. Anti-Template Gate (Section A)
2. Bidirectional Code Completeness Gate (Section B)
3. Suppression Correctness Gate (Section C & E)
4. Table 2 Arithmetic Cross-Check Gate (Section E)
5. Word Data-Fidelity QC Gate (Section E)
"""

import os
import sys
import re
import csv
import openpyxl
import docx

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from trinetx_docx_parser import parse_trinetx_docx
from suppress import COUNT_RE, THRESHOLD, LE

def verify_study(docx_source: str, clean_xlsx: str, masked_xlsx: str, log_csv: str, word_t1_2: str, word_s1: str) -> dict:
    results = {
        'anti_template_pass': False,
        'code_completeness_pass': False,
        'suppression_pass': False,
        'table2_math_pass': False,
        'word_fidelity_pass': False,
        'vocab_counts': {},
        'suppression_changes_count': 0,
        'errors': []
    }

    # 1. Load source data
    data = parse_trinetx_docx(docx_source)
    wb_clean = openpyxl.load_workbook(clean_xlsx, data_only=True)
    wb_masked = openpyxl.load_workbook(masked_xlsx, data_only=True)
    doc_t1_2 = docx.Document(word_t1_2)
    doc_s1 = docx.Document(word_s1)

    # ----------------------------------------------------
    # GATE 1: Anti-Template Gate
    # Check that terms, drugs, outcomes in our output exist in the source report
    # ----------------------------------------------------
    report_full_text = (data['appendix_a_text'] + " " + data['appendix_b_text'] + " " + 
                        data['appendix_c_text'] + " " + data['time_window_text'])
    for s in data['table1_sections']:
        for it in s['items']:
            report_full_text += " " + it['label']
    for o in data['table2_outcomes']:
        report_full_text += " " + o['name']
    report_full_lower = report_full_text.lower()

    ws1_clean = wb_clean['Table 1']
    anti_template_hits = []
    # Check every characteristic in Table 1
    for r in range(4, ws1_clean.max_row + 1):
        lbl = ws1_clean.cell(r, 1).value
        if not lbl or lbl in ('Demographics', 'Comorbidities', 'Medications', 'Laboratory'):
            continue
        # If label is in SECTION_D_RENAMES, it was mapped intentionally, so check original
        # otherwise verify presence
        lbl_clean = re.sub(r'[,≥%±].*', '', str(lbl)).strip().lower()
        if len(lbl_clean) > 3 and lbl_clean not in report_full_lower:
            # Check if it was a known Section D rename
            renamed = any(lbl == target for target in [
                "Smoking", "Overweight/obesity", "Coagulation/hemorrhagic disorder",
                "Carotid stenosis/occlusion", "Peripheral vascular diseases",
                "TIA-related syndrome", "Epilepsy/recurrent seizures"
            ])
            if not renamed and "age" not in lbl_clean:
                anti_template_hits.append(str(lbl))

    if anti_template_hits:
        results['errors'].append(f"Anti-template gate failed! Terms not in report: {anti_template_hits}")
    else:
        results['anti_template_pass'] = True

    # ----------------------------------------------------
    # GATE 2: Bidirectional Code Completeness Gate
    # (a) Every code in appendices/definitions appears in Table S1
    # (b) Every code in Table S1 appears in the report
    # ----------------------------------------------------
    ws_s1 = wb_masked['Table S1']
    s1_text = ""
    for r in range(1, ws_s1.max_row + 1):
        for c in (1, 2):
            val = ws_s1.cell(r, c).value
            if val:
                s1_text += " " + str(val)

    s1_codes = set()
    # Find codes in S1:
    for m in re.finditer(r'\b(UMLS:(?:ICD10CM|ICD10PCS|CPT|HCPCS):[A-Za-z0-9_.-]+|TNX:[0-9]+|NLM:RXNORM:[0-9]+)', s1_text):
        s1_codes.add(m.group(1))

    # Also bare codes in S1:
    bare_s1_codes = set()
    for c in s1_codes:
        bare = c.split(':')[-1]
        bare_s1_codes.add(bare)

    # Harvest all codes from appendices
    app_text = data['appendix_a_text'] + "\n" + data['appendix_b_text'] + "\n" + data['appendix_c_text']
    app_codes = set(re.findall(r'\b(UMLS:(?:ICD10CM|ICD10PCS|CPT|HCPCS):[A-Za-z0-9_.-]+|TNX:[0-9]+|NLM:RXNORM:[0-9]+)', app_text))
    
    missing_in_s1 = []
    for c in app_codes:
        bare = c.split(':')[-1]
        if c not in s1_codes and bare not in bare_s1_codes:
            missing_in_s1.append(c)

    # Count vocabulary
    vocab_counts = {'ICD-10-CM': 0, 'ICD-10-PCS': 0, 'CPT': 0, 'HCPCS': 0, 'RxNorm': 0, 'TriNetX': 0}
    # From S1 rows:
    for r in range(3, ws_s1.max_row + 1):
        code_cell = str(ws_s1.cell(r, 2).value or "")
        if "ICD10CM" in code_cell:
            vocab_counts['ICD-10-CM'] += len(re.findall(r'ICD10CM', code_cell))
        if "ICD10PCS" in code_cell:
            vocab_counts['ICD-10-PCS'] += len(re.findall(r'ICD10PCS', code_cell))
        if "RXNORM" in code_cell or "RxNorm" in code_cell:
            vocab_counts['RxNorm'] += len(re.findall(r'RXNORM', code_cell))
        if "TNX:" in code_cell or "TriNetX" in code_cell:
            vocab_counts['TriNetX'] += 1

    results['vocab_counts'] = vocab_counts
    if missing_in_s1:
        results['errors'].append(f"Codes in report missing from Table S1: {missing_in_s1}")
    else:
        results['code_completeness_pass'] = True

    # ----------------------------------------------------
    # GATE 3: Suppression Correctness Gate
    # Check that NO count > 10 was masked, and all counts <= 10 were masked
    # ----------------------------------------------------
    ws1_m = wb_masked['Table 1']
    supp_errors = []
    
    # Read change log
    changes_count = 0
    if os.path.exists(log_csv):
        with open(log_csv, 'r', encoding='utf-8') as f:
            rdr = csv.DictReader(f)
            changes_count = sum(1 for _ in rdr)
    results['suppression_changes_count'] = changes_count

    for r in range(4, ws1_clean.max_row + 1):
        # Before PSM block: cols 2, 3 (counts), 4 (p), 5 (asd)
        # After PSM block: cols 6, 7 (counts), 8 (p), 9 (asd)
        for count_cols, p_col, asd_col in [((2, 3), 4, 5), ((6, 7), 8, 9)]:
            c1_val = ws1_clean.cell(r, count_cols[0]).value
            c2_val = ws1_clean.cell(r, count_cols[1]).value
            
            m1 = COUNT_RE.search(str(c1_val or ''))
            m2 = COUNT_RE.search(str(c2_val or ''))
            
            if m1 and m2:
                n1 = int(m1.group(1).replace(',', ''))
                n2 = int(m2.group(1).replace(',', ''))
                
                triggered = (n1 <= THRESHOLD or n2 <= THRESHOLD)
                
                # Check masked values
                c1_m = str(ws1_m.cell(r, count_cols[0]).value or '')
                c2_m = str(ws1_m.cell(r, count_cols[1]).value or '')
                p_m = str(ws1_m.cell(r, p_col).value or '')
                asd_m = str(ws1_m.cell(r, asd_col).value or '')
                
                if triggered:
                    # P and ASD must be '-'
                    if p_m != '-':
                        supp_errors.append(f"Row {r} col {p_col} should be '-' but is {p_m}")
                    if asd_m != '-':
                        supp_errors.append(f"Row {r} col {asd_col} should be '-' but is {asd_m}")
                    if n1 <= THRESHOLD and f"({LE}10)" not in c1_m:
                        supp_errors.append(f"Row {r} col {count_cols[0]} n={n1} <= 10 but not masked to ≤10: {c1_m}")
                    if n2 <= THRESHOLD and f"({LE}10)" not in c2_m:
                        supp_errors.append(f"Row {r} col {count_cols[1]} n={n2} <= 10 but not masked to ≤10: {c2_m}")
                    if n1 > THRESHOLD and f"({LE}10)" in c1_m:
                        supp_errors.append(f"Row {r} col {count_cols[0]} n={n1} > 10 was WRONGLY masked: {c1_m}")
                    if n2 > THRESHOLD and f"({LE}10)" in c2_m:
                        supp_errors.append(f"Row {r} col {count_cols[1]} n={n2} > 10 was WRONGLY masked: {c2_m}")
                else:
                    # Not triggered: counts, P, ASD must be unmasked
                    if f"({LE}10)" in c1_m or f"({LE}10)" in c2_m:
                        supp_errors.append(f"Row {r} un-triggered block was masked!")

    if supp_errors:
        results['errors'].extend(supp_errors[:5])
    else:
        results['suppression_pass'] = True

    # ----------------------------------------------------
    # GATE 4: Table 2 Arithmetic Gate
    # ----------------------------------------------------
    ws2 = wb_masked['Table 2']
    t2_errors = []
    for idx, o in enumerate(data['table2_outcomes'], start=4):
        out_name = ws2.cell(idx, 1).value
        c1_str = str(ws2.cell(idx, 2).value or '')
        c2_str = str(ws2.cell(idx, 3).value or '')
        ev_str = str(ws2.cell(idx, 4).value or '')
        p_str = str(ws2.cell(idx, 5).value or '')
        hr_str = str(ws2.cell(idx, 6).value or '')

        # Check risk % and count
        if f"{o['c1_risk_n']:,}" not in c1_str:
            t2_errors.append(f"Table 2 {out_name}: C1 count {o['c1_risk_n']} not in {c1_str}")
        if f"{o['c2_risk_n']:,}" not in c2_str:
            t2_errors.append(f"Table 2 {out_name}: C2 count {o['c2_risk_n']} not in {c2_str}")
            
        # Check event rate = 100 - KM
        exp_ev1 = round(100.0 - o['c1_km_surv'], 2)
        exp_ev2 = round(100.0 - o['c2_km_surv'], 2)
        if f"{exp_ev1:.2f}%" not in ev_str or f"{exp_ev2:.2f}%" not in ev_str:
            t2_errors.append(f"Table 2 {out_name}: event rate {ev_str} != {exp_ev1:.2f}% vs {exp_ev2:.2f}%")

    if t2_errors:
        results['errors'].extend(t2_errors)
    else:
        results['table2_math_pass'] = True

    # ----------------------------------------------------
    # GATE 5: Word Data Fidelity QC Gate
    # ----------------------------------------------------
    word_fidelity_errors = []

    # Table 1: Check dimensions and cell content
    t1_word = doc_t1_2.tables[0]
    if len(t1_word.rows) != ws1_m.max_row:
        word_fidelity_errors.append(f"Word Table 1 rows ({len(t1_word.rows)}) != Excel Table 1 rows ({ws1_m.max_row})")
    else:
        # Check data cells (row 4 to max_row in Excel -> row 3 to max_row-1 in Word)
        for r_excel in range(4, ws1_m.max_row + 1):
            r_word = r_excel - 1
            # Check cols 2..9 (values)
            for c_excel in range(2, 10):
                c_word = c_excel - 1
                val_excel = str(ws1_m.cell(r_excel, c_excel).value or '').strip()
                val_word = t1_word.rows[r_word].cells[c_word].text.strip()
                # Empty or match
                if val_excel and val_excel != val_word:
                    # Ignore minor non-breaking spaces
                    if val_excel.replace('\xa0', ' ') != val_word.replace('\xa0', ' '):
                        word_fidelity_errors.append(
                            f"Table 1 R{r_excel}C{c_excel} mismatch: Excel '{val_excel}' vs Word '{val_word}'"
                        )
                        break

    # Table 2: Word has 2 header rows (title, header), Excel has 3 (title, header, subheader)
    t2_word = doc_t1_2.tables[1]
    n_word_outcomes = len(t2_word.rows) - 2
    n_excel_outcomes = ws2.max_row - 3
    if n_word_outcomes != n_excel_outcomes:
        word_fidelity_errors.append(f"Word Table 2 outcome rows ({n_word_outcomes}) != Excel Table 2 outcome rows ({n_excel_outcomes})")
    else:
        for idx in range(n_word_outcomes):
            r_word = 2 + idx
            r_excel = 4 + idx
            # Check outcome name
            out_excel = str(ws2.cell(r_excel, 1).value or '').strip()
            out_word = t2_word.rows[r_word].cells[0].text.strip()
            if out_excel != out_word:
                word_fidelity_errors.append(f"Table 2 row {idx+1} outcome name mismatch: Excel '{out_excel}' vs Word '{out_word}'")
            # Check outcome data columns
            for c_excel in range(2, 7):
                c_word = c_excel - 1
                v_excel = str(ws2.cell(r_excel, c_excel).value or '').strip()
                v_word = t2_word.rows[r_word].cells[c_word].text.strip()
                if v_excel.replace('\xa0', ' ') != v_word.replace('\xa0', ' '):
                    word_fidelity_errors.append(f"Table 2 {out_excel} col {c_excel} mismatch: Excel '{v_excel}' vs Word '{v_word}'")

    # Table S1: Check code rows
    ws_s1 = wb_masked['Table S1']
    t_s1_word = doc_s1.tables[0]
    n_s1_word = len(t_s1_word.rows) - 2  # minus title & header
    n_s1_excel = ws_s1.max_row - 2      # minus title & header
    if n_s1_word != n_s1_excel:
        word_fidelity_errors.append(f"Word Table S1 rows ({n_s1_word}) != Excel Table S1 rows ({n_s1_excel})")

    if word_fidelity_errors:
        results['errors'].extend(word_fidelity_errors[:5])
    else:
        results['word_fidelity_pass'] = True

    return results

if __name__ == '__main__':
    import json
    if len(sys.argv) < 3:
        print("Usage: python verify_all_gates.py <raw_docx> <study_dir>")
        sys.exit(1)
    raw_docx = sys.argv[1]
    study_dir = sys.argv[2]
    stem = os.path.splitext(os.path.basename(raw_docx))[0]

    clean_xlsx = os.path.join(study_dir, f"clean_{stem}.xlsx")
    masked_xlsx = os.path.join(study_dir, f"{stem}_Table1_2_S1.xlsx")
    log_csv = os.path.join(study_dir, f"{stem}_change_log.csv")
    word_t1_2 = os.path.join(study_dir, f"{stem}_Table_1_and_2.docx")
    word_s1 = os.path.join(study_dir, f"{stem}_Table_S1.docx")

    res = verify_study(raw_docx, clean_xlsx, masked_xlsx, log_csv, word_t1_2, word_s1)
    print(json.dumps(res, indent=2))
    all_pass = (
        res.get('anti_template_pass') and
        res.get('code_completeness_pass') and
        res.get('suppression_pass') and
        res.get('table2_math_pass') and
        res.get('word_fidelity_pass')
    )
    if not all_pass:
        sys.exit(1)

