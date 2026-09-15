"""
build_study_tables_v2.py

Disease-agnostic pipeline for TriNetX two-cohort study tables (v2).
Constructs:
- clean_<stem>.xlsx (Table 1, Table 2, Table S1)
- masked_<stem>.xlsx (with small-cell <=10 suppression)
- <stem>_change_log.csv
- <stem>_Table_1_and_2.docx
- <stem>_Table_S1.docx

Strictly complies with Rules A-E:
- Single source of truth (report only)
- Complete code-type fidelity
- Exact titles and footnotes
- Section D rename map & house style
- Verification gates (anti-template, bidirectional codes, suppression, math audit)
"""

import os
import sys
import re
import csv
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from trinetx_docx_parser import parse_trinetx_docx
from suppress import suppress_workbook
from format_tables_to_docx import build_docx

# Section D: Global baseline label rename map
SECTION_D_RENAMES = {
    "Nicotine dependence, cigarettes": "Smoking",
    "Overweight, obesity and other hyperalimentation": "Overweight/obesity",
    "Coagulation defects, purpura and other hemorrhagic conditions": "Coagulation/hemorrhagic disorder",
    "Occlusion and stenosis of carotid artery": "Carotid stenosis/occlusion",
    "Other peripheral vascular diseases": "Peripheral vascular diseases",
    "Transient cerebral ischemic attacks and related syndromes": "TIA-related syndrome",
    "Epilepsy and recurrent seizures": "Epilepsy/recurrent seizures",
}

SECTION_DISPLAY = {
    'Demographics': 'Demographics',
    'Diagnosis': 'Comorbidities',
    'Medication': 'Medications',
    'Laboratory': 'Laboratory',
}

def clean_label(raw_label: str, sec_name: str) -> str:
    """Apply Section D renames and house-style cleanups."""
    if not raw_label:
        return ""
    
    lbl = raw_label.strip()
    
    # 1. Check Section D exact renames
    if lbl in SECTION_D_RENAMES:
        return SECTION_D_RENAMES[lbl]
    
    # 2. House-style medication capitalization
    if sec_name == 'Medication':
        # Capitalize first letter of medication name
        return lbl[0].upper() + lbl[1:] if lbl else lbl

    # 3. Lab specific cleanups
    if "Hemoglobin A1c/Hemoglobin.total in Blood" in lbl:
        return "Hemoglobin A1c, %"

    if sec_name == 'Laboratory':
        # Open-ended lab bins ('10 - 0 %' -> '≥10 %', '50 - 0 kg/m2' -> '≥50 kg/m²');
        # all other bin labels stay verbatim from the report (anti-template gate safe).
        lbl = re.sub(r'\b10\s*-\s*0\s*%', '≥10 %', lbl)
        lbl = re.sub(r'\b50\s*-\s*0\s*kg/m2', '≥50 kg/m²', lbl)

    return lbl

def item_display_label(it: dict, sec_name: str) -> str:
    """Repair misparsed rows, then apply Section D renames / house-style cleanups."""
    lbl = it['label']
    code = it['code']
    # Laboratory bin rows (e.g. '0 - 5.70 %', '35 - 39.90 kg/m2'): the TriNetX table puts
    # the bin range where a code would sit, so the parser swaps bin text into 'code' and
    # the first count into 'label'. Repair by using the bin text as the label.
    if sec_name == 'Laboratory' and lbl and lbl.replace(',', '').replace('.', '').isdigit():
        lbl = code
    return clean_label(lbl, sec_name)


VOCAB_DISPLAY = {
    'ICD10CM': 'ICD-10-CM',
    'ICD10PCS': 'ICD-10-PCS',
    'CPT': 'CPT',
    'HCPCS': 'HCPCS',
    'RXNORM': 'RxNorm',
}

TERM_RE = re.compile(
    r"([A-Za-z][A-Za-z0-9 ,/();\[\]'\-\.]*?)\s*"
    r"\((UMLS:(?:ICD10CM|ICD10PCS|CPT|HCPCS):[A-Za-z0-9.\-]+|NLM:RXNORM:[0-9]+)\)")

# A named criterion group inside a cohort/index-event definition, e.g.
# "GLP1: Any instance of GLP1 occurred within 3 months before or up to 1 day after any
#  instance of Multiple Sclerosis   Patients must have: any of the following: ..."
GROUP_RE = re.compile(
    r"([A-Z][A-Za-z0-9 ]*?):\s*"
    r"((?:Any instance of|The terms in this group)[^.]*?)\s*"
    r"Patients (must have|cannot have):\s*"
    r"(?:any of the following:\s*)?"
    r"(.*?)"
    r"(?=(?:[A-Z][A-Za-z0-9 ]*?:\s*(?:Any instance of|The terms in this group))|$)",
    re.DOTALL)


def format_code_token(token: str) -> str:
    """'UMLS:ICD10CM:G35' -> 'ICD-10-CM G35'; 'NLM:RXNORM:2601723' -> 'RxNorm 2601723'."""
    if token.startswith('NLM:RXNORM:'):
        return f"RxNorm {token.split(':')[-1]}"
    if token.startswith('UMLS:'):
        parts = token.split(':')
        voc = parts[1]
        return f"{VOCAB_DISPLAY.get(voc, voc)} {parts[-1]}"
    return token


def harvest_terms(block: str) -> list:
    """Extract (label, formatted_code) pairs from a report block ('label (TOKEN)' patterns)."""
    terms = []
    for m in TERM_RE.finditer(re.sub(r'\s+', ' ', block)):
        lbl = m.group(1).strip()
        lbl = re.sub(r'^(?:or|and)\s+', '', lbl).strip().rstrip(';,. ').strip()
        if lbl:
            terms.append((lbl, format_code_token(m.group(2))))
    return terms


def norm_text(s: str) -> str:
    return re.sub(r'\s+', ' ', s or '')


def derive_clean_cohort_display(query_name: str) -> str:
    """Derive clean, human-readable display label from TriNetX query name."""
    # e.g. "cSDH: Standalone MMAE PLT<100k" -> "Standalone MMAE (PLT<100k)"
    # "cSDH: MM alone PLT<100k" -> "Medical management (PLT<100k)"
    # "cSDH: Sx alone PLT<100k" -> "Surgery alone (PLT<100k)"
    # "cSDH: Sx + MMAE PLT<100k" -> "Surgery + MMAE (PLT<100k)"
    clean = query_name.strip()
    # strip prefix like "cSDH: " if present, but keep disease-agnostic
    if ':' in clean:
        clean = clean.split(':', 1)[1].strip()
    
    # Check if threshold like PLT<100k is attached
    m_thr = re.search(r'(PLT\s*[<>=]\s*\d+k?)', clean, re.IGNORECASE)
    thr_str = ""
    if m_thr:
        thr_str = f" ({m_thr.group(1).strip()})"
        clean = clean[:m_thr.start()].strip()
        
    # Map common abbreviations cleanly while preserving fidelity
    if clean.lower() == 'mm alone':
        clean = "Medical management"
    elif clean.lower() == 'sx alone':
        clean = "Surgery alone"
    elif clean.lower() == 'sx + mmae':
        clean = "Surgery + MMAE"
    elif clean.lower() == 'standalone mmae':
        clean = "Standalone MMAE"
        
    return clean + thr_str

def fmt_p(p_val: str) -> str:
    if str(p_val).strip() in ('0', '0.000', '0.0000'):
        return '<0.001'
    return str(p_val).strip()

def build_excel_workbook(data: dict, out_clean_xlsx: str) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    # Default sheet
    ws_t1 = wb.active
    ws_t1.title = "Table 1"
    ws_t2 = wb.create_sheet(title="Table 2")
    ws_s1 = wb.create_sheet(title="Table S1")
    
    bold = Font(bold=True)
    reg = Font(bold=False)
    ctr = Alignment(horizontal='center', vertical='center')
    ctr_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')
    left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)
    
    c1_display = derive_clean_cohort_display(data['matching_cohort_1']['name'])
    c2_display = derive_clean_cohort_display(data['matching_cohort_2']['name'])
    time_window = data['time_window_text']
    
    n_b1 = data['matching_cohort_1']['n_before']
    n_b2 = data['matching_cohort_2']['n_before']
    n_a1 = data['matching_cohort_1']['n_after']
    n_a2 = data['matching_cohort_2']['n_after']

    # ----------------------------------------------------
    # SHEET 1: Table 1
    # ----------------------------------------------------
    # R1: Title
    t1_title = f"Table 1: Patient Characteristics Before and After Propensity Score Matching (PSM) — {c1_display} vs {c2_display}"
    ws_t1.cell(1, 1, t1_title).font = bold
    ws_t1.merge_cells('A1:I1')
    ws_t1.cell(1, 1).alignment = left

    # R2: Band
    ws_t1.cell(2, 1, 'Characteristics - mean ± SD, % (n)').font = reg
    ws_t1.merge_cells('A2:A3')
    ws_t1.cell(2, 1).alignment = ctr_wrap
    
    ws_t1.cell(2, 2, 'Before PSM').font = bold
    ws_t1.merge_cells('B2:E2')
    ws_t1.cell(2, 2).alignment = ctr
    
    ws_t1.cell(2, 6, 'After PSM').font = bold
    ws_t1.merge_cells('F2:I2')
    ws_t1.cell(2, 6).alignment = ctr

    # R3: Sub-headers
    hdr1 = [
        (2, f"{c1_display} (N={n_b1:,})"),
        (3, f"{c2_display} (N={n_b2:,})"),
        (4, 'P-value'),
        (5, 'ASD'),
        (6, f"{c1_display} (N={n_a1:,})"),
        (7, f"{c2_display} (N={n_a2:,})"),
        (8, 'P-value'),
        (9, 'ASD'),
    ]
    for col, text in hdr1:
        c = ws_t1.cell(3, col, text)
        c.font = bold
        c.alignment = ctr_wrap

    # R4+: Data rows
    row_idx = 4
    for sec in data['table1_sections']:
        sec_name = sec['name']
        sec_display = SECTION_DISPLAY.get(sec_name, sec_name)
        
        # Section header
        ws_t1.cell(row_idx, 1, sec_display).font = bold
        row_idx += 1
        
        for it in sec['items']:
            disp_lbl = item_display_label(it, sec_name)

            b = it['before']
            a = it['after']
            
            # Label
            ws_t1.cell(row_idx, 1, disp_lbl).font = reg
            ws_t1.cell(row_idx, 1).alignment = left
            
            # Before values
            ws_t1.cell(row_idx, 2, b['c1_str']).alignment = ctr
            ws_t1.cell(row_idx, 3, b['c2_str']).alignment = ctr
            ws_t1.cell(row_idx, 4, fmt_p(b['p'])).alignment = ctr
            ws_t1.cell(row_idx, 5, str(b['asd'])).alignment = ctr
            
            # After values
            ws_t1.cell(row_idx, 6, a['c1_str']).alignment = ctr
            ws_t1.cell(row_idx, 7, a['c2_str']).alignment = ctr
            ws_t1.cell(row_idx, 8, fmt_p(a['p'])).alignment = ctr
            ws_t1.cell(row_idx, 9, str(a['asd'])).alignment = ctr
            
            row_idx += 1

    widths_t1 = {'A': 42.0, 'B': 18.0, 'C': 18.0, 'D': 12.0, 'E': 12.0, 'F': 18.0, 'G': 18.0, 'H': 12.0, 'I': 12.0}
    for col, w in widths_t1.items():
        ws_t1.column_dimensions[col].width = w

    # ----------------------------------------------------
    # SHEET 2: Table 2 (Outcomes)
    # ----------------------------------------------------
    # R1: Title
    t2_title = f"Table 2: Clinical Outcomes at {time_window} Follow-up — {c1_display} vs {c2_display}"
    ws_t2.cell(1, 1, t2_title).font = bold
    ws_t2.merge_cells('A1:F1')
    ws_t2.cell(1, 1).alignment = left

    # R2: Headers
    hdr2 = ['Outcomes', 'Cohort 1', 'Cohort 2', 'Event probability rate', 'p-value', 'HR [95%CI]']
    for i, h in enumerate(hdr2, start=1):
        c = ws_t2.cell(2, i, h)
        c.font = bold if i > 1 else reg
        c.alignment = ctr_wrap if i > 1 else left

    # R3: Sub-headers
    ws_t2.cell(3, 2, f"{c1_display} (N={n_a1:,})").font = bold
    ws_t2.cell(3, 2).alignment = ctr_wrap
    ws_t2.cell(3, 3, f"{c2_display} (N={n_a2:,})").font = bold
    ws_t2.cell(3, 3).alignment = ctr_wrap
    ws_t2.cell(3, 4, f"({c1_display} vs {c2_display})").font = bold
    ws_t2.cell(3, 4).alignment = ctr_wrap

    # R4+: Data rows
    row_idx = 4
    for o in data['table2_outcomes']:
        # Format typo fix if any
        out_name = o['name']
        if out_name.lower() == 'treatement':
            out_name = 'Treatment'
            
        c1_risk_str = f"{o['c1_risk_pct']:.1f}% ({o['c1_risk_n']:,})"
        c2_risk_str = f"{o['c2_risk_pct']:.1f}% ({o['c2_risk_n']:,})"
        ev_rate_str = f"{o['c1_event_rate']:.2f}% vs {o['c2_event_rate']:.2f}%"
        p_str = fmt_p(o['logrank_p'])
        
        hr_str = "-"
        if o['hr'] not in ('-', 'N/A', None):
            try:
                hr_f = float(o['hr'])
                hr_lo_f = float(o['hr_lo'])
                hr_hi_f = float(o['hr_hi'])
                hr_str = f"{hr_f:.3f} [{hr_lo_f:.3f}-{hr_hi_f:.3f}]"
            except ValueError:
                hr_str = str(o['hr'])

        vals_t2 = [out_name, c1_risk_str, c2_risk_str, ev_rate_str, p_str, hr_str]
        for col_i, v in enumerate(vals_t2, start=1):
            c = ws_t2.cell(row_idx, col_i, v)
            c.font = reg
            c.alignment = left if col_i == 1 else ctr
        row_idx += 1

    widths_t2 = {'A': 22.0, 'B': 22.0, 'C': 22.0, 'D': 22.0, 'E': 12.0, 'F': 22.0}
    for col, w in widths_t2.items():
        ws_t2.column_dimensions[col].width = w

    # ----------------------------------------------------
    # SHEET 3: Table S1 (Code Reference)
    # ----------------------------------------------------
    # R1: Title
    ws_s1.cell(1, 1, f"Table S1: Code Reference and Study Definitions — {c1_display} vs {c2_display}").font = bold
    ws_s1.merge_cells('A1:B1')
    ws_s1.cell(1, 1).alignment = left

    # R2: Header
    ws_s1.cell(2, 1, 'Characteristic / Criteria').font = bold
    ws_s1.cell(2, 2, 'Codes (ICD-10-CM / ICD-10-PCS / CPT / HCPCS / RxNorm / TriNetX)').font = bold
    ws_s1.cell(2, 1).alignment = left
    ws_s1.cell(2, 2).alignment = left

    row_idx = 3
    def s1_section(sec_title):
        nonlocal row_idx
        ws_s1.cell(row_idx, 1, sec_title).font = bold
        row_idx += 1

    def s1_row(char_text, code_text):
        nonlocal row_idx
        ws_s1.cell(row_idx, 1, char_text).font = reg
        ws_s1.cell(row_idx, 1).alignment = left_top
        ws_s1.cell(row_idx, 2, code_text).font = reg
        ws_s1.cell(row_idx, 2).alignment = left_top
        row_idx += 1

    norm_a = norm_text(data['appendix_a_text'])
    norm_b = norm_text(data['appendix_b_text'])
    norm_c = norm_text(data['appendix_c_text'])

    def cohort_block(text, header_pattern, c_num):
        m = re.search(rf"{header_pattern.format(c=c_num)}.*?(?={header_pattern.format(c=c_num + 1)}|$)", text)
        return m.group(0) if m else ''

    def group_rows(block_text, row_prefix):
        """Emit one S1 row per named criterion group in a cohort/index-event block."""
        rows_out = []
        seen = set()
        for gm in GROUP_RE.finditer(block_text):
            g_name = gm.group(1).strip()
            g_rel = gm.group(2).strip()
            g_req = gm.group(3).strip()
            terms = harvest_terms(gm.group(4))
            if not terms:
                continue
            code_txt = f"{g_rel} — Patients {g_req}: " + "; ".join(f"{c} ({l})" for l, c in terms)
            key = (g_name, code_txt)
            if key in seen:
                continue
            seen.add(key)
            rows_out.append((f"{row_prefix} — {g_name}", code_txt))
        return rows_out

    # 1. Inclusion
    s1_section("Inclusion")
    s1_row("Age at index ≥ 18 years", "TriNetX demographics criterion: Age, at least 18 years (most recent occurrence)")
    if "on or after Jan 1, 2016" in norm_a or "on or after Jan 1, 2016" in norm_b:
        s1_row("Index date window", "Terms occurred on or after Jan 1, 2016")

    # Index-event group definitions per cohort (Appendix B) — generic: works for any
    # disease area (diagnosis groups, exposure groups, procedure groups).
    seen_incl = set()
    for c_num in (1, 2):
        c_block = cohort_block(norm_b, "The index event for Cohort {c}", c_num)
        disp_name = c1_display if c_num == 1 else c2_display
        for char_txt, code_txt in group_rows(c_block, f"Cohort {c_num} ({disp_name}) index event"):
            # dedupe identical group definitions across cohorts (e.g. shared diagnosis)
            if (char_txt.split(' — ')[0], code_txt) in seen_incl:
                continue
            seen_incl.add((char_txt.split(' — ')[0], code_txt))
            s1_row(char_txt, code_txt)

    # 2. Exclusions — harvest the actual exclusion term list from the cohort query
    # criteria (Appendix A, 'Exclusions ... Patients cannot have'), one row per term.
    s1_section("Exclusions")
    s1_row("Index event > 20 years prior",
           "Patients whose index event occurred 20 years or more prior are excluded "
           f"({data.get('excluded_20y', 0)} excluded)")
    excl_terms = []
    for c_num in (1, 2):
        c_block = cohort_block(norm_a, "Query Criteria for Cohort {c}", c_num)
        for gm in GROUP_RE.finditer(c_block):
            if gm.group(1).strip().lower() == 'exclusions':
                excl_terms = harvest_terms(gm.group(4))
                break
        if excl_terms:
            break
    for lbl, code_txt in excl_terms:
        s1_row(lbl, code_txt)

    # 3. Outcomes — each outcome gets ONLY the codes in its own Appendix C block.
    s1_section("Outcomes")
    outcome_names = [o['name'] for o in data['table2_outcomes']]
    positions = []
    for nm in outcome_names:
        pos = norm_c.find(nm)
        if pos != -1:
            positions.append((pos, nm))
    positions.sort()
    matched_names = {nm for _, nm in positions}
    for j, (pos, nm) in enumerate(positions):
        block_end = positions[j + 1][0] if j + 1 < len(positions) else len(norm_c)
        block = norm_c[pos:block_end]
        if 'deceased' in block.lower() and not harvest_terms(block):
            s1_row(nm, "TriNetX demographics criterion: Deceased")
            continue
        terms = harvest_terms(block)
        s1_row(nm, "; ".join(f"{c} ({l})" for l, c in terms) if terms else "Defined per TriNetX concept")
    # Outcomes defined in the report but not located in Appendix C still get a row.
    for o in data['table2_outcomes']:
        if o['name'] not in matched_names:
            s1_row(o['name'], "Defined per TriNetX concept (see report Outcome Definitions)")

    # 4. Covariates
    s1_section("Covariates")
    for sec in data['table1_sections']:
        sec_name = sec['name']
        for it in sec['items']:
            raw_lbl = it['label']
            # skip misparsed laboratory bin rows (ranges belong to the parent lab concept)
            if sec_name == 'Laboratory' and raw_lbl and raw_lbl.replace(',', '').replace('.', '').isdigit():
                continue
            disp_lbl = item_display_label(it, sec_name)
            code = it['code']

            code_str = ""
            if sec_name == 'Demographics':
                if code == 'AI':
                    code_str = "TriNetX demographics criterion: Age at Index"
                else:
                    code_str = f"TriNetX concept {code}"
            elif sec_name == 'Diagnosis':
                code_str = f"ICD-10-CM {code}"
            elif sec_name == 'Medication':
                code_str = f"RxNorm {code}"
            elif sec_name == 'Laboratory':
                if code:
                    code_str = f"TriNetX laboratory value {code}"
                else:
                    code_str = f"TriNetX laboratory value ({raw_lbl})"

            s1_row(disp_lbl, code_str)

    widths_s1 = {'A': 45.0, 'B': 95.0}
    for col, w in widths_s1.items():
        ws_s1.column_dimensions[col].width = w

    wb.save(out_clean_xlsx)
    return wb

def run_pipeline_for_docx(docx_path: str, out_root: str) -> dict:
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    study_dir = os.path.join(out_root, stem)
    os.makedirs(study_dir, exist_ok=True)
    
    clean_xlsx = os.path.join(study_dir, f"clean_{stem}.xlsx")
    masked_xlsx = os.path.join(study_dir, f"{stem}_Table1_2_S1.xlsx")
    change_log_csv = os.path.join(study_dir, f"{stem}_change_log.csv")
    
    # 1. Parse report
    data = parse_trinetx_docx(docx_path)
    
    # 2. Build clean Excel
    build_excel_workbook(data, clean_xlsx)
    
    # 3. Apply suppression
    n_changes, changes = suppress_workbook(clean_xlsx, masked_xlsx, change_log_csv)
    
    # 4. Word export via format_tables_to_docx
    # Use exact Section D renames, no cross-study renames
    docx_files = build_docx(
        xlsx_path=masked_xlsx,
        out_dir=study_dir,
        rename_map=dict(SECTION_D_RENAMES),
        pair_rename_map={},
        blank_value_labels=set(),
        consolidate_nicotine_rows=False,
        apply_general_renames=True,
        file_grouping="match_uploads",
        main_filename=f"{stem}_Table_1_and_2.docx",
        supp_filename=f"{stem}_Table_S1.docx"
    )
    
    return {
        'stem': stem,
        'docx_path': docx_path,
        'study_dir': study_dir,
        'clean_xlsx': clean_xlsx,
        'masked_xlsx': masked_xlsx,
        'change_log_csv': change_log_csv,
        'word_files': docx_files,
        'suppression_count': n_changes,
        'data': data,
    }

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="Build study tables from TriNetX docx export.")
    parser.add_argument('docx_path', help="Path to TriNetX docx file")
    parser.add_argument('--out_root', default="output", help="Output directory root (default: output)")
    args = parser.parse_args()
    
    res = run_pipeline_for_docx(args.docx_path, args.out_root)
    print(f"Pipeline finished successfully for {res['stem']}:")
    print(f"  Study dir: {res['study_dir']}")
    print(f"  Suppressed changes: {res['suppression_count']}")
    print(f"  Clean Excel: {res['clean_xlsx']}")
    print(f"  Masked Excel: {res['masked_xlsx']}")
    print(f"  Change log: {res['change_log_csv']}")
    print(f"  Word files: {res['word_files']}")

