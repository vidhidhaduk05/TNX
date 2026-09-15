"""
Build publication-quality clinical study workbooks from TriNetX cSDH PDF exports.

Produces ONE workbook per PDF with THREE sheets (locked skill requirement):
    Table 1  -- baseline characteristics (before/after PSM), 9 columns
    Table 2  -- study outcomes, 6 columns (1 block, matching example)
    Table S1 -- TriNetX cohort & outcome code reference, 2 columns

Formatting reproduces the user's example workbooks EXACTLY:
    - Table 1: R1 title (merged A1:I1); R2 band Before/After PSM; R3 sub-headers
      C1(N=)|C2(N=)|P-value|ASD x2; section rows (col A bold); data as
      '%(n)' (categorical) or 'mean +/- SD' (continuous); missing -> '-'.
    - Table 2: R1 title; R2 headers; R3 cohort-label/N + comparison sub-header;
      one outcome block (2 rows for this study).
    - Table S1: R1 title; R2 header; section headers (Inclusion/Outcomes/Covariates);
      code rows with system tag literal from UMLS prefix.

Clean labels come from labels.py (verbatim PDF text; NOT the garbled extractor label).

This module builds the CLEAN workbook. The <=10 suppression pass (Table 1 only)
is applied afterward by suppress.py to produce the delivered masked workbook.
"""
import sys
sys.path.insert(0, '/workspace')

import re
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from trinetx_coord import extract_pdf_sections
from trinetx_outcomes import parse_outcomes
from labels import LABELS, code_system, pretty_bin

# ---------------------------------------------------------------------------
# Study configuration -- all 5 PDFs
# ---------------------------------------------------------------------------
# Each study carries everything needed to build source-faithful tables:
#   pdf         : source PDF path
#   stem        : exact original filename stem (used for output naming)
#   c1 / c2     : cohort display labels (user-confirmed style, threshold per file)
#   thr_k       : platelet threshold in units of 1000/uL (50 or 100) -> title text
#   thr_val     : the numeric PLT inclusion cutoff text as printed ("50.00"/"100.00")
#   c1_kind     : 'mmae' (standalone MMAE) or 'sx_mmae' (surgery + MMAE)
#   c2_kind     : 'surgery' | 'medical' | 'surgery'  (drives S1 cohort-definition wording)
STUDIES = {
    'p50_Sx': {
        'pdf': '/mnt/user-uploads/PLT__50__Sx_vs_MMAE_202607261450.pdf',
        'stem': 'PLT__50__Sx_vs_MMAE_202607261450',
        'c1': 'Standalone MMAE (PLT<50k)', 'c2': 'Surgery alone (PLT<50k)',
        'thr_k': 50, 'thr_val': '50.00', 'c1_kind': 'mmae', 'c2_kind': 'surgery',
    },
    'p50_MM': {
        'pdf': '/mnt/user-uploads/PLT__50__MM_vs_MMAE_202607261435.pdf',
        'stem': 'PLT__50__MM_vs_MMAE_202607261435',
        'c1': 'Standalone MMAE (PLT<50k)', 'c2': 'Medical management (PLT<50k)',
        'thr_k': 50, 'thr_val': '50.00', 'c1_kind': 'mmae', 'c2_kind': 'medical',
    },
    'p100_MM': {
        'pdf': '/mnt/user-uploads/PLT__100__MM_vs_MMAE_202607261353.pdf',
        'stem': 'PLT__100__MM_vs_MMAE_202607261353',
        'c1': 'Standalone MMAE (PLT<100k)', 'c2': 'Medical management (PLT<100k)',
        'thr_k': 100, 'thr_val': '100.00', 'c1_kind': 'mmae', 'c2_kind': 'medical',
    },
    'p100_Sx': {
        'pdf': '/mnt/user-uploads/PLT__100__Sx_vs_MMAE_202607261455.pdf',
        'stem': 'PLT__100__Sx_vs_MMAE_202607261455',
        'c1': 'Standalone MMAE (PLT<100k)', 'c2': 'Surgery alone (PLT<100k)',
        'thr_k': 100, 'thr_val': '100.00', 'c1_kind': 'mmae', 'c2_kind': 'surgery',
    },
    'p100_SxM': {
        'pdf': '/mnt/user-uploads/PLT__100__Sx_vs_Sx_+_MMAE_202607261402.pdf',
        'stem': 'PLT__100__Sx_vs_Sx_+_MMAE_202607261402',
        'c1': 'Surgery + MMAE (PLT<100k)', 'c2': 'Surgery alone (PLT<100k)',
        'thr_k': 100, 'thr_val': '100.00', 'c1_kind': 'sx_mmae', 'c2_kind': 'surgery',
    },
}

# Back-compat aliases used throughout the builders.
PDFS = {k: v['pdf'] for k, v in STUDIES.items()}
COHORTS = {k: {'c1': v['c1'], 'c2': v['c2']} for k, v in STUDIES.items()}

# Section display names (match example: 'Comorbidities'/'Medications' plural;
# study adds 'Laboratory'). Extractor section -> display header.
SECTION_DISPLAY = {
    'Demographics': 'Demographics',
    'Diagnosis': 'Comorbidities',
    'Medication': 'Medications',
    'Laboratory': 'Laboratory',
}
SECTION_ORDER = ['Demographics', 'Diagnosis', 'Medication', 'Laboratory']

MISSING = '-'   # not-calculable / missing sentinel (example uses '-')

# ---------------------------------------------------------------------------
# Value coercion helpers (extractor returns strings; %-suffixed pcts)
# ---------------------------------------------------------------------------
def _num(x):
    """Coerce a possibly-string, possibly-%-suffixed, comma-bearing value to
    float. Returns None for missing/'--'/non-numeric."""
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(',', '').rstrip('%').strip()
    if s in ('', '--', '-'):
        return None
    try:
        return float(s)
    except ValueError:
        return None

def _int(x):
    """Coerce count value (string/comma) to int, else None."""
    if x is None:
        return None
    if isinstance(x, int):
        return x
    s = str(x).strip().replace(',', '')
    if s in ('', '--', '-'):
        return None
    try:
        return int(round(float(s)))
    except ValueError:
        return None

def parse_cohort_ns(title):
    """Extract (n1, n2) integers from a characteristics-table title like
    'Cohort 1 (N = 152) and cohort 2 (N = 425) characteristics ...'."""
    nums = re.findall(r'N\s*=\s*([\d,]+)', title)
    if len(nums) < 2:
        raise ValueError(f'could not parse two Ns from title: {title!r}')
    return int(nums[0].replace(',', '')), int(nums[1].replace(',', ''))

# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
def fmt_p(p):
    """p=0.000 -> '<0.001'; '--'/None -> '-'; TriNetX bare '1' -> '1.000'; else 3dp."""
    if p in (None, '--', '', '-'):
        return MISSING
    try:
        v = float(p)
    except (TypeError, ValueError):
        return str(p)
    if v == 0:
        return '<0.001'
    return f'{v:.3f}'          # 1 -> '1.000', 0.011 -> '0.011'

def fmt_sd(s):
    """Std diff: '--'/None -> '-'; '<0.001' passthrough; else 3dp."""
    if s in (None, '--', '', '-'):
        return MISSING
    if isinstance(s, str) and s.strip() == '<0.001':
        return '<0.001'
    try:
        return f'{float(s):.3f}'
    except (TypeError, ValueError):
        return str(s)

def fmt_pct_n(pct, count):
    """Categorical cell '52.4% (33,586)'. Accepts string/%-suffixed/comma inputs.
    Count rendered with thousands separators to match example."""
    p = _num(pct)
    c = _int(count)
    if p is None or c is None:
        return MISSING
    # percent: keep the PDF's own precision feel -> 1 dp like example
    return f'{p:.1f}% ({c:,})'

def fmt_mean_sd(mean, sd):
    """Continuous cell 'mean +/- SD' with unicode +/-. 1 dp like example."""
    m = _num(mean)
    s = _num(sd)
    if m is None or s is None:
        return MISSING
    return f'{m:.1f} \u00b1 {s:.1f}'

def row_label(r):
    """Clean display label for a characteristics row."""
    if r['code'] is None:            # laboratory range-bin
        return pretty_bin(r['label'])
    return LABELS.get(r['code'], r['label'])

# ---------------------------------------------------------------------------
# Table 1 builder
# ---------------------------------------------------------------------------
def _cat_cells(r):
    """Return (c1_cell, c2_cell) '%(n)' strings for a categorical row."""
    return (fmt_pct_n(r.get('c1_pct'), r.get('c1_count')),
            fmt_pct_n(r.get('c2_pct'), r.get('c2_count')))

def build_table1(ws, study, char_tables):
    """Assemble the Table 1 sheet on worksheet ws."""
    before = char_tables[study]['before']
    after = char_tables[study]['after']
    coh = COHORTS[study]

    # cohort N per phase (parsed from table titles)
    n_b1, n_b2 = parse_cohort_ns(before['title'])
    n_a1, n_a2 = parse_cohort_ns(after['title'])

    bold = Font(bold=True)
    reg = Font(bold=False)
    ctr = Alignment(horizontal='center')
    ctr_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left')

    # R1 title
    title = (f'Table 1: Patient Characteristics Before and After Propensity Score '
             f'Matching (PSM) \u2014 {coh["c1"]} vs {coh["c2"]}')
    ws.cell(1, 1, title).font = bold
    ws.merge_cells('A1:I1')
    ws.cell(1, 1).alignment = left

    # R2 band
    ws.cell(2, 1, 'Characteristics - mean \u00b1 SD, % (n)').font = reg
    ws.merge_cells('A2:A3')
    ws.cell(2, 1).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.cell(2, 2, 'Before PSM').font = bold
    ws.merge_cells('B2:E2'); ws.cell(2, 2).alignment = ctr
    ws.cell(2, 6, 'After PSM').font = bold
    ws.merge_cells('F2:I2'); ws.cell(2, 6).alignment = ctr

    # R3 sub-headers
    hdr = [
        (2, f'{coh["c1"]} (N={n_b1:,})'),
        (3, f'{coh["c2"]} (N={n_b2:,})'),
        (4, 'P-value'),
        (5, 'ASD'),
        (6, f'{coh["c1"]} (N={n_a1:,})'),
        (7, f'{coh["c2"]} (N={n_a2:,})'),
        (8, 'P-value'),
        (9, 'ASD'),
    ]
    for col, text in hdr:
        c = ws.cell(3, col, text)
        c.font = bold
        c.alignment = ctr_wrap

    # Data rows, grouped by section
    row = 4
    # index rows by code within each phase for aligned lookup
    b_by_code = {r['code']: r for r in before['rows'] if r['code'] is not None}
    a_by_code = {r['code']: r for r in after['rows'] if r['code'] is not None}
    # lab bins keyed by cleaned label
    b_by_bin = {r['label']: r for r in before['rows'] if r['code'] is None}
    a_by_bin = {r['label']: r for r in after['rows'] if r['code'] is None}

    for sec in SECTION_ORDER:
        # section header row (col A bold only)
        ws.cell(row, 1, SECTION_DISPLAY[sec]).font = bold
        row += 1
        # rows belonging to this section, in the extractor's order
        sec_rows = [r for r in before['rows'] if r['section'] == sec]
        for br in sec_rows:
            code = br['code']
            if code is not None:
                ar = a_by_code.get(code, {})
            else:
                ar = a_by_bin.get(br['label'], {})
            label = row_label(br)
            ws.cell(row, 1, label).font = reg
            ws.cell(row, 1).alignment = left

            if br.get('is_mean'):
                # continuous row: mean +/- SD, both phases
                ws.cell(row, 2, fmt_mean_sd(br.get('c1_mean'), br.get('c1_sd')))
                ws.cell(row, 3, fmt_mean_sd(br.get('c2_mean'), br.get('c2_sd')))
                ws.cell(row, 6, fmt_mean_sd(ar.get('c1_mean'), ar.get('c1_sd')))
                ws.cell(row, 7, fmt_mean_sd(ar.get('c2_mean'), ar.get('c2_sd')))
            else:
                c1b, c2b = _cat_cells(br)
                ws.cell(row, 2, c1b)
                ws.cell(row, 3, c2b)
                if ar:
                    c1a, c2a = _cat_cells(ar)
                    ws.cell(row, 6, c1a)
                    ws.cell(row, 7, c2a)
                else:
                    ws.cell(row, 6, MISSING); ws.cell(row, 7, MISSING)

            # P-value + ASD both phases
            ws.cell(row, 4, fmt_p(br.get('p')))
            ws.cell(row, 5, fmt_sd(br.get('sd_stat')))
            ws.cell(row, 8, fmt_p(ar.get('p')) if ar else MISSING)
            ws.cell(row, 9, fmt_sd(ar.get('sd_stat')) if ar else MISSING)
            # center the numeric data columns
            for col in (2, 3, 4, 5, 6, 7, 8, 9):
                ws.cell(row, col).alignment = ctr
            row += 1

    # (Footnotes intentionally omitted per user decision -- both Excel and Word
    #  carry clean tables with no explanatory notes.)

    # column widths (match example)
    widths = {'A': 48.3, 'B': 20.5, 'C': 20.5, 'D': 9.0, 'E': 7.5,
              'F': 20.5, 'G': 20.5, 'H': 9.0, 'I': 7.5}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    return row  # first empty row after table


# ---------------------------------------------------------------------------
# Table 2 builder
# ---------------------------------------------------------------------------
def build_table2(ws, study, outcomes, after_n):
    """Assemble Table 2 sheet. after_n = (n1, n2) matched cohort sizes."""
    coh = COHORTS[study]
    n1, n2 = after_n
    bold = Font(bold=True); reg = Font(bold=False)
    ctr = Alignment(horizontal='center')
    ctr_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left')

    # R1 title
    ws.cell(1, 1, 'Table 2: Outcomes').font = bold
    ws.merge_cells('A1:F1'); ws.cell(1, 1).alignment = left

    # R2 headers
    headers = ['Outcomes', 'Cohort 1', 'Cohort 2',
               'Event probability rate', 'p-value', 'HR [95%CI]']
    for i, h in enumerate(headers, start=1):
        c = ws.cell(2, i, h)
        if i == 1:
            c.font = reg; c.alignment = left
        else:
            c.font = bold; c.alignment = ctr_wrap

    # R3 sub-header: cohort labels + N; comparison label under Event col
    ws.cell(3, 2, f'{coh["c1"]} (N={n1:,})').font = bold
    ws.cell(3, 2).alignment = ctr_wrap
    ws.cell(3, 3, f'{coh["c2"]} (N={n2:,})').font = bold
    ws.cell(3, 3).alignment = ctr_wrap
    ws.cell(3, 4, f'({coh["c1"]} vs {coh["c2"]})').font = bold
    ws.cell(3, 4).alignment = ctr_wrap

    # outcome rows (one block for this study)
    row = 4
    for o in outcomes:
        c1_pct = float(o['c1_risk']) * 100
        c2_pct = float(o['c2_risk']) * 100
        c1_ev = 100.0 - float(o['c1_km_surv'])
        c2_ev = 100.0 - float(o['c2_km_surv'])
        vals = [
            o['outcome'],
            f"{c1_pct:.1f}% ({int(o['c1_n_outcome']):,})",
            f"{c2_pct:.1f}% ({int(o['c2_n_outcome']):,})",
            f"{c1_ev:.2f}% vs {c2_ev:.2f}%",
            fmt_p(o['logrank_p']),
            f"{float(o['hr']):.3f} [{float(o['hr_lo']):.3f}-{float(o['hr_hi']):.3f}]",
        ]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(row, i, v)
            c.font = reg
            c.alignment = left if i == 1 else ctr
        row += 1

    widths = {'A': 20.2, 'B': 22.0, 'C': 22.0, 'D': 21.0, 'E': 11.0, 'F': 21.5}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    return row


# ---------------------------------------------------------------------------
# Table S1 builder
# ---------------------------------------------------------------------------
# MMAE procedures (ICD-10-PCS) -- shared by both PDFs
MMAE_PCS = ['03LG3BZ', '03LG3CZ', '03LG3DZ', '03LG3ZZ',
            '03VG3BZ', '03VG3CZ', '03VG3DZ', '03VG3HZ', '03VG3ZZ']
# Surgery procedures (ICD-10-PCS)
SURG_PCS = ['00N0', '00N1', '00N2', '00N7', '00C4', '0094', '00D2']
# SDH inclusion diagnosis codes
SDH_DX = [('I62.00', 'Nontraumatic subdural hemorrhage, unspecified'),
          ('I62.02', 'Nontraumatic subacute subdural hemorrhage'),
          ('I62.03', 'Nontraumatic chronic subdural hemorrhage')]

def _sys_tag(code):
    """Prefix a code with its display system tag (Table S1 style)."""
    s = code_system(code)
    if s == 'ICD-10':
        return f'ICD-10 {code}'
    if s == 'ICD-10-PCS':
        return f'ICD-10-PCS {code}'
    if s == 'RxNorm':
        return f'RxNorm {code}'
    if s == 'TNX':
        return f'TNX:{code}' if code != 'AI' else 'A1'
    return code

def build_table_s1(ws, study, char_tables, outcomes):
    """Assemble Table S1 code-reference sheet.

    Cohort-definition wording and the platelet threshold are taken from the
    per-study STUDIES config so each PDF's S1 is source-faithful. The Outcomes
    section is driven by the parsed outcome names (Mortality + Surgery+Mortality
    for PLT<50; Mortality + Surgery for PLT<100)."""
    cfg = STUDIES[study]
    coh = COHORTS[study]
    thr_k = cfg['thr_k']; thr_val = cfg['thr_val']
    bold = Font(bold=True); reg = Font(bold=False)
    left = Alignment(horizontal='left', vertical='top', wrap_text=True)
    left_nowrap = Alignment(horizontal='left')

    # R1 title
    ws.cell(1, 1, 'Table S1: Codes used in this study').font = reg
    ws.merge_cells('A1:B1'); ws.cell(1, 1).alignment = left_nowrap
    # R2 header
    ws.cell(2, 1, 'Characteristic').font = bold
    ws.cell(2, 2, 'Codes (ICD-10 / ICD-10-PCS / RxNorm / TriNetX)').font = bold
    ws.cell(2, 2).alignment = left_nowrap

    row = 3
    def section(name):
        nonlocal row
        ws.cell(row, 1, name).font = bold
        row += 1
    def entry(char, codes):
        nonlocal row
        ws.cell(row, 1, char).font = reg
        ws.cell(row, 1).alignment = left
        ws.cell(row, 2, codes).font = reg
        ws.cell(row, 2).alignment = left
        row += 1

    mmae_codes = ', '.join(f'ICD-10-PCS {c}' for c in MMAE_PCS)
    surg_codes = ', '.join(f'ICD-10-PCS {c}' for c in SURG_PCS)

    # ---- Inclusion ----
    section('Inclusion')
    entry('Age at index \u2265 18 years', 'A1 (Age at Index) \u2265 18')
    sdh_codes = ', '.join(f'ICD-10 {c} ({lbl})' for c, lbl in SDH_DX)
    entry('Chronic subdural hematoma (SDH) diagnosis (any of)', sdh_codes)
    entry(f'Platelet count \u2264 {thr_k} \u00d710\u00b3/\u00b5L',
          f'TNX:9020 (Platelets [#/volume] in Blood) \u2264 {thr_val} 10*3/uL')
    entry('Index date on or after Jan 1, 2016', 'Encounter date \u2265 2016-01-01')

    # ---- cohort definitions (per-study, source-faithful) ----
    # C1
    if cfg['c1_kind'] == 'sx_mmae':
        entry(f'Cohort 1 \u2014 {coh["c1"]}',
              f'Having surgical evacuation [{surg_codes}] AND MMAE [{mmae_codes}]')
    else:  # standalone MMAE
        entry(f'Cohort 1 \u2014 {coh["c1"]}',
              f'Having MMAE [{mmae_codes}]; NOT having surgical evacuation '
              f'[{surg_codes}]')
    # C2
    if cfg['c2_kind'] == 'surgery':
        # if C1 is surgery+MMAE, C2 (surgery alone) is distinguished by NOT having MMAE
        entry(f'Cohort 2 \u2014 {coh["c2"]}',
              f'Having surgical evacuation [{surg_codes}]; NOT having MMAE '
              f'[{mmae_codes}]')
    else:  # medical management
        entry(f'Cohort 2 \u2014 {coh["c2"]}',
              f'NOT having MMAE [{mmae_codes}]; NOT having surgical evacuation '
              f'[{surg_codes}] (i.e., medical management)')

    # ---- Outcomes (source-faithful: names from the parsed outcomes) ----
    section('Outcomes')
    outcome_names = [o['outcome'] for o in outcomes[study]]
    for oname in outcome_names:
        key = oname.lower().replace(' ', '').replace('+', '')
        if key == 'mortality':
            entry('Mortality', 'Deceased')
        elif key == 'surgerymortality':
            entry('Surgery + Mortality',
                  f'Any surgical evacuation [{surg_codes}] OR Deceased')
        elif key == 'surgery':
            entry('Surgery', f'Any surgical evacuation [{surg_codes}]')
        else:
            entry(oname, '')  # unexpected -- leave codes blank rather than fabricate

    # ---- Covariates ----
    section('Covariates')
    # use the before-PSM row order (identical to after) for covariates
    before = char_tables[study]['before']
    for r in before['rows']:
        code = r['code']
        if code is None:
            continue  # lab range-bins are derived from the platelet lab (below)
        if code == 'AI':
            entry(LABELS['AI'], 'A1')
        elif code == 'F':
            entry(LABELS['F'], 'F (Female)')
        elif code == '9020':
            entry(LABELS['9020'], 'TNX:9020')
        else:
            entry(LABELS.get(code, code), _sys_tag(code))

    # column widths
    ws.column_dimensions['A'].width = 42.0
    ws.column_dimensions['B'].width = 72.0
    return row


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def load_all():
    char_tables, outcomes = {}, {}
    for study, path in PDFS.items():
        tabs = extract_pdf_sections(path)
        char_tables[study] = {t['phase']: t for t in tabs}
        outcomes[study] = parse_outcomes(path)
    return char_tables, outcomes

def build_workbook(study, char_tables, outcomes):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = 'Table 1'
    ws2 = wb.create_sheet('Table 2')
    ws3 = wb.create_sheet('Table S1')

    build_table1(ws1, study, char_tables)
    after = char_tables[study]['after']
    n_a1, n_a2 = parse_cohort_ns(after['title'])
    build_table2(ws2, study, outcomes[study], (n_a1, n_a2))
    build_table_s1(ws3, study, char_tables, outcomes)
    return wb


if __name__ == '__main__':
    char_tables, outcomes = load_all()
    for study in PDFS:
        wb = build_workbook(study, char_tables, outcomes)
        out = f'/workspace/clean_{study}.xlsx'
        wb.save(out)
        print(f'saved {out} | sheets={wb.sheetnames}')
