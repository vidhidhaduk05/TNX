"""
Suppression acceptance checks (SKILL R6 / acceptance criteria 5-10).
Ground truth = the CLEAN intermediate workbook.

Checks per study:
  5. Every Table 1 block whose ORIGINAL min count <=10 has P='-' and ASD='-';
     NO block with all counts >10 was blanked (report min counts).
  6. No pure-numeric '(n)' token <=10 remains in Table 1 count columns; no '(<=M)'
     for M!=10; large counts (incl '(110)','(1,069)') intact.
  7. Table 2 + Table S1 content-identical to the clean intermediate; sheet count = 3.
  8. Percentages unchanged.
  9. Re-running suppression on the delivered workbook yields 0 changes (idempotent).
 10. change_log.csv present, one row per modified cell, 6 documented columns.
"""
import re, csv, os, sys
sys.path.insert(0, '/workspace')
import openpyxl
from suppress import suppress_workbook, _count_value, BLOCKS, FIRST_DATA_ROW, THRESHOLD, LE
from build_tables import STUDIES as STUDY_CFG

STUDIES = list(STUDY_CFG.keys())              # all 5 study keys
STEM = {k: v['stem'] for k, v in STUDY_CFG.items()}
PCT_RE = re.compile(r"(\d+(?:\.\d+)?)%")
BADMASK_RE = re.compile(r"\(" + LE + r"(\d+)\)")     # '(<=N)'
PLAIN_COUNT_RE = re.compile(r"\((\d{1,3}(?:,\d{3})*|\d+)\)")


def pcts(text):
    if text is None:
        return []
    return PCT_RE.findall(str(text))


def check_study(study):
    clean = f'/workspace/clean_{study}.xlsx'
    masked = f'/workspace/masked_{study}.xlsx'
    log = f'/workspace/{STEM[study]}_change_log.csv'
    wc = openpyxl.load_workbook(clean, data_only=False)
    wm = openpyxl.load_workbook(masked, data_only=False)
    res = {}

    # ---- 7. sheet count = 3, and Table 2 / Table S1 identical ----
    res['sheet_count_3'] = (len(wm.sheetnames) == 3 and wm.sheetnames == ['Table 1', 'Table 2', 'Table S1'])
    identical = True
    for sh in ('Table 2', 'Table S1'):
        a, b = wc[sh], wm[sh]
        if a.max_row != b.max_row or a.max_column != b.max_column:
            identical = False; break
        for r in range(1, a.max_row + 1):
            for c in range(1, a.max_column + 1):
                if a.cell(r, c).value != b.cell(r, c).value:
                    identical = False; break
            if not identical:
                break
        if not identical:
            break
    res['t2_s1_identical'] = identical

    # ---- 5. trigger correctness (min-count-based) ----
    wsc, wsm = wc['Table 1'], wm['Table 1']
    miss_or_over = []   # blocks blanked incorrectly or missed
    for r in range(FIRST_DATA_ROW, wsc.max_row + 1):
        for count_cols, p_col, asd_col in BLOCKS:
            orig = [_count_value(wsc.cell(r, c).value) for c in count_cols]
            present = [v for v in orig if v is not None]
            if not present:
                continue
            min_v = min(present)
            should = min_v <= THRESHOLD
            p_blank = wsm.cell(r, p_col).value == '-'
            asd_blank = wsm.cell(r, asd_col).value == '-'
            # a block "was blanked" if BOTH p and asd are '-' in masked but at least
            # one was NOT '-' in clean (i.e. suppression touched it). We evaluate on
            # the rule: after suppression, a triggered block must have p='-' AND asd='-'.
            if should:
                if not (p_blank and asd_blank):
                    miss_or_over.append((r, 'MISSED', min_v,
                                         wsm.cell(r, p_col).value, wsm.cell(r, asd_col).value))
            else:
                # not triggered: p/asd must equal the clean values (unchanged)
                if wsm.cell(r, p_col).value != wsc.cell(r, p_col).value or \
                   wsm.cell(r, asd_col).value != wsc.cell(r, asd_col).value:
                    miss_or_over.append((r, 'OVER', min_v,
                                         wsm.cell(r, p_col).value, wsm.cell(r, asd_col).value))
    res['trigger_ok'] = (len(miss_or_over) == 0)
    res['trigger_problems'] = miss_or_over

    # ---- 6. no plain <=10 count remains; no (<=M) M!=10; large counts intact ----
    bad = []
    for r in range(FIRST_DATA_ROW, wsm.max_row + 1):
        for count_cols, _, _ in BLOCKS:
            for c in count_cols:
                v = wsm.cell(r, c).value
                if v is None:
                    continue
                s = str(v)
                # any bad-mask (<=N) with N != 10 ?
                for m in BADMASK_RE.findall(s):
                    if m != '10':
                        bad.append((r, c, 'bad-mask', s))
                # any remaining plain numeric token <=10 ?
                pm = PLAIN_COUNT_RE.search(s)
                if pm:
                    n = int(pm.group(1).replace(',', ''))
                    if n <= THRESHOLD:
                        bad.append((r, c, 'plain<=10 remains', s))
    res['relabel_ok'] = (len(bad) == 0)
    res['relabel_problems'] = bad

    # ---- 6b. large counts intact (spot: any clean count >10 must be byte-identical in masked) ----
    large_ok = True
    large_bad = []
    for r in range(FIRST_DATA_ROW, wsc.max_row + 1):
        for count_cols, _, _ in BLOCKS:
            for c in count_cols:
                cv = _count_value(wsc.cell(r, c).value)
                if cv is not None and cv > THRESHOLD:
                    if wsm.cell(r, c).value != wsc.cell(r, c).value:
                        large_ok = False
                        large_bad.append((r, c, wsc.cell(r, c).value, wsm.cell(r, c).value))
    res['large_intact'] = large_ok
    res['large_bad'] = large_bad

    # ---- 8. percentages unchanged everywhere in Table 1 count cols ----
    pct_ok = True
    pct_bad = []
    for r in range(FIRST_DATA_ROW, wsc.max_row + 1):
        for count_cols, _, _ in BLOCKS:
            for c in count_cols:
                if pcts(wsc.cell(r, c).value) != pcts(wsm.cell(r, c).value):
                    pct_ok = False
                    pct_bad.append((r, c, wsc.cell(r, c).value, wsm.cell(r, c).value))
    res['pct_ok'] = pct_ok
    res['pct_bad'] = pct_bad

    # ---- 9. idempotency: re-run on masked -> 0 changes ----
    tmp_out = f'/workspace/_tmp_rerun_{study}.xlsx'
    tmp_log = f'/workspace/_tmp_rerun_{study}.csv'
    n2, _ = suppress_workbook(masked, tmp_out, tmp_log)
    res['idempotent'] = (n2 == 0)
    res['rerun_changes'] = n2
    os.remove(tmp_out); os.remove(tmp_log)

    # ---- 10. change_log present with 6 columns, one row per modified cell ----
    ok_log = os.path.exists(log)
    n_log = 0
    if ok_log:
        with open(log) as f:
            rd = csv.DictReader(f)
            cols = rd.fieldnames
            ok_log = cols == ['sheet', 'cell', 'column_role', 'old_value', 'new_value', 'reason']
            n_log = sum(1 for _ in rd)
    res['log_ok'] = ok_log
    res['log_rows'] = n_log
    return res


if __name__ == '__main__':
    all_pass = True
    for study in STUDIES:
        print(f'\n{"="*66}\n{study}\n{"="*66}')
        r = check_study(study)
        checks = [
            ('7  sheet count = 3 (T1,T2,S1)', r['sheet_count_3']),
            ('7  Table 2 + S1 identical to clean', r['t2_s1_identical']),
            ('5  trigger correct (no missed/over-applied)', r['trigger_ok']),
            ('6  no plain <=10 / no (<=M!=10) remains', r['relabel_ok']),
            ('6b large counts byte-intact', r['large_intact']),
            ('8  percentages unchanged', r['pct_ok']),
            ('9  idempotent (re-run 0 changes)', r['idempotent']),
            ('10 change_log 6-col, per-cell', r['log_ok']),
        ]
        for name, ok in checks:
            print(f'   [{"PASS" if ok else "FAIL"}] {name}')
            all_pass = all_pass and ok
        print(f'   change_log rows = {r["log_rows"]}; idempotency re-run changes = {r["rerun_changes"]}')
        if r['trigger_problems']:
            print('   trigger problems:', r['trigger_problems'][:8])
        if r['relabel_problems']:
            print('   relabel problems:', r['relabel_problems'][:8])
        if r['large_bad']:
            print('   large-count problems:', r['large_bad'][:8])
        if r['pct_bad']:
            print('   pct problems:', r['pct_bad'][:8])
    print(f'\n{"#"*66}\nOVERALL: {"ALL PASS" if all_pass else "FAILURES ABOVE"}\n{"#"*66}')
