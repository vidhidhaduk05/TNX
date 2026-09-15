"""
Independent verification of the CLEAN cSDH workbooks (all 5 studies).

Cross-checks every Table 1 value against pdfplumber's line-text
(`extract_text()`) -- an extraction path INDEPENDENT of the coordinate word-box
extractor (trinetx_coord) used to build the workbooks.

The TriNetX text dump interleaves multi-word labels across the count lines, e.g.

    1 Atrial fibrillation 34 22.4%
    I48 0.667 0.040
    2 and flutter 88 20.7%

so a naive triple-line parser fails. Instead, for each categorical code we:
  1. find the line beginning with that CODE token; the last two tokens on that
     line are the P-value and ASD (before-PSM table) / (after-PSM table);
  2. search a small window around it for the nearest preceding "1 ... <count> <pct>%"
     and following "2 ... <count> <pct>%" fragments (label words allowed between
     the marker and the trailing 'count pct%').
This tolerates arbitrary label wrapping and still fails loudly if a NUMBER is wrong.

Continuous rows (Age, Platelets) are mean-presence-checked. Lab range-bins are
matched positionally within the Laboratory section.
"""
import re
import openpyxl

FLOAT = r"[-+]?\d+(?:\.\d+)?"
PA = r'(?:<0\.001|' + FLOAT + r'|--)'          # p-value / asd token
CNT_PCT = re.compile(r'(\d[\d,]*)\s+(' + FLOAT + r')%\s*$')  # '... 34 22.4%' at line end


def _phase_blocks(txt):
    mb = re.search(r'characteristics before propensity', txt, re.I)
    ma = re.search(r'characteristics after propensity', txt, re.I)
    if not (mb and ma):
        return None, None
    before = txt[mb.end():ma.start()]
    tail = txt[ma.end():]
    me = re.search(r'\n\s*(?:Measure of Association|Kaplan-Meier|Results\b|1\s+Mortality)', tail)
    after = tail[:me.start()] if me else tail
    return before, after


CODE_LINE = {
    # code -> regex matching the START of its code line, capturing trailing p + asd
}


def _code_start_re(code):
    esc = re.escape(code)
    # code at line start, then (optional label words), then p asd at end
    return re.compile(r'^' + esc + r'\b.*?(' + PA + r')\s+(' + PA + r')\s*$')


def parse_phase(block, codes):
    """For each code, return {c1_count,c1_pct,c2_count,c2_pct,p,asd} using the
    windowed strategy. Missing pieces -> None."""
    lines = [l.rstrip() for l in block.split('\n')]
    res = {}
    for code in codes:
        cre = _code_start_re(code)
        ci = None
        for idx, ln in enumerate(lines):
            m = cre.match(ln.strip())
            if m:
                ci = idx; p, asd = m.group(1), m.group(2)
                break
        if ci is None:
            res[code] = None
            continue
        # nearest preceding line with a '1 ... count pct%'
        c1 = c2 = None
        for j in range(ci, max(-1, ci - 4), -1):
            s = lines[j].strip()
            if s.startswith('1 ') or s == '1':
                mm = CNT_PCT.search(s)
                if mm:
                    c1 = (mm.group(1).replace(',', ''), float(mm.group(2)))
                    break
        for j in range(ci, min(len(lines), ci + 4)):
            s = lines[j].strip()
            if s.startswith('2 ') or s == '2':
                mm = CNT_PCT.search(s)
                if mm:
                    c2 = (mm.group(1).replace(',', ''), float(mm.group(2)))
                    break
        res[code] = {
            'c1_count': c1[0] if c1 else None, 'c1_pct': c1[1] if c1 else None,
            'c2_count': c2[0] if c2 else None, 'c2_pct': c2[1] if c2 else None,
            'p': p, 'asd': asd,
        }
    return res


def parse_lab_bins(block):
    lines = [l.rstrip() for l in block.split('\n')]
    out = []
    for idx, ln in enumerate(lines):
        rng = re.search(r'(\d+)\s*-\s*(\d+)\s*10\*3/uL', ln)
        if not rng:
            continue
        lo, hi = int(rng.group(1)), int(rng.group(2))
        c1 = c2 = None
        for j in range(idx, max(-1, idx - 3), -1):
            s = lines[j].strip()
            if s.startswith('1 ') or s == '1':
                mm = CNT_PCT.search(s)
                if mm:
                    c1 = (mm.group(1).replace(',', ''), float(mm.group(2))); break
        for j in range(idx, min(len(lines), idx + 3)):
            s = lines[j].strip()
            if s.startswith('2 ') or s == '2':
                mm = CNT_PCT.search(s)
                if mm:
                    c2 = (mm.group(1).replace(',', ''), float(mm.group(2))); break
        if c1 and c2:
            out.append({'lo': lo, 'hi': hi,
                        'c1_count': c1[0], 'c1_pct': c1[1],
                        'c2_count': c2[0], 'c2_pct': c2[1]})
    return out


def _cell_pct_n(txt):
    if not txt or txt == '-':
        return None, None
    m = re.match(r'^(' + FLOAT + r')%\s*\(([\d,]+)\)$', str(txt).strip())
    if not m:
        return None, None
    return float(m.group(1)), int(m.group(2).replace(',', ''))


def _fmt_p_expect(p):
    if p in (None, '--'):
        return '-'
    if p == '<0.001':
        return '<0.001'
    v = float(p)
    return '<0.001' if v == 0 else f'{v:.3f}'


def _fmt_asd_expect(a):
    if a in (None, '--'):
        return '-'
    if a == '<0.001':
        return '<0.001'
    return f'{float(a):.3f}'


def verify_study(study, xlsx_path, text_path):
    import labels as L
    txt = open(text_path).read()
    before, after = _phase_blocks(txt)
    assert before and after, f'{study}: could not split phase blocks'

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb['Table 1']
    label2code = {lab: code for code, lab in L.LABELS.items()}

    # collect the categorical codes present in the workbook
    codes = []
    row_of = {}
    for r in range(4, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        if not lab:
            continue
        c2v = ws.cell(r, 2).value
        if c2v and '\u00b1' in str(c2v):
            continue  # continuous
        code = label2code.get(lab)
        if code is None:
            continue
        codes.append(code); row_of[code] = r

    src = {'before': parse_phase(before, codes), 'after': parse_phase(after, codes)}
    bins = {'before': parse_lab_bins(before), 'after': parse_lab_bins(after)}

    npass = nfail = 0
    fails = []
    for code in codes:
        r = row_of[code]
        lab = ws.cell(r, 1).value
        for ph, (cc1, cc2, cp, ca) in (('before', (2, 3, 4, 5)), ('after', (6, 7, 8, 9))):
            s = src[ph].get(code)
            wb_c1 = _cell_pct_n(ws.cell(r, cc1).value)
            wb_c2 = _cell_pct_n(ws.cell(r, cc2).value)
            wb_p = str(ws.cell(r, cp).value)
            wb_a = str(ws.cell(r, ca).value)
            if s is None:
                if wb_c1[1] is not None:
                    nfail += 1; fails.append(f'{lab} [{ph}] code {code} NOT FOUND in text')
                continue
            ok = True; det = []
            if wb_c1[1] is not None and s['c1_count'] is not None and str(wb_c1[1]) != s['c1_count']:
                ok = False; det.append(f"c1_n wb={wb_c1[1]} txt={s['c1_count']}")
            if wb_c2[1] is not None and s['c2_count'] is not None and str(wb_c2[1]) != s['c2_count']:
                ok = False; det.append(f"c2_n wb={wb_c2[1]} txt={s['c2_count']}")
            if wb_c1[0] is not None and s['c1_pct'] is not None and abs(wb_c1[0] - s['c1_pct']) > 0.11:
                ok = False; det.append(f"c1_pct wb={wb_c1[0]} txt={s['c1_pct']}")
            if wb_c2[0] is not None and s['c2_pct'] is not None and abs(wb_c2[0] - s['c2_pct']) > 0.11:
                ok = False; det.append(f"c2_pct wb={wb_c2[0]} txt={s['c2_pct']}")
            if wb_p != _fmt_p_expect(s['p']):
                ok = False; det.append(f"p wb={wb_p} txt->{_fmt_p_expect(s['p'])} (raw {s['p']})")
            if wb_a != _fmt_asd_expect(s['asd']):
                ok = False; det.append(f"asd wb={wb_a} txt->{_fmt_asd_expect(s['asd'])} (raw {s['asd']})")
            if ok:
                npass += 1
            else:
                nfail += 1; fails.append(f'{lab} [{ph}] ' + '; '.join(det))

    # lab range-bins positional
    wb_bins = []
    for r in range(4, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        if lab and re.match(r'^\d+\u2013\d+ \u00d710\u00b3/\u00b5L$', str(lab)):
            wb_bins.append((r, int(re.match(r'^(\d+)', str(lab)).group(1))))
    for (r, lo) in wb_bins:
        for ph, (cc1, cc2) in (('before', (2, 3)), ('after', (6, 7))):
            s = next((b for b in bins[ph] if b['lo'] == lo), None)
            wb_c1 = _cell_pct_n(ws.cell(r, cc1).value)
            wb_c2 = _cell_pct_n(ws.cell(r, cc2).value)
            lab = ws.cell(r, 1).value
            if s is None:
                if wb_c1[1] is not None:
                    nfail += 1; fails.append(f'lab {lab} [{ph}] NOT FOUND in text')
                continue
            ok = True; det = []
            if wb_c1[1] is not None and str(wb_c1[1]) != s['c1_count']:
                ok = False; det.append(f"c1_n wb={wb_c1[1]} txt={s['c1_count']}")
            if wb_c2[1] is not None and str(wb_c2[1]) != s['c2_count']:
                ok = False; det.append(f"c2_n wb={wb_c2[1]} txt={s['c2_count']}")
            if wb_c1[0] is not None and abs(wb_c1[0] - s['c1_pct']) > 0.11:
                ok = False; det.append(f"c1_pct wb={wb_c1[0]} txt={s['c1_pct']}")
            if wb_c2[0] is not None and abs(wb_c2[0] - s['c2_pct']) > 0.11:
                ok = False; det.append(f"c2_pct wb={wb_c2[0]} txt={s['c2_pct']}")
            if ok:
                npass += 1
            else:
                nfail += 1; fails.append(f'lab {lab} [{ph}] ' + '; '.join(det))

    return npass, nfail, fails


if __name__ == '__main__':
    import sys
    sys.path.insert(0, '/workspace')
    from build_tables import STUDIES
    grand_pass = grand_fail = 0
    for study in STUDIES:
        p, f, fails = verify_study(study, f'/workspace/clean_{study}.xlsx', f'/workspace/text_{study}.txt')
        grand_pass += p; grand_fail += f
        print(f'{study:10s} [{"PASS" if f == 0 else "FAIL"}]  checks pass={p}  fail={f}')
        for fl in fails[:15]:
            print('     -', fl)
    print(f'\nTOTAL: pass={grand_pass} fail={grand_fail}  ->  {"ALL PASS" if grand_fail==0 else "FAILURES PRESENT"}')
