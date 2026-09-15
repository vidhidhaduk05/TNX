"""
Coordinate-aware extractor for TriNetX "Compare Outcomes" characteristics tables.

Strategy
--------
The characteristics tables have a fixed column layout with a header row:
    Cohort | <code> <label> | Mean ± SD | Median | IQR | Patients | % of Cohort | P-Value | Std diff.
Each data row spans TWO physical text lines (cohort 1 on top, cohort 2 below),
with the cohort marker column showing '1' then '2'. Categorical rows leave the
Mean/Median/IQR columns blank and populate Patients + %; continuous ("mean") rows
populate Mean±SD/Median/IQR and leave Patients as the N.

Rather than parse the fragile token stream, we:
  1. Locate each "Cohort  Mean  SD  Median IQR Patients % of Cohort P-Value Std diff."
     header line and read the x-centers of Median / IQR / Patients / %/ P-Value / Std diff.
  2. Define column x-windows from those anchors (Mean±SD is everything left of Median
     and right of the label block).
  3. For every subsequent data row (until the next section header / table), group words
     into cohort-1 vs cohort-2 by y (two sub-rows), and bucket each numeric token into a
     column by its x-center.

We key rows by the ICD/RxNorm/lab code when present, else by the reconstructed label.
Mean±SD tokens are read positionally from the Mean±SD column (top value = mean, the
'+/-' and the SD are stacked); this eliminates the token-wrap ambiguity.

Only the fields needed downstream are emitted:
  - categorical: c1_count, c1_pct, c2_count, c2_pct, p, sd, code, label
  - continuous : c1_mean, c1_sd, c2_mean, c2_sd, p, sd, code, label
"""
import re
import pdfplumber

CODE_RE = re.compile(
    r"^(AI|A1|F|M|Z87\.891|[A-Z]\d{2}(\.[0-9A-Za-z]+)?|[A-Z]\d{2}-[A-Z]\d{2}|\d{4,7})$"
)
FLOAT = re.compile(r"^\d+(\.\d+)?$")
PCT = re.compile(r"^\d+(\.\d+)?%$")

SECTION_HEADERS = ("Demographics", "Diagnosis", "Medication", "Laboratory")
# Header line tokens we anchor on
HEADER_ANCHORS = ["Median", "IQR", "Patients", "P-Value", "Std"]


def _lines_by_top(words, tol=2.0):
    """Group words into physical lines by their 'top' coordinate."""
    lines = []
    for w in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
        placed = False
        for ln in lines:
            if abs(ln["top"] - w["top"]) <= tol:
                ln["words"].append(w)
                ln["top"] = (ln["top"] * (len(ln["words"]) - 1) + w["top"]) / len(ln["words"])
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})
    for ln in lines:
        ln["words"].sort(key=lambda w: w["x0"])
        ln["text"] = " ".join(w["text"] for w in ln["words"])
    lines.sort(key=lambda ln: ln["top"])
    return lines


def _xcenter(w):
    return (w["x0"] + w["x1"]) / 2.0


def extract_pdf_sections(pdf_path):
    """Return list of table dicts, each: {title, page, rows:[...]}.

    A new table starts at a line matching 'Cohort N ( ... ) characteristics
    (before|after) propensity score matching'. Rows accumulate under the
    current section header until the next table title.
    """
    pdf = pdfplumber.open(pdf_path)
    tables = []
    cur = None
    cur_section = None
    col_anchors = None  # dict of column -> x window

    title_re = re.compile(r"characteristics (before|after) propensity score matching")
    # A characteristics table ends when the Results / outcomes section begins.
    end_re = re.compile(r"^(Results|Follow-up Time|Risk analysis|Kaplan\s*-\s*Meier)\b")

    for pno, page in enumerate(pdf.pages, start=1):
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        if not words:
            continue
        lines = _lines_by_top(words)
        i = 0
        while i < len(lines):
            ln = lines[i]
            txt = ln["text"]

            m = title_re.search(txt)
            if m:
                # start a new table
                cur = {"title": txt, "phase": m.group(1), "page": pno, "rows": []}
                tables.append(cur)
                cur_section = None
                col_anchors = None
                i += 1
                continue

            # end of a characteristics table: Results / outcomes section begins
            if cur is not None and end_re.search(txt):
                cur = None
                cur_section = None
                col_anchors = None
                i += 1
                continue

            # section header (single token that is a known section, alone-ish)
            first_tok = ln["words"][0]["text"] if ln["words"] else ""
            if first_tok in SECTION_HEADERS and len(ln["words"]) <= 2 and cur is not None:
                cur_section = first_tok
                i += 1
                continue

            # column header row: "Cohort Mean SD Median IQR Patients % of Cohort P-Value Std diff."
            if cur is not None and "Median" in txt and "IQR" in txt and "P-Value" in txt:
                # build column anchors from x-centers
                anchor = {}
                for w in ln["words"]:
                    t = w["text"]
                    if t == "Median":
                        anchor["median_x"] = _xcenter(w)
                    elif t == "IQR":
                        anchor["iqr_x"] = _xcenter(w)
                    elif t == "Patients":
                        anchor["patients_x"] = _xcenter(w)
                    elif t == "Cohort":
                        # header has TWO 'Cohort' tokens: the leftmost cohort-marker
                        # column (~x127) and one inside '% of Cohort' (~x431). Always
                        # keep the LEFTMOST; the right one would wrongly widen the
                        # cohort-marker guard and swallow a P-value '1'.
                        xc = _xcenter(w)
                        if "cohort_x" not in anchor or xc < anchor["cohort_x"]:
                            anchor["cohort_x"] = xc
                    elif t == "%":
                        anchor["pct_x"] = _xcenter(w)
                    elif t == "P-Value":
                        anchor["p_x"] = _xcenter(w)
                    elif t.startswith("Std"):
                        anchor["sd_x"] = _xcenter(w)
                col_anchors = anchor
                i += 1
                continue

            # data row detection: a line whose first token is cohort marker '1'
            if cur is not None and col_anchors is not None and first_tok == "1":
                # gather this line (cohort1) + the following cohort2 line + any wrap lines
                # We collect all words until we hit the next line starting with '1' or a
                # section/header/title, then split by y into two cohort sub-rows.
                block_lines = [ln]
                # For continuous (mean) rows, the cohort-1 mean token ("64.9 +/-") and part
                # of the label ("Platelets") wrap to line(s) ABOVE the '1' marker. Walk
                # backward and pull in any preceding lines that belong to this row: mean
                # fragments (contain '+/-') or non-structural label fragments. Stop at the
                # previous row's markers, a section/title/header, or a completed data line.
                k = i - 1
                pulled = []
                while k >= 0 and (i - k) <= 4:
                    prev = lines[k]
                    ptxt = prev["text"].strip()
                    pfirst = prev["words"][0]["text"] if prev["words"] else ""
                    if pfirst in ("1", "2"):
                        break
                    if (title_re.search(ptxt) or
                            (pfirst in SECTION_HEADERS and len(prev["words"]) <= 2) or
                            ("Median" in ptxt and "IQR" in ptxt and "P-Value" in ptxt)):
                        break
                    is_meanfrag = "+/-" in ptxt
                    # label fragment: begins with an alphabetic label token (e.g. 'Platelets',
                    # 'Blood', 'Other'); may carry a stray median/IQR number merged by line
                    # grouping (e.g. 'Platelets 20.5'). We still pull it in because the '1'/'2'
                    # marker check above is what protects against crossing into the prior row.
                    starts_alpha = bool(re.match(r"^[A-Za-z\[]", pfirst))
                    is_labelfrag = starts_alpha and len(prev["words"]) <= 4
                    if is_meanfrag or is_labelfrag:
                        pulled.append(prev)
                        k -= 1
                        continue
                    break
                for p in pulled:  # pulled is nearest-first; insert to keep document order
                    block_lines.insert(0, p)
                j = i + 1
                while j < len(lines):
                    nxt = lines[j]
                    nfirst = nxt["words"][0]["text"] if nxt["words"] else ""
                    ntxt = nxt["text"]
                    if nfirst == "1" or title_re.search(ntxt) or end_re.search(ntxt) or \
                       (nfirst in SECTION_HEADERS and len(nxt["words"]) <= 2) or \
                       ("Median" in ntxt and "IQR" in ntxt and "P-Value" in ntxt):
                        break
                    block_lines.append(nxt)
                    j += 1
                row = _decode_row(block_lines, col_anchors, cur_section)
                if row:
                    cur["rows"].append(row)
                i = j
                continue

            i += 1
    return tables


def _decode_row(block_lines, cols, section):
    """Decode one data row (spanning 2 cohorts) using column x-windows."""
    all_words = []
    for ln in block_lines:
        all_words.extend(ln["words"])
    # cohort marker positions: markers '1' and '2' with small x (< cohort label area).
    # Cap well left of the numeric columns (Median starts ~x297) so a balanced-covariate
    # P-value '1' (x~466) can NEVER be mistaken for a row marker.
    marker_x_max = min(cols.get("cohort_x", 120) + 40, 200)
    markers = [w for w in all_words if w["text"] in ("1", "2") and w["x0"] < marker_x_max]
    # y split: use the '2' marker top as boundary; fallback to midpoint
    tops = sorted(w["top"] for w in all_words)
    two_marker = [w for w in markers if w["text"] == "2"]
    if two_marker:
        ysplit = min(w["top"] for w in two_marker) - 1.0
    else:
        ysplit = (tops[0] + tops[-1]) / 2.0

    # code + label = tokens left of Mean column (x < mean window right edge) that are not markers
    # Mean±SD column right edge ~ median_x - a bit
    median_x = cols["median_x"]
    iqr_x = cols["iqr_x"]
    patients_x = cols["patients_x"]
    pct_x = cols.get("pct_x", patients_x + 40)
    p_x = cols.get("p_x", pct_x + 60)
    sd_x = cols.get("sd_x", p_x + 45)

    # Column geometry (observed): label/code sits at small x (~130-215); the Mean±SD
    # numeric column is centered ~35-55 px left of the Median column; Median ~ median_x.
    # So define:
    #   label region : xc <  mean_lo
    #   Mean±SD col   : mean_lo <= xc < median_x - 12
    mean_lo = median_x - 60          # left edge of Mean±SD numeric column
    label_right = mean_lo            # label/code strictly left of the Mean column
    mean_hi = median_x - 12

    code = None
    label_words = []
    for w in sorted(all_words, key=lambda w: (w["top"], w["x0"])):
        t = w["text"]
        if t in ("1", "2") and w["x0"] < marker_x_max:
            continue
        xc = _xcenter(w)
        if xc < label_right:
            if code is None and CODE_RE.match(t):
                code = t
            else:
                label_words.append(t)
    label = " ".join(label_words)
    label = re.sub(r"\s+", " ", label).strip()

    def col_tokens(xlo, xhi, cohort):
        out = []
        for w in all_words:
            if w["text"] in ("1", "2") and w["x0"] < marker_x_max:
                continue
            xc = _xcenter(w)
            if xlo <= xc < xhi:
                is_c1 = w["top"] < ysplit
                if (cohort == 1 and is_c1) or (cohort == 2 and not is_c1):
                    out.append((w["top"], w["text"]))
        out.sort()
        return [t for _, t in out]

    is_mean = any("+/-" in w["text"] or w["text"] == "+/-" for w in all_words) or \
              any(w["text"].endswith("+/-") for w in all_words)

    # mean_lo, mean_hi already defined above from median_x geometry
    pat_lo, pat_hi = patients_x - 22, patients_x + 22
    pct_lo, pct_hi = pct_x - 30, pct_x + 30
    p_lo, p_hi = p_x - 28, p_x + 28
    sd_lo, sd_hi = sd_x - 28, sd_x + 40

    row = {"section": section, "code": code, "label": label, "is_mean": is_mean}

    # p-value and std-diff (shared, appear once per row on cohort-1 or spanning both)
    p_toks = col_tokens(p_lo, p_hi, 1) + col_tokens(p_lo, p_hi, 2)
    sd_toks = col_tokens(sd_lo, sd_hi, 1) + col_tokens(sd_lo, sd_hi, 2)
    p_toks = [t for t in p_toks if FLOAT.match(t) or t in ("<0.001", "--", "1")]
    sd_toks = [t for t in sd_toks if FLOAT.match(t) or t in ("<0.001", "--", "1")]
    row["p"] = p_toks[0] if p_toks else None
    row["sd_stat"] = sd_toks[0] if sd_toks else None

    if is_mean:
        # Mean rows stack four numeric tokens in the Mean±SD column, top-to-bottom:
        #   [C1 mean, C1 SD, C2 mean, C2 SD].  The '+/-' glyphs sit next to the means.
        # Collect ALL numeric tokens in the Mean window regardless of ysplit, ordered by top.
        mean_col = []
        for w in all_words:
            if w["text"] in ("1", "2") and w["x0"] < marker_x_max:
                continue
            xc = _xcenter(w)
            if mean_lo <= xc < mean_hi and FLOAT.match(w["text"]):
                mean_col.append((w["top"], w["text"]))
        mean_col.sort()
        vals = [t for _, t in mean_col]
        row["c1_mean"] = vals[0] if len(vals) >= 1 else None
        row["c1_sd"] = vals[1] if len(vals) >= 2 else None
        row["c2_mean"] = vals[2] if len(vals) >= 3 else None
        row["c2_sd"] = vals[3] if len(vals) >= 4 else None
        row["_mean_col_raw"] = vals
        # patients (N) for context
        c1_pat = [t for t in col_tokens(pat_lo, pat_hi, 1) if re.match(r"^[\d,]+$", t)]
        c2_pat = [t for t in col_tokens(pat_lo, pat_hi, 2) if re.match(r"^[\d,]+$", t)]
        row["c1_n"] = c1_pat[0] if c1_pat else None
        row["c2_n"] = c2_pat[0] if c2_pat else None
    else:
        c1_pat = [t for t in col_tokens(pat_lo, pat_hi, 1) if re.match(r"^[\d,]+$", t)]
        c2_pat = [t for t in col_tokens(pat_lo, pat_hi, 2) if re.match(r"^[\d,]+$", t)]
        c1_pct = [t for t in col_tokens(pct_lo, pct_hi, 1) if PCT.match(t) or t == "0%"]
        c2_pct = [t for t in col_tokens(pct_lo, pct_hi, 2) if PCT.match(t) or t == "0%"]
        row["c1_count"] = c1_pat[0] if c1_pat else None
        row["c2_count"] = c2_pat[0] if c2_pat else None
        row["c1_pct"] = c1_pct[0] if c1_pct else None
        row["c2_pct"] = c2_pct[0] if c2_pct else None
    return row


if __name__ == "__main__":
    import sys
    tables = extract_pdf_sections(sys.argv[1])
    for t in tables:
        print(f"\n===== TABLE: {t['phase'].upper()} (page {t['page']}) : {len(t['rows'])} rows =====")
        for r in t["rows"]:
            if r["is_mean"]:
                print(f"  [{(r['section'] or '?')[:4]}] {r['code']!s:9} MEAN c1={r['c1_mean']}\u00b1{r['c1_sd']} c2={r['c2_mean']}\u00b1{r['c2_sd']} p={r['p']} sd={r['sd_stat']} N=({r['c1_n']},{r['c2_n']}) | {r['label']}")
            else:
                print(f"  [{(r['section'] or '?')[:4]}] {r['code']!s:9} c1={r['c1_count']}({r['c1_pct']}) c2={r['c2_count']}({r['c2_pct']}) p={r['p']} sd={r['sd_stat']} | {r['label']}")
