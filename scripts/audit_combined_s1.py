"""Audit: guarantee the combined Table S1 misses NO codes from any of the 5 PDFs.

Two independent directions:

(A) FORWARD  - every atomic code appearing in the combined master table appears
    literally in the raw text of at least one source PDF. (No invented codes.)

(B) REVERSE  - every code present in each source PDF is represented in the combined
    master. We assemble each PDF's code universe from BOTH:
      (1) Appendix-A / cohort-definition tokens: UMLS:ICD10CM|ICD10PCS|RXNORM:<code>,
          TNX:<code>   (via harvest_codes.harvest)
      (2) the per-study verified Table S1 row codes (which already encode the
          covariate ICD-10 codes + RxNorm drug codes parsed from the Table-1 region).
    The union of (1) and (2), per PDF, is the ground-truth universe. Every element
    must be found in the combined master's code strings.

Atomic-code extraction from a free-text code string uses tolerant patterns that
match how codes are written in the master:
    ICD-10-CM/PCS : letters+digits with optional dots/ranges, e.g. I62.03, E08-E13,
                    D68.0, G93.4, 03LG3BZ, 00N0, 0094
    RxNorm        : 'RxNorm <digits>'
    TNX           : 'TNX:<digits>'
"""
import re
import sys

sys.path.insert(0, "/workspace")
from build_tables import STUDIES  # noqa: E402
from harvest_codes import harvest  # noqa: E402
from build_combined_s1 import build_merged, STUDY_ORDER  # noqa: E402

# ---- atomic code extractors -------------------------------------------------
# ICD-10 code or range token, e.g. I10, I62.03, E08-E13, D68.0, G93.4, C00-D49
RE_ICD = re.compile(r"\b([A-TV-Z]\d{2}(?:\.\d+)?(?:-[A-TV-Z]\d{2}(?:\.\d+)?)?)\b")
# ICD-10-PCS 7-char code (letters+digits, no dot), or the 4-char surgical roots (00N0..)
RE_PCS = re.compile(r"\b((?:03[LV]G3[A-Z]{2})|(?:00[A-Z]?\d[A-Z0-9]?)|(?:0094)|(?:00C4)|(?:00D2))\b")
RE_RXNORM = re.compile(r"RxNorm\s+(\d+)")
RE_TNX = re.compile(r"TNX:(\d+)")


def codes_in_string(s):
    """Extract the set of atomic codes present in a master 'code' cell string."""
    found = set()
    for m in RE_RXNORM.finditer(s):
        found.add(("RxNorm", m.group(1)))
    for m in RE_TNX.finditer(s):
        found.add(("TNX", m.group(1)))
    # PCS first (7-char), then ICD; strip PCS spans so ICD regex doesn't re-hit them
    pcs_spans = []
    for m in RE_PCS.finditer(s):
        found.add(("ICD-10-PCS", m.group(1)))
        pcs_spans.append((m.start(), m.end()))
    # mask PCS regions before ICD scan
    masked = list(s)
    for a, b in pcs_spans:
        for i in range(a, b):
            masked[i] = " "
    masked = "".join(masked)
    for m in RE_ICD.finditer(masked):
        found.add(("ICD-10-CM", m.group(1)))
    return found


def master_code_universe(rows):
    uni = set()
    per_row = []
    for row in rows:
        cs = codes_in_string(row["code"])
        per_row.append((row, cs))
        uni |= cs
    return uni, per_row


def normalize_for_text(system, code):
    """How this code should appear literally in raw PDF text (for FORWARD check)."""
    # ICD codes in raw text sometimes appear as the bare code (covariates) or inside
    # UMLS:ICD10CM:. Ranges like E08-E13 appear literally. PCS appear bare or in UMLS.
    return code


def raw_pdf_text(pdf):
    import pdfplumber
    return "\n".join((pg.extract_text() or "") for pg in pdfplumber.open(pdf).pages)


def main():
    rows = build_merged()
    uni, per_row = master_code_universe(rows)
    by_sys = {}
    for s, c in uni:
        by_sys.setdefault(s, set()).add(c)
    print("MASTER atomic-code universe:", sum(len(v) for v in by_sys.values()), "unique")
    for s in sorted(by_sys):
        print(f"   {s:12s}: {len(by_sys[s]):3d}  ->", ", ".join(sorted(by_sys[s])))

    texts = {st: raw_pdf_text(cfg["pdf"]) for st, cfg in STUDIES.items()}

    # ---------- (A) FORWARD: every master code literally in >=1 PDF ----------
    print("\n=== (A) FORWARD: master code -> present literally in some PDF ===")
    fwd_fail = []
    for system, code in sorted(uni):
        lit = normalize_for_text(system, code)
        if not any(lit in t for t in texts.values()):
            fwd_fail.append((system, code))
    if fwd_fail:
        print("  FAIL - master codes NOT found in any PDF text:")
        for s, c in fwd_fail:
            print("     ", s, c)
    else:
        print(f"  PASS - all {len(uni)} master codes appear literally in >=1 source PDF")

    # ---------- (B) REVERSE: every PDF code represented in master ----------
    print("\n=== (B) REVERSE: every source-PDF code -> present in master ===")
    # master code strings concatenated for substring membership
    master_blob = " || ".join(r["code"] for r in rows)
    rev_fail = {}
    for study, cfg in STUDIES.items():
        # universe (1): appendix/cohort tokens
        h = harvest(cfg["pdf"])
        uni1 = h["codes"]  # set of (system, code)
        # universe (2): atomic codes from this study's own verified S1 rows
        from build_combined_s1 import read_s1
        uni2 = set()
        for _sec, _label, codeval in read_s1(study):
            uni2 |= codes_in_string(codeval)
        pdf_uni = uni1 | uni2
        missing = []
        for system, code in sorted(pdf_uni):
            # is this code present anywhere in the master code strings?
            if code not in master_blob:
                missing.append((system, code))
        if missing:
            rev_fail[study] = missing
        print(f"  {study:9s}: PDF universe={len(pdf_uni):3d} (appendix={len(uni1)}, S1={len(uni2)})  "
              f"-> {'MISSING ' + str(len(missing)) if missing else 'all present'}")
        for s, c in missing:
            print("        MISSING:", s, c)

    print("\n" + "=" * 60)
    ok = (not fwd_fail) and (not rev_fail)
    print("AUDIT RESULT:", "PASS - no codes missed in either direction" if ok
          else "FAIL - see above")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
