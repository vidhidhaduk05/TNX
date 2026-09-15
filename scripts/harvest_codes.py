"""Harvest the COMPLETE universe of coded terms from each source PDF's raw text.

Ground-truth for the "no codes missed" audit. Captures every coded token of the
forms seen in TriNetX exports:
    UMLS:ICD10CM:<code>      diagnosis (ICD-10-CM)
    UMLS:ICD10PCS:<code>     procedure (ICD-10-PCS)
    UMLS:RXNORM:<code>       medication (RxNorm)
    TNX:<code>               TriNetX custom (labs, etc.)
    UMLS:CPT:<code> / HCPCS  procedures (if present)

Descriptions in the raw text wrap across lines and interleave; we capture the
code + the immediate trailing text fragment on the same physical line, then also
keep a normalized set of bare codes (system, code) for a strict set-membership
audit that is independent of any description wrapping.
"""
import re
import pdfplumber

# system tokens as they literally appear in the raw text  ->  canonical system name
SYS = {
    "ICD10CM": "ICD-10-CM",
    "ICD10PCS": "ICD-10-PCS",
    "RXNORM": "RxNorm",
    "CPT": "CPT",
    "HCPCS": "HCPCS",
    "LOINC": "LOINC",
    "SNOMED": "SNOMED",
}

# UMLS:<SYS>:<code>   (code = run of non-space chars, TriNetX codes have no spaces)
RE_UMLS = re.compile(r"\bUMLS:(" + "|".join(SYS) + r"):(\S+)")
# TNX:<code>
RE_TNX = re.compile(r"\bTNX:(\S+)")


def _raw_text(pdf_path):
    return "\n".join((pg.extract_text() or "") for pg in pdfplumber.open(pdf_path).pages)


def harvest(pdf_path):
    """Return dict with:
    codes: set of (system, code) bare tuples  <- strict audit universe
    detail: sorted list of (system, code) with a best-effort trailing text fragment
    """
    txt = _raw_text(pdf_path)
    codes = set()
    detail = {}
    for m in RE_UMLS.finditer(txt):
        sys_raw, code = m.group(1), m.group(2)
        # strip trailing punctuation that clings to a code at line end
        code = code.rstrip(".,;:)")
        system = SYS[sys_raw]
        codes.add((system, code))
        # trailing fragment on same line (until newline)
        tail = txt[m.end():m.end() + 80].split("\n", 1)[0].strip()
        detail.setdefault((system, code), tail)
    for m in RE_TNX.finditer(txt):
        code = m.group(1).rstrip(".,;:)")
        codes.add(("TNX", code))
        tail = txt[m.end():m.end() + 80].split("\n", 1)[0].strip()
        detail.setdefault(("TNX", code), tail)
    return {"codes": codes, "detail": dict(sorted(detail.items()))}


if __name__ == "__main__":
    import sys
    sys.path.insert(0, "/workspace")
    from build_tables import STUDIES

    allcodes = {}
    for study, cfg in STUDIES.items():
        h = harvest(cfg["pdf"])
        allcodes[study] = h["codes"]
        by_sys = {}
        for s, c in h["codes"]:
            by_sys.setdefault(s, set()).add(c)
        print("=" * 70)
        print(study, cfg["stem"])
        for s in sorted(by_sys):
            print(f"   {s:12s}: {len(by_sys[s]):3d} unique  ->", ", ".join(sorted(by_sys[s])[:12]),
                  ("..." if len(by_sys[s]) > 12 else ""))
    # union across all studies
    union = set().union(*allcodes.values())
    print("=" * 70)
    print("UNION across 5 PDFs:", len(union), "unique (system,code)")
    ubys = {}
    for s, c in union:
        ubys.setdefault(s, set()).add(c)
    for s in sorted(ubys):
        print(f"   {s:12s}: {len(ubys[s])}")
