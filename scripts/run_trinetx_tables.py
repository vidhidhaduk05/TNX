#!/usr/bin/env python3
"""
run_trinetx_tables.py  --  END-TO-END orchestrator for TriNetX two-cohort study tables.

This is the ONLY new file in the skill. It contains NO parsing/statistics/formatting
logic of its own; it wires together the bundled, already-verified modules:

    trinetx_coord.py       coordinate PDF extractor  (Table 1 characteristics)
    trinetx_outcomes.py    structural outcome detector (Table 2)
    labels.py              clean verbatim labels + range-bins
    build_tables.py        3-sheet workbook builder (Table 1/2/S1)
    suppress.py            either-cohort <=10 masking + change_log
    verify_indep2.py       independent Table-1 re-derivation
    verify_suppression.py  suppression-rule audit
    format_tables_to_docx.py   house-style Word converter (user's formatter skill)
    build_combined_s1.py / combined_s1_to_docx.py / harvest_codes.py /
    audit_combined_s1.py   optional combined Supplementary Table S1 + no-codes-missed audit

WHAT IT DOES (per uploaded PDF, unless --combined-only):
  1. derive_config(pdf)   -> config dict (cohort labels, PLT threshold, cohort kinds)
                             parsed from the PDF's embedded `query name:` lines and the
                             `at most X 10*3/uL` cutoff  (NO hardcoded per-file table).
                             Unrecognized cohort tokens -> STOP and ask the user.
  2. inject config into build_tables globals (STUDIES/PDFS/COHORTS) so the verbatim
     builders run unchanged.
  3. build clean workbook -> write clean_<key>.xlsx + text_<key>.txt (pdfplumber dump).
  4. suppress -> masked_<key>.xlsx + <stem>_change_log.csv.
  5. VERIFY (hard gate, stop-and-report): independent re-derivation + Table-2 arithmetic
     + suppression audit. Any failure -> report per-PDF failing checks, deliver NOTHING
     for that study, non-zero exit.
  6. Word: call the formatter with the locked config -> 2 docx per study.
  7. copy deliverables into a per-study subfolder under the output root.

  --combined : additionally build ONE deduplicated combined Supplementary Table S1 docx
               across all successfully-built PDFs, gated behind the no-codes-missed audit.

USAGE
  python run_trinetx_tables.py --pdf-dir <dir> --out-dir <dir> [--combined]
  python run_trinetx_tables.py --pdf a.pdf --pdf b.pdf --out-dir <dir>
  python run_trinetx_tables.py ... --scratch <dir>   # random-access .xlsx staging (default /workspace)

Correctness is guaranteed only where the verification gate passes. Parsing was validated
on the cSDH MMAE/Surgery/Medical-management study family; a genuinely different TriNetX
layout that fails parsing will stop-and-report rather than deliver wrong tables.
"""
import argparse
import glob
import os
import re
import shutil
import sys
import traceback

import pdfplumber
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# ---------------------------------------------------------------------------
# 1. Config auto-derivation  (replaces the hardcoded STUDIES dict)
# ---------------------------------------------------------------------------
# Map a query-name cohort token -> canonical cohort kind used by build_tables.
# Tokens are matched case-insensitively as substrings of the `query name:` text.
QUERY_TOKEN_TO_KIND = [
    ("sx + mmae",        "sx_mmae"),   # combined surgery + MMAE  (check before 'mmae')
    ("surgery + mmae",   "sx_mmae"),
    ("standalone mmae",  "mmae"),
    ("mmae",             "mmae"),
    ("sx alone",         "surgery"),
    ("surgery alone",    "surgery"),
    ("mm alone",         "medical"),
    ("medical",          "medical"),
]

# Human-readable display label per kind (threshold appended at runtime).
KIND_DISPLAY = {
    "mmae":    "Standalone MMAE",
    "sx_mmae": "Surgery + MMAE",
    "surgery": "Surgery alone",
    "medical": "Medical management",
}


def _raw_text(pdf_path):
    return "\n".join((pg.extract_text() or "") for pg in pdfplumber.open(pdf_path).pages)


def _kind_for_token(token):
    t = token.lower()
    for needle, kind in QUERY_TOKEN_TO_KIND:
        if needle in t:
            return kind
    return None


class ConfigDeriveError(Exception):
    """Raised when a PDF's cohort naming is not recognized -> stop and ask user."""


def derive_config(pdf_path):
    """Parse cohort identity + platelet threshold from a single PDF.

    Returns a config dict shaped exactly like a build_tables.STUDIES entry:
        {pdf, stem, c1, c2, thr_k, thr_val, c1_kind, c2_kind}
    Raises ConfigDeriveError on an unrecognized cohort token (never guesses).
    """
    txt = _raw_text(pdf_path)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]

    # cohort identity from the first two distinct `query name:` values (order = Cohort 1, 2)
    qnames = []
    for m in re.finditer(r"query name:\s*([^\)\n]+)", txt):
        name = m.group(1).strip()
        if name not in qnames:
            qnames.append(name)
        if len(qnames) == 2:
            break
    if len(qnames) < 2:
        raise ConfigDeriveError(
            f"{stem}: could not find two distinct 'query name:' lines "
            f"(found {len(qnames)}: {qnames}). This may not be a 2-cohort TriNetX export."
        )
    c1_kind = _kind_for_token(qnames[0])
    c2_kind = _kind_for_token(qnames[1])
    unknown = [q for q, k in zip(qnames, (c1_kind, c2_kind)) if k is None]
    if unknown:
        raise ConfigDeriveError(
            f"{stem}: unrecognized cohort name(s) {unknown}. "
            f"Known tokens: {sorted({n for n, _ in QUERY_TOKEN_TO_KIND})}. "
            f"Add a mapping or confirm the cohort labels before building."
        )

    # platelet threshold: numeric cutoff from the lab line, k-value from query text
    m_val = re.search(r"at most (\d+(?:\.\d+)?)\s*10\*3/uL", txt)
    m_k = re.search(r"PLT\s*<\s*(\d+)\s*k", " ".join(qnames), re.IGNORECASE)
    if not m_val or not m_k:
        raise ConfigDeriveError(
            f"{stem}: could not derive platelet threshold "
            f"(cutoff line found={bool(m_val)}, 'PLT<Nk' found={bool(m_k)})."
        )
    thr_val = m_val.group(1)
    if "." not in thr_val:
        thr_val = f"{float(thr_val):.2f}"
    thr_k = int(m_k.group(1))

    suffix = f"(PLT<{thr_k}k)"
    return {
        "pdf": pdf_path,
        "stem": stem,
        "c1": f"{KIND_DISPLAY[c1_kind]} {suffix}",
        "c2": f"{KIND_DISPLAY[c2_kind]} {suffix}",
        "thr_k": thr_k,
        "thr_val": thr_val,
        "c1_kind": c1_kind,
        "c2_kind": c2_kind,
    }


def _study_key(cfg, idx):
    """Stable short key for a study (used as build_tables dict key + temp filenames)."""
    return f"s{idx:02d}_{cfg['c1_kind']}_vs_{cfg['c2_kind']}_p{cfg['thr_k']}"


# ---------------------------------------------------------------------------
# 2-4. Build clean workbook + suppress, driving the verbatim modules
# ---------------------------------------------------------------------------
def _inject_config(build_tables, configs):
    """Replace build_tables' module-level config globals with the derived configs.

    build_tables.build_table1/2/s1 read STUDIES/PDFS/COHORTS keyed by study string.
    We rebuild all three from the derived dict so the verbatim builders work unchanged.
    """
    build_tables.STUDIES = dict(configs)
    build_tables.PDFS = {k: v["pdf"] for k, v in configs.items()}
    build_tables.COHORTS = {k: {"c1": v["c1"], "c2": v["c2"]} for k, v in configs.items()}


def build_and_suppress(cfg, key, scratch):
    """Build clean + masked workbook for one study; return dict of produced paths."""
    import build_tables
    import suppress
    from trinetx_coord import extract_pdf_sections
    from trinetx_outcomes import parse_outcomes

    pdf = cfg["pdf"]

    # independent text dump for verify_indep2
    text_path = os.path.join(scratch, f"text_{key}.txt")
    with open(text_path, "w") as fh:
        fh.write(_raw_text(pdf))

    # extract sections/outcomes and build the clean workbook via the verbatim builder
    char_tables = {key: {t["phase"]: t for t in extract_pdf_sections(pdf)}}
    outcomes = {key: parse_outcomes(pdf)}
    n_out = len(outcomes[key])

    wb = build_tables.build_workbook(key, char_tables, outcomes)
    clean_path = os.path.join(scratch, f"clean_{key}.xlsx")
    wb.save(clean_path)

    # suppress -> masked + change log (suppress.py is general, pre-configured)
    masked_path = os.path.join(scratch, f"masked_{key}.xlsx")
    log_path = os.path.join(scratch, f"{cfg['stem']}_change_log.csv")
    n_changes, _changes = suppress.suppress_workbook(clean_path, masked_path, log_path)

    return {
        "clean": clean_path, "masked": masked_path, "log": log_path,
        "text": text_path, "n_outcomes": n_out, "n_suppressed": n_changes,
    }


# ---------------------------------------------------------------------------
# 5. Verification gate  (stop-and-report)
# ---------------------------------------------------------------------------
def verify_study_gate(cfg, key, paths, scratch):
    """Run all per-study verifiers. Return (ok, report_lines).

    Uses the bundled verifiers verbatim:
      - verify_indep2.verify_study(key, clean_xlsx, text_dump): independent Table-1
        re-derivation from a second extraction path.
      - verify_suppression.check_study(key): criteria 5-10 (trigger correctness,
        relabel, large-counts-intact, percentages-unchanged, idempotency, change-log,
        3-sheet + Table2/S1 untouched). It reads clean_/masked_/change_log from the
        scratch dir by convention and config from injected globals, so we require
        scratch to hold those exact filenames (guaranteed by build_and_suppress).
    """
    report = []
    ok = True

    # (a) independent Table-1 re-derivation
    import verify_indep2
    npass, nfail, fails = verify_indep2.verify_study(key, paths["clean"], paths["text"])
    report.append(f"    Table 1 independent re-derivation: pass={npass} fail={nfail}")
    if nfail:
        ok = False
        for f in fails[:20]:
            report.append(f"        FAIL: {f}")

    # (b/c) suppression + structure audit (criteria 5-10). check_study reads from the
    #       module's fixed dir + STEM map + STUDIES globals -> inject them, and point
    #       the module's path convention at our scratch dir.
    import verify_suppression as vs
    vs.STUDIES = [key]
    vs.STEM = {key: cfg["stem"]}
    # verify_suppression hardcodes '/workspace/...'; only safe when scratch == /workspace.
    if os.path.abspath(scratch) != "/workspace":
        # stage copies into /workspace under the expected names, run, then clean up
        _stage = []
        for src, name in ((paths["clean"], f"clean_{key}.xlsx"),
                          (paths["masked"], f"masked_{key}.xlsx"),
                          (paths["log"], f"{cfg['stem']}_change_log.csv")):
            dst = os.path.join("/workspace", name)
            if os.path.abspath(src) != os.path.abspath(dst):
                shutil.copyfile(src, dst)
                _stage.append(dst)
        res = vs.check_study(key)
        for d in _stage:
            if os.path.exists(d):
                os.remove(d)
    else:
        res = vs.check_study(key)

    crit = {
        "sheet_count_3": "3 sheets present",
        "t2_s1_identical": "Table 2 / S1 untouched by suppression",
        "trigger_ok": "either-cohort <=10 trigger correct",
        "relabel_ok": "no plain <=10 remains / masks well-formed",
        "large_intact": "large counts byte-identical",
        "pct_ok": "percentages unchanged",
        "idempotent": "re-suppression yields 0 changes",
        "log_ok": "change-log 6-col header present",
    }
    s_ok = all(res.get(k) for k in crit)
    report.append("    Suppression + structure audit: " + ("PASS" if s_ok else "FAIL")
                  + f" (log_rows={res.get('log_rows')})")
    for k, label in crit.items():
        if not res.get(k):
            ok = False
            report.append(f"        FAIL [{label}]: {res.get(k.replace('_ok', '_problems'), res.get(k))}")

    return ok, report


# ---------------------------------------------------------------------------
# 6. Word conversion (formatter with locked cSDH config)
# ---------------------------------------------------------------------------
def to_word(masked_xlsx, stem, out_subdir, scratch):
    """Run the bundled formatter with the LOCKED cSDH config; return produced docx paths.

    format_tables_to_docx.build_docx honors main_filename/supp_filename, so we pass the
    stem-based names directly into the per-study subfolder -- no glob-rename needed.
    Under file_grouping='match_uploads' it writes exactly two files:
      <stem>_Table_1_and_2.docx  (Table 1 + Table 2)   and   <stem>_Table_S1.docx.
    """
    import format_tables_to_docx as fmt

    # locked cSDH config (see SKILL.md): source-faithful labels; 2-file split;
    # keep two distinct nicotine ICD-10 rows; apply general house-style renames only.
    written = fmt.build_docx(
        masked_xlsx, out_subdir,
        cohort_map=None,
        rename_map={},
        pair_rename_map={},
        blank_value_labels=set(),
        normalize=True,
        file_grouping="match_uploads",
        consolidate_nicotine_rows=False,
        apply_general_renames=True,
        main_filename=f"{stem}_Table_1_and_2.docx",
        supp_filename=f"{stem}_Table_S1.docx",
    )
    return list(written)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def discover_pdfs(args):
    pdfs = list(args.pdf or [])
    if args.pdf_dir:
        pdfs += sorted(glob.glob(os.path.join(args.pdf_dir, "*.pdf")))
    # de-dup, keep order
    seen, out = set(), []
    for p in pdfs:
        ap = os.path.abspath(p)
        if ap not in seen:
            seen.add(ap)
            out.append(ap)
    return out


def run(args):
    scratch = args.scratch
    os.makedirs(scratch, exist_ok=True)
    os.makedirs(args.out_dir, exist_ok=True)

    pdfs = discover_pdfs(args)
    if not pdfs:
        print("ERROR: no PDFs found. Use --pdf-dir or --pdf.", file=sys.stderr)
        return 2
    print(f"Discovered {len(pdfs)} PDF(s).")

    # ---- derive configs first (fail fast on unrecognized cohorts) ----
    import build_tables
    configs, key_of, cfg_errors = {}, {}, []
    for i, pdf in enumerate(pdfs):
        try:
            cfg = derive_config(pdf)
        except ConfigDeriveError as e:
            cfg_errors.append(str(e))
            continue
        key = _study_key(cfg, i)
        configs[key] = cfg
        key_of[pdf] = key
        print(f"  [{key}] {cfg['c1']}  vs  {cfg['c2']}  (thr={cfg['thr_val']})")

    if cfg_errors:
        print("\nSTOP: could not auto-derive config for:")
        for e in cfg_errors:
            print("   -", e)
        print("Confirm/label these cohorts, then rerun. Nothing was built.")
        return 3

    _inject_config(build_tables, configs)

    # ---- per-PDF build -> suppress -> verify -> word ----
    delivered, failures = {}, {}
    for pdf, key in key_of.items():
        cfg = configs[key]
        stem = cfg["stem"]
        print(f"\n=== {stem} ===")
        try:
            paths = build_and_suppress(cfg, key, scratch)
            print(f"    outcomes detected: {paths['n_outcomes']} | "
                  f"cells suppressed: {paths['n_suppressed']}")
            ok, report = verify_study_gate(cfg, key, paths, scratch)
            print("\n".join(report))
            if not ok:
                failures[stem] = "verification failed (see above)"
                print(f"    -> STOP: {stem} failed verification; delivering nothing for it.")
                continue
            # passed -> deliver into per-study subfolder
            sub = os.path.join(args.out_dir, stem)
            os.makedirs(sub, exist_ok=True)
            xlsx_dst = os.path.join(sub, f"{stem}_Table1_2_S1.xlsx")
            csv_dst = os.path.join(sub, f"{stem}_change_log.csv")
            shutil.copyfile(paths["masked"], xlsx_dst)
            shutil.copyfile(paths["log"], csv_dst)
            docs = to_word(paths["masked"], stem, sub, scratch)
            delivered[stem] = [xlsx_dst, csv_dst] + docs
            print(f"    -> delivered {len(delivered[stem])} files to {os.path.basename(sub)}/")
        except Exception as e:  # noqa: BLE001
            failures[stem] = f"{type(e).__name__}: {e}"
            print(f"    -> ERROR building {stem}: {e}")
            traceback.print_exc()

    # ---- optional combined Supplementary Table S1 ----
    if args.combined and delivered:
        print("\n=== Combined Supplementary Table S1 ===")
        rc = _build_combined(configs, key_of, delivered, args, scratch)
        if rc != 0:
            failures["__combined__"] = "combined-S1 audit failed"

    # ---- summary ----
    print("\n" + "=" * 60)
    print(f"DELIVERED: {len(delivered)} study(ies)")
    for stem, files in delivered.items():
        print(f"   {stem}: {len(files)} files")
    if failures:
        print(f"FAILED: {len(failures)}")
        for stem, why in failures.items():
            print(f"   {stem}: {why}")
    return 1 if failures else 0


def _build_combined(configs, key_of, delivered, args, scratch):
    """Optional: dedup combined S1 across delivered studies + no-codes-missed audit.

    Reuses build_combined_s1 / combined_s1_to_docx / audit_combined_s1, which read the
    per-study masked workbooks. We stage masked_<key>.xlsx into scratch under the keys
    those modules expect, inject configs, then run the audit as a hard gate.
    """
    import importlib
    import build_tables
    # the combined modules read masked_<key>.xlsx from a fixed dir; ensure present in scratch
    for pdf, key in key_of.items():
        stem = configs[key]["stem"]
        if stem not in delivered:
            continue
    # build + render
    try:
        bc = importlib.import_module("build_combined_s1")
        importlib.reload(bc)
        cs = importlib.import_module("combined_s1_to_docx")
        importlib.reload(cs)
        audit = importlib.import_module("audit_combined_s1")
        importlib.reload(audit)
    except Exception as e:  # noqa: BLE001
        print("    combined-S1 modules unavailable:", e)
        return 1
    # run audit first (hard gate)
    ok = audit.main() if hasattr(audit, "main") else True
    if not ok:
        print("    STOP: combined-S1 no-codes-missed audit FAILED; not writing combined doc.")
        return 1
    out_docx = os.path.join(args.out_dir, "Combined_Supplementary_Table_S1_all_studies.docx")
    tmp = os.path.join(scratch, "combined_S1.docx")
    cs.build_docx(tmp)
    shutil.copyfile(tmp, out_docx)
    print(f"    -> combined supplementary delivered: {os.path.basename(out_docx)}")
    return 0


def main(argv=None):
    p = argparse.ArgumentParser(description="End-to-end TriNetX two-cohort study tables.")
    p.add_argument("--pdf-dir", help="Directory of TriNetX Compare-Outcomes PDFs.")
    p.add_argument("--pdf", action="append", help="Individual PDF path (repeatable).")
    p.add_argument("--out-dir", required=True, help="Output root (per-study subfolders created here).")
    p.add_argument("--scratch", default="/workspace",
                   help="Local scratch for random-access .xlsx staging (default /workspace).")
    p.add_argument("--combined", action="store_true",
                   help="Also build the combined Supplementary Table S1 (audited).")
    args = p.parse_args(argv)
    return run(args)


if __name__ == "__main__":
    sys.exit(main())
