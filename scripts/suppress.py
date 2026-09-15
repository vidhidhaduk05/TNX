"""
Small-cell (<=10) suppression engine for the cSDH TriNetX workbooks.

Implements the locked policy (SKILL R3b / R8), Table 1 sheet ONLY:
  - THRESHOLD = 10, TRIGGER = "either" (suppress a block if EITHER cohort count <=10)
  - Evaluate the trigger from the ORIGINAL counts BEFORE relabeling.
  - On a triggered block: blank that block's P-value AND ASD to '-', and relabel the
    count token '(n)' -> '(<=10)' keeping the leading '%'. Force text format ('@').
  - Relabel regex anchored on parentheses:  \\(([0-9]|10)\\)  (comma-tolerant parse;
    '(0)' triggers). '(110)' / '(1,069)' are never partially matched.
  - Skip continuous 'mean +/- SD' rows and section headers (no '(n)' token).
  - NEVER touch Table 2 or Table S1.
  - Idempotent (re-running yields 0 changes).
  - Always emit change_log.csv with columns:
        sheet, cell, column_role, old_value, new_value, reason
  - .xlsx is random-access -> write to /workspace, then the caller cp's to /mnt/results.
  - Load with data_only=False.

Config below matches the locked engine; only SRC/OUT_XLSX/OUT_LOG are set per call.
"""
import re
import csv
import shutil
import openpyxl
from openpyxl.utils import get_column_letter

# ---- Locked policy config ----
SHEET_NAMES = ["Table 1"]        # literal; Table 2 / Table S1 never touched
THRESHOLD = 10
TRIGGER = "either"
RELABEL_COUNTS = True
KEEP_PERCENT = True
FIRST_DATA_ROW = 4
# blocks: (count_cols, p_col, asd_col)
BLOCKS = [((2, 3), 4, 5), ((6, 7), 8, 9)]
LE = "\u2264"                     # <= sign (ChrW(8804))
MISS = "-"

# count token anchored on parens: 0..9 or 10 exactly, comma-tolerant inside
COUNT_RE = re.compile(r"\((\d{1,3}(?:,\d{3})*|\d+)\)")


def _count_value(cell_text):
    """Return the integer count inside '(...)' of a '% (n)' cell, else None.
    Comma-tolerant. Continuous/section cells (no paren token) -> None."""
    if cell_text is None:
        return None
    m = COUNT_RE.search(str(cell_text))
    if not m:
        return None
    return int(m.group(1).replace(",", ""))


def _relabel(cell_text):
    """Relabel '(n)' -> '(<=10)' keeping everything else (the leading '%')."""
    return COUNT_RE.sub(f"({LE}10)", str(cell_text), count=1)


def suppress_workbook(src, out_xlsx, out_log, results_copy=None):
    """Apply suppression to `src`, write masked workbook to `out_xlsx` and audit
    to `out_log`. If results_copy given, cp the masked workbook there too.
    Returns (n_changes, list_of_change_rows)."""
    wb = openpyxl.load_workbook(src, data_only=False)
    changes = []

    for sheet in SHEET_NAMES:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for r in range(FIRST_DATA_ROW, ws.max_row + 1):
            for count_cols, p_col, asd_col in BLOCKS:
                # read ORIGINAL counts first (before any relabel)
                cvals = []
                for c in count_cols:
                    cvals.append((_count_value(ws.cell(r, c).value), c))
                present = [(v, c) for v, c in cvals if v is not None]
                if not present:
                    continue  # continuous / section / empty block -> skip
                # trigger: either cohort count <= threshold
                min_v = min(v for v, _ in present)
                if TRIGGER == "either":
                    triggered = min_v <= THRESHOLD
                else:  # 'both'
                    triggered = all(v <= THRESHOLD for v, _ in present)
                if not triggered:
                    continue

                # (a) blank P and ASD to '-'
                for col, role in ((p_col, "P-value"), (asd_col, "ASD")):
                    old = ws.cell(r, col).value
                    if old not in (MISS, None):
                        ws.cell(r, col).value = MISS
                        ws.cell(r, col).number_format = "@"
                        changes.append({
                            "sheet": sheet,
                            "cell": f"{get_column_letter(col)}{r}",
                            "column_role": f"{role} ({'Before' if col in (4,5) else 'After'} PSM)",
                            "old_value": old,
                            "new_value": MISS,
                            "reason": f"either cohort count {LE}{THRESHOLD} (min={min_v})",
                        })
                # (b) relabel each count cell with an original count <= threshold
                #     (only the small cell is relabeled; large partner stays intact)
                for v, c in cvals:
                    if v is None:
                        continue
                    if v <= THRESHOLD and RELABEL_COUNTS:
                        old = ws.cell(r, c).value
                        new = _relabel(old)
                        if new != old:
                            ws.cell(r, c).value = new
                            ws.cell(r, c).number_format = "@"
                            role = f"Cohort {1 if c in (2,6) else 2} count " \
                                   f"({'Before' if c in (2,3) else 'After'} PSM)"
                            changes.append({
                                "sheet": sheet,
                                "cell": f"{get_column_letter(c)}{r}",
                                "column_role": role,
                                "old_value": old,
                                "new_value": new,
                                "reason": f"count {v} {LE}{THRESHOLD}",
                            })

    # write masked workbook to local disk (random-access safe)
    wb.save(out_xlsx)
    # audit trail
    with open(out_log, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["sheet", "cell", "column_role",
                                          "old_value", "new_value", "reason"])
        w.writeheader()
        for row in changes:
            w.writerow(row)
    if results_copy:
        shutil.copy(out_xlsx, results_copy)
    return len(changes), changes


if __name__ == "__main__":
    import sys
    src = sys.argv[1]
    out = sys.argv[2]
    log = sys.argv[3]
    n, _ = suppress_workbook(src, out, log)
    print(f"{src} -> {out}: {n} cells modified; log={log}")
