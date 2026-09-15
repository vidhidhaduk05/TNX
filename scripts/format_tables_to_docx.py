"""
format_tables_to_docx.py
=========================
Convert a multi-sheet Excel workbook of clinical study tables into
publication-formatted Word (.docx) documents.

This reproduces a specific "journal table" house style:
  * Times New Roman, portrait, 1-inch margins.
  * No outer table border and no shading.
  * A vertical rule separating the row-label column from the data columns
    (plus a second vertical rule between grouped blocks, e.g. Before/After PSM).
  * A horizontal rule under the header block.
  * Bold "Table N:" title prefix, bold-italic section rows, centered data values.

Three table archetypes are auto-detected from each sheet's header signature:
  1. baseline-grouped  (e.g. "Table 1": Characteristic | Before PSM x4 | After PSM x4)
  2. outcomes          (e.g. "Table 2": Outcomes | Cohort1 | Cohort2 | EPR | p | HR)
  3. codes             (e.g. "Table S1": Characteristic | Codes)

Public API
----------
build_docx(xlsx_path, out_dir, ...) -> list[str]
    Render all recognized sheets and write .docx file(s); returns output paths.

CLI
---
python format_tables_to_docx.py INPUT.xlsx --out-dir ./out

Dependencies: python-docx, openpyxl  (both standard in the Biomni env).

NOTE on file system: python-docx (a ZIP writer) needs random-access writes,
which S3-FUSE mounts like /mnt/results do not support. Always write to a local
path (e.g. /workspace or tmp) and copy the finished file to /mnt/results.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import warnings
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

# --------------------------------------------------------------------------- #
# Style constants (the locked house style)                                    #
# --------------------------------------------------------------------------- #
FONT_NAME = "Times New Roman"
FONT_SIZE_PT = 11.0
BORDER_SZ = 4          # eighths of a point -> 0.5pt single line
BORDER_VAL = "single"
BORDER_COLOR = "auto"
SUBROW_INDENT = "     "  # 5 spaces, matches the template's sub-bin indent
PAGE_MARGIN_TWIPS = 1440      # 1 inch
PAGE_WIDTH_TWIPS = 15840      # 11 in  (Tabloid portrait, matches the template)
PAGE_HEIGHT_TWIPS = 24480     # 17 in

# GENERAL default renames applied to EVERY study (house-wide conventions the
# user confirmed are always-on). Keep this MINIMAL — only edits that are
# unambiguous and correct for ANY dataset belong here.
GENERAL_RENAME_MAP = {
    # "Nicotine dependence" is reported as "Smoking" house-wide. The narrower
    # ", cigarettes" row is removed and its code folded in by consolidate_nicotine().
    "Nicotine dependence": "Smoking",
    # Age row: drop "at Index" but KEEP a lowercase "mean ± SD" descriptor (house
    # style). All common raw spellings map to the same display string; the
    # _MEAN_SD_KEEP_LABELS exemption stops the normalizer from re-stripping it.
    "Age at Index, Mean \u00b1 SD": "Age, mean \u00b1 SD",
    "Age at Index, mean \u00b1 SD": "Age, mean \u00b1 SD",
    "Age at Index": "Age, mean \u00b1 SD",
}

# Default dataset-specific label edits (IIH/AMD example). Swap for other studies.
# Keyed on the raw Excel label (trimmed); value is the exact display string.
# NOTE: the GENERAL_RENAME_MAP above is merged in on top of whatever rename_map
# the caller supplies, so general house rules survive a per-study override.
DEFAULT_RENAME_MAP = {
    # --- Table 1 row labels ---
    "Black or African American": "Black ",
    "Essential (primary) hypertension": "Essential hypertension",
    "Hemoglobin A1c (HbA1c), Mean \u00b1 SD": "Hemoglobin A1c ",
    "Body mass index (BMI), Mean \u00b1 SD": "Body mass index (BMI)",
    # --- Table S1 right-column code descriptions (manuscript-shortened) ---
    "Age (at least 18 years, most recent occurrence)": "Age (at least 18 years)",
    "ICD-10 G93.2 (Benign intracranial hypertension); on or after Jan 1, 2016":
        "ICD-10 G93.2 ",
}

# Context-specific S1 label edits keyed on the (label, code) PAIR, so that two
# rows sharing a label (e.g. 'Age at Index' appears in both Inclusion and
# Covariates) can be edited independently. Checked before the flat rename map.
DEFAULT_PAIR_RENAME_MAP = {
    ("Age at Index", "Age (at least 18 years, most recent occurrence)"): "Age ",
}

# Row labels whose DATA cells are intentionally left blank in the final document
# (parent rows that only introduce their sub-bins). Dataset-specific.
DEFAULT_BLANK_VALUE_LABELS = {
    "Body mass index (BMI), Mean \u00b1 SD",
    "Hemoglobin A1c (HbA1c), Mean \u00b1 SD",
}

# Section labels rendered bold-italic (no data cells). Detection is structural
# (a label row whose data cells are blank); this set is a safety net.
KNOWN_SECTION_LABELS = {
    "Demographics", "Comorbidities", "Medications", "Laboratory",
    "Inclusion", "Outcomes", "Covariates",
}


# --------------------------------------------------------------------------- #
# Low-level OOXML helpers                                                      #
# --------------------------------------------------------------------------- #
def _set_cell_borders(cell, **sides):
    """Set individual cell borders. sides: top/bottom/left/right -> bool."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(existing)
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        if sides.get(side):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), BORDER_VAL)
            el.set(qn("w:sz"), str(BORDER_SZ))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), BORDER_COLOR)
            tcBorders.append(el)
    if len(tcBorders):
        tcPr.append(tcBorders)


def _set_cell_valign(cell, valign="center"):
    cell.vertical_alignment = (
        WD_ALIGN_VERTICAL.CENTER if valign == "center" else WD_ALIGN_VERTICAL.BOTTOM
    )


def _set_zero_cell_margins(table):
    """Set table-level left/right cell margins to 0 (matches template)."""
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _set_table_width(table, total_twips, col_twips):
    """Fix the table layout and per-column widths (disables autofit jitter)."""
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total_twips))
    tblW.set(qn("w:type"), "dxa")
    tblGrid = table._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        table._tbl.remove(tblGrid)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_twips:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w)))
        tblGrid.append(gc)
    table._tbl.insert(list(table._tbl).index(tblPr) + 1, tblGrid)
    table.autofit = False


def _merge_across(row, start, end):
    """Horizontally merge cells [start..end] in a row; returns the merged cell."""
    merged = row.cells[start]
    for ci in range(start + 1, end + 1):
        merged = merged.merge(row.cells[ci])
    return merged


# --------------------------------------------------------------------------- #
# Cell text writing with run-level formatting                                 #
# --------------------------------------------------------------------------- #
def _style_run(run, *, bold=False, italic=False):
    run.font.name = FONT_NAME
    run.font.size = Pt(FONT_SIZE_PT)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), FONT_NAME)
    run.bold = bold
    run.italic = italic


def _title_prefix_split(text):
    """Split a table title into (bold_prefix, rest).

    The house style bolds the "Table N" / "Table SN" label and its terminating
    punctuation. That terminator may be a colon OR a period (e.g. "Table 1:" or
    "Table 1." / "Table S1."), so we anchor on the leading 'Table <id>' token
    and bold up to and including the first ':' or '.' that follows it. If no such
    delimiter is found, fall back to the first ':' or '.' anywhere; otherwise no
    prefix is bolded.
    """
    m = re.match(r"^(\s*Table\s+S?\d+\s*[:.])", text, re.IGNORECASE)
    if m:
        idx = m.end()
        return text[:idx], text[idx:]
    # Fallback: first ':' or '.' anywhere in the string.
    cands = [text.index(ch) for ch in (":", ".") if ch in text]
    if cands:
        idx = min(cands) + 1
        return text[:idx], text[idx:]
    return None, text


def _write_cell(cell, text, *, bold=False, italic=False, align="left",
                valign="center", title_prefix=False):
    """Write text into a cell with the house font/format.

    text may contain '\n' for hard line breaks.
    title_prefix=True bolds the "Table N"/"Table SN" label and its terminating
    ':' or '.' (see _title_prefix_split); the remainder is regular weight.
    """
    para = cell.paragraphs[0]
    for r in list(para.runs):
        r._element.getparent().remove(r._element)
    para.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    )
    _set_cell_valign(cell, valign)

    text = "" if text is None else str(text)

    if title_prefix:
        prefix, rest = _title_prefix_split(text)
        if prefix is not None:
            _style_run(para.add_run(prefix), bold=True, italic=italic)
            last = para.add_run(rest) if rest else para.add_run("")
            if rest:
                _style_run(last, bold=False, italic=italic)
            # The house style terminates the title line with a hard break so the
            # title cell text ends in '\n' (matches the example documents exactly).
            last.add_break()
            return

    lines = text.split("\n")
    for li, line in enumerate(lines):
        run = para.add_run(line)
        _style_run(run, bold=bold, italic=italic)
        if li < len(lines) - 1:
            run.add_break()


# --------------------------------------------------------------------------- #
# Text normalization                                                          #
# --------------------------------------------------------------------------- #
_MEAN_SD_RE = re.compile(r",\s*mean\s*\u00b1\s*sd\s*$", re.IGNORECASE)
_PLUS_RE = re.compile(r"(\d+)\+")          # capture the FULL number (40+ -> >=40)
_LOWER_WORD_RE = re.compile(r"^[a-z][a-z]+$")

# --- Safe mechanical cleanups (learned from user hand-edits; see SKILL.md "R5") ---
# These two are the ONLY label rewrites that generalize without semantic risk.
# A trailing ", unspecified" is a redundant ICD gloss the manuscript drops
# (e.g. "Hyperlipidemia, unspecified" -> "Hyperlipidemia").
_UNSPECIFIED_RE = re.compile(r",\s*unspecified\s*$", re.IGNORECASE)
# A literal "(both cohorts)" parenthetical is a structural gloss, not clinical content
# (e.g. "Carotid stenosis (both cohorts)" -> "Carotid stenosis").
_BOTH_COHORTS_RE = re.compile(r"\s*\(both cohorts\)", re.IGNORECASE)

# Labels EXEMPT from the ", mean ± SD" strip: the descriptor is kept here by
# explicit user preference (the Age row carries "mean ± SD" in the manuscript).
# Compared case-insensitively against the post-rename core.
_MEAN_SD_KEEP_LABELS = {"age, mean \u00b1 sd"}


def normalize_label(raw, rename_map, normalize=True):
    """Apply (1) explicit rename_map then (2) general normalizers to a row label.

    Returns the display string (may include a leading 5-space indent for sub-rows).
    """
    if raw is None:
        return ""
    s = str(raw)
    stripped = s.lstrip(" ")
    had_indent = (len(s) - len(stripped)) >= 2

    # (1) explicit rename map keyed on raw then trimmed text
    if s in rename_map:
        return rename_map[s]
    if stripped in rename_map:
        out = rename_map[stripped]
        if had_indent and not out.startswith(" "):
            out = SUBROW_INDENT + out
        return out

    core = stripped
    if normalize:
        # Strip ", mean ± SD" EXCEPT for explicitly exempted labels (e.g. Age).
        if core.strip().lower() not in _MEAN_SD_KEEP_LABELS:
            core = _MEAN_SD_RE.sub("", core)
        core = _UNSPECIFIED_RE.sub("", core)          # ", unspecified" -> ""
        core = _BOTH_COHORTS_RE.sub("", core)         # "(both cohorts)" -> ""
        core = _PLUS_RE.sub("\u2265" + r"\1", core)   # 40+ -> >=40 ; 9+ -> >=9
        if _LOWER_WORD_RE.match(core):
            core = core.capitalize()
        core = re.sub(r"\s{2,}", " ", core).strip()

    if had_indent:
        core = SUBROW_INDENT + core
    return core


def _clean_value(v):
    """Coerce a numeric/None cell to a clean display string (values pass through)."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return ("%g" % v)
    return str(v)


# --------------------------------------------------------------------------- #
# Sheet parsing                                                               #
# --------------------------------------------------------------------------- #
@dataclass
class ParsedSheet:
    kind: str
    title: str = ""
    rows: list = field(default_factory=list)
    meta: dict = field(default_factory=dict)


def detect_table_type(ws):
    """Classify a worksheet by its header signature."""
    max_r = min(ws.max_row, 6)
    max_c = ws.max_column
    blob = []
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                blob.append(v.lower())
    text = " | ".join(blob)

    if ("before" in text and "after" in text) or ("asd" in text and max_c >= 8):
        return "baseline"
    if "event probability" in text or ("p-value" in text and "hr" in text) \
       or ("hr [95%ci]" in text):
        return "outcomes"
    if max_c <= 2 or "code" in text:
        return "codes"
    return "unknown"


def _parse_cohort_label(s):
    """Split 'IIH with AMD (N=34,764)' -> ('IIH with AMD', 'N=34,764')."""
    if s is None:
        return "", ""
    s = str(s).replace("\n", " ").strip()
    m = re.search(r"\(([^)]*)\)\s*$", s)
    if m:
        return s[: m.start()].strip(), m.group(1).strip()
    return s, ""


def derive_cohort_names(wb, cohort_map=None):
    """Derive semantic cohort names from the outcomes (Table 2) sub-header.

    Returns c1_name / c2_name (display, 'no'->'without') and
    c1_name_raw / c2_name_raw (original, for the 'vs' string).
    """
    names = {
        "c1_name": "Cohort 1", "c2_name": "Cohort 2",
        "c1_name_raw": "Cohort 1", "c2_name_raw": "Cohort 2",
    }
    for ws in wb.worksheets:
        if detect_table_type(ws) != "outcomes":
            continue
        for r in range(1, min(ws.max_row, 8) + 1):
            b = ws.cell(row=r, column=2).value
            c = ws.cell(row=r, column=3).value
            if isinstance(b, str) and isinstance(c, str) and ("n=" in b.lower()):
                c1_raw, _ = _parse_cohort_label(b)
                c2_raw, _ = _parse_cohort_label(c)
                names["c1_name_raw"] = c1_raw
                names["c2_name_raw"] = c2_raw
                names["c1_name"] = re.sub(r"\bno\b", "without", c1_raw)
                names["c2_name"] = re.sub(r"\bno\b", "without", c2_raw)
                break
        break
    if cohort_map:
        names.update({k: v for k, v in cohort_map.items() if k in names})
    return names


def parse_baseline_sheet(ws):
    """Parse a baseline-grouped sheet (Table 1).

    Layout (rows): title; group-header (Before/After); sub-header (cohort/N/P/ASD);
    then alternating section rows (col A only) and data rows (A..I).
    Returns ParsedSheet with rows = list of ('section'|'data', label, [8 values]).
    meta carries group labels and the 4 N strings.
    """
    title = ws.cell(row=1, column=1).value or ""
    # find the sub-header row (has 'P-value'/'ASD' or 'Cohort')
    sub_r = None
    for r in range(2, min(ws.max_row, 6) + 1):
        rowtext = " ".join(
            str(ws.cell(row=r, column=c).value or "") for c in range(1, ws.max_column + 1)
        ).lower()
        if "cohort" in rowtext or ("p-value" in rowtext and "asd" in rowtext):
            sub_r = r
            break
    if sub_r is None:
        sub_r = 3
    group_r = sub_r - 1

    # group labels (Before/After) from the row above the sub-header
    group_labels = []
    for c in range(2, ws.max_column + 1):
        v = ws.cell(row=group_r, column=c).value
        if v:
            group_labels.append((c, str(v)))

    # Ns from the sub-header cohort cells (cols B,C and F,G in the 8-col layout)
    def _n_of(col):
        _, n = _parse_cohort_label(ws.cell(row=sub_r, column=col).value)
        return n
    meta = {
        "group_labels": group_labels,
        "n_b1": _n_of(2), "n_b2": _n_of(3),
        "n_a1": _n_of(6), "n_a2": _n_of(7),
    }

    rows = []
    for r in range(sub_r + 1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a is None:
            continue
        vals = [ws.cell(row=r, column=c).value for c in range(2, 10)]  # B..I = 8 cols
        nonempty = [v for v in vals if v not in (None, "")]
        if not nonempty:
            rows.append(("section", str(a), [""] * 8))
        else:
            rows.append(("data", a, [_clean_value(v) for v in vals]))
    return ParsedSheet(kind="baseline", title=str(title), rows=rows, meta=meta)


def _maybe_blank_values(label, vals, blank_labels):
    """Return blanked values if this label is configured as a label-only row."""
    if blank_labels and str(label).strip() in blank_labels:
        return [""] * len(vals)
    return vals


def parse_outcomes_blocks(ws):
    """Parse all stacked outcome blocks in a Table-2 sheet.

    Returns list of dicts: {title, follow_up_years, c1_n, c2_n, comparison, rows}
    where rows = list of (label, [c1, c2, epr, p, hr]).
    """
    blocks = []
    r = 1
    max_r = ws.max_row
    while r <= max_r:
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.lower().startswith("table"):
            title = v.strip()
            header_r = r + 1
            sub_r = r + 2
            c1_name, c1_n = _parse_cohort_label(ws.cell(row=sub_r, column=2).value)
            c2_name, c2_n = _parse_cohort_label(ws.cell(row=sub_r, column=3).value)
            comparison = ws.cell(row=sub_r, column=4).value or ""
            # follow-up years from the title ("within N years")
            m = re.search(r"within\s+(\d+)\s*year", title, re.IGNORECASE)
            fy = int(m.group(1)) if m else -1
            # data rows from sub_r+1 until a blank label or next 'Table'
            rows = []
            rr = sub_r + 1
            while rr <= max_r:
                lab = ws.cell(row=rr, column=1).value
                if lab is None:
                    rr += 1
                    # stop if we've hit a gap AND the next non-empty is a new Table
                    nxt = ws.cell(row=rr, column=1).value if rr <= max_r else None
                    if isinstance(nxt, str) and nxt.lower().startswith("table"):
                        break
                    if lab is None and (rr > max_r or ws.cell(row=rr, column=2).value is None):
                        # allow a single trailing blank then break on next table
                        continue
                    continue
                if isinstance(lab, str) and lab.lower().startswith("table"):
                    break
                vals = [ws.cell(row=rr, column=c).value for c in range(2, 7)]  # B..F
                rows.append((str(lab), [_clean_value(x) for x in vals]))
                rr += 1
            blocks.append({
                "title": title, "follow_up_years": fy,
                "c1_name_raw": c1_name, "c2_name_raw": c2_name,
                "c1_n": c1_n, "c2_n": c2_n, "comparison": str(comparison),
                "rows": rows,
            })
            r = rr
        else:
            r += 1
    return blocks


def parse_codes_sheet(ws):
    """Parse a 2-column codes sheet (Table S1)."""
    title = ws.cell(row=1, column=1).value or ""
    # header row = row 2 (Characteristic | Codes...)
    header = (ws.cell(row=2, column=1).value, ws.cell(row=2, column=2).value)
    rows = []
    for r in range(3, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a is None and b is None:
            continue
        if (b is None or str(b).strip() == ""):
            rows.append(("section", str(a), ""))
        else:
            rows.append(("data", str(a), str(b)))
    return ParsedSheet(kind="codes", title=str(title), rows=rows,
                       meta={"header": header})


# --------------------------------------------------------------------------- #
# Post-parse content transforms                                               #
# --------------------------------------------------------------------------- #
# Recognizes the broad "Nicotine dependence" parent row and the narrower
# ", cigarettes" child row (any label that is the parent text + a ", ...ycigarettes"
# qualifier). Matching is on the RAW Excel label, before rename.
_NICOTINE_PARENT_RE = re.compile(r"^\s*nicotine dependence\s*$", re.IGNORECASE)
_NICOTINE_CHILD_RE = re.compile(
    r"^\s*nicotine dependence\s*,\s*.*cigarett", re.IGNORECASE)


def _row_label(row):
    """Return the label element of a parsed row tuple (works for all kinds)."""
    # baseline/codes: ("section"|"data", label, values); some code rows ("data", label, codestr)
    if len(row) >= 2 and isinstance(row[1], str):
        return row[1]
    # outcomes block rows: (label, [values])
    if len(row) >= 1 and isinstance(row[0], str):
        return row[0]
    return ""


def consolidate_nicotine(parsed_rows, kind):
    """Collapse the two nicotine rows into a single 'Smoking' row (house rule).

    The narrower 'Nicotine dependence, cigarettes' (e.g. ICD-10 F17.21) row is
    removed; the broader 'Nicotine dependence' (e.g. ICD-10 F17) row survives and
    is later renamed to 'Smoking' via GENERAL_RENAME_MAP. For codes tables the
    removed child's code is appended to the surviving row so no code is lost
    ('ICD-10 F17' -> 'ICD-10 F17, ICD-10 F17.21').

    `parsed_rows` is the list stored on a ParsedSheet (baseline/codes) OR an
    outcomes block's "rows" list. Returns a NEW list; input is not mutated.
    `kind` is 'baseline' | 'codes' | 'outcomes'.
    """
    # locate child + parent indices
    child_idx = parent_idx = None
    for i, row in enumerate(parsed_rows):
        lab = _row_label(row)
        if _NICOTINE_CHILD_RE.match(lab):
            child_idx = i
        elif _NICOTINE_PARENT_RE.match(lab):
            parent_idx = i

    if child_idx is None:
        return list(parsed_rows)  # nothing to consolidate

    child_code = ""
    if kind == "codes":
        crow = parsed_rows[child_idx]
        # codes data row is ("data", label, codestring)
        if len(crow) >= 3 and isinstance(crow[2], str):
            child_code = crow[2].strip()

    out = []
    for i, row in enumerate(parsed_rows):
        if i == child_idx:
            continue  # drop the narrower row
        if i == parent_idx and kind == "codes" and child_code:
            # append child's code to the surviving parent's code cell
            lab = row[1]
            parent_code = row[2].strip() if len(row) >= 3 and isinstance(row[2], str) else ""
            merged = parent_code
            if child_code and child_code not in parent_code:
                merged = f"{parent_code}, {child_code}" if parent_code else child_code
            out.append(("data", lab, merged))
        else:
            out.append(row)

    # If there was a child but no broad parent row, promote the child itself so its
    # data/codes are not lost (rename to Smoking still applies via the label).
    if parent_idx is None:
        crow = parsed_rows[child_idx]
        if kind == "codes":
            out.insert(min(child_idx, len(out)), ("data", "Nicotine dependence", child_code))
        elif kind == "baseline":
            out.insert(min(child_idx, len(out)), ("data", "Nicotine dependence", crow[2]))
        else:  # outcomes
            out.insert(min(child_idx, len(out)), ("Nicotine dependence", crow[1]))
    return out


# --------------------------------------------------------------------------- #
# Renderers                                                                   #
# --------------------------------------------------------------------------- #
# Column-width templates (twips), copied from the example documents.
BASELINE_COL_TW = [3419, 1820, 2102, 872, 636, 1625, 1961, 868, 633]   # 9 cols
OUTCOMES_COL_TW = [2605, 2070, 2165, 3510, 1440, 2148]                  # 6 cols
CODES_COL_TW = [3060, 6300]                                            # 2 cols


def _new_table(doc, n_rows, n_cols, col_tw):
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = "Table Grid"          # start from a bordered grid...
    # ...then clear ALL borders; we add only the ones the house style needs.
    tblPr = table._tbl.tblPr
    for b in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(b)
    no_borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        no_borders.append(el)
    tblPr.append(no_borders)
    _set_zero_cell_margins(table)
    _set_table_width(table, sum(col_tw), col_tw)
    table.alignment = None
    return table


def render_baseline_table(doc, parsed, names):
    """Render a Table-1 baseline-grouped table into doc."""
    rows = parsed.rows
    n_body = len(rows)
    n_rows = 3 + n_body            # title + group-hdr + sub-hdr + body
    t = _new_table(doc, n_rows, 9, BASELINE_COL_TW)

    # ---- Row 0: title spanning all 9 cols ----
    title_cell = _merge_across(t.rows[0], 0, 8)
    _write_cell(title_cell, parsed.title, align="left", valign="bottom",
                title_prefix=True)

    # ---- Row 1: group headers; col 0 vMerge restart (label) ----
    r1 = t.rows[1]
    # group label text (Before PSM / After PSM) - derive from meta
    g_labels = [g[1] for g in parsed.meta.get("group_labels", [])]
    before_lbl = g_labels[0] if len(g_labels) > 0 else "Before"
    after_lbl = g_labels[1] if len(g_labels) > 1 else "After"
    g1 = _merge_across(r1, 1, 4)
    _write_cell(g1, before_lbl, bold=True, align="center", valign="center")
    g2 = _merge_across(r1, 5, 8)
    _write_cell(g2, after_lbl, bold=True, align="center", valign="center")

    # ---- Row 2: sub-header (cohort / P-value / ASD) x2 ----
    r2 = t.rows[2]
    m = parsed.meta
    sub_texts = [
        None,  # col 0 is vMerged with row1
        f"{names['c1_name']}\n({m['n_b1']})" if m['n_b1'] else names['c1_name'],
        f"{names['c2_name']}\n({m['n_b2']})" if m['n_b2'] else names['c2_name'],
        "P-value", "ASD",
        f"{names['c1_name']}\n({m['n_a1']})" if m['n_a1'] else names['c1_name'],
        f"{names['c2_name']}\n({m['n_a2']})" if m['n_a2'] else names['c2_name'],
        "P-value", "ASD",
    ]
    for ci in range(1, 9):
        _write_cell(r2.cells[ci], sub_texts[ci], bold=True, align="center",
                    valign="center")

    # ---- col-0 vertical merge across rows 1 & 2, italic header ----
    c0_merged = t.rows[1].cells[0].merge(t.rows[2].cells[0])
    _write_cell(c0_merged, "Characteristics \u2014 mean \u00b1 SD, %(n)",
                italic=True, align="left", valign="bottom")

    # ---- Body rows ----
    for i, (kind, label, vals) in enumerate(rows):
        ri = 3 + i
        row = t.rows[ri]
        if kind == "section":
            _write_cell(row.cells[0], label, bold=True, italic=True,
                        align="left", valign="bottom")
            for ci in range(1, 9):
                _write_cell(row.cells[ci], "", align="center")
        else:
            disp = normalize_label(label, parsed.meta.get("_rename", {}),
                                   parsed.meta.get("_normalize", True))
            shown_vals = _maybe_blank_values(
                label, vals, parsed.meta.get("_blank_labels", set()))
            _write_cell(row.cells[0], disp, align="left", valign="center")
            for ci in range(1, 9):
                _write_cell(row.cells[ci], shown_vals[ci - 1], align="center",
                            valign="center")

    _apply_baseline_borders(t)
    return t


def _apply_baseline_borders(table):
    """Border pattern for the 3-row-header grouped baseline table (Table 1).

    Matches the example exactly:
      * col 0 (label): Right border on every row except the title row; no
        horizontal rules pass through the label column.
      * Row 1 (group band): each band cell (cols 1-4, 5-8) gets Bottom + the
        vertical dividers (Left/Right) -> a box under 'Before/After PSM'.
      * Row 2 (sub-header): Top + Bottom + vertical dividers.
      * Row 3 (first body row): Top + vertical dividers.
      * All other body rows: vertical dividers only.
    Vertical dividers: Left on cols 1 & 5, Right on cols 0, 4 & 8.
    """
    n_rows = len(table.rows)

    def vdiv(ci):
        s = {}
        if ci == 0:
            s["right"] = True
        if ci == 1:
            s["left"] = True
        if ci == 4:
            s["right"] = True
        if ci == 5:
            s["left"] = True
        if ci == 8:
            pass  # col 8 is the outer edge; no right border in the example
        return s

    for ri in range(n_rows):
        row = table.rows[ri]
        if ri == 0:
            continue
        for ci in range(9):
            sides = vdiv(ci)
            if ri == 1 and ci >= 1:
                sides["bottom"] = True               # under group band
            if ri == 2 and ci >= 1:
                sides["top"] = True
                sides["bottom"] = True               # main header rule
            if ri == 3:
                sides["top"] = True                  # top side of header rule
                                                     # (incl. col 0 -> 'TR')
            _merge_cell_borders(row.cells[ci], sides)


def render_outcomes_table(doc, block, names):
    """Render one Table-2 outcomes block into doc."""
    rows = block["rows"]
    n_rows = 2 + len(rows)         # title + header + body
    t = _new_table(doc, n_rows, 6, OUTCOMES_COL_TW)

    # ---- Row 0: title ----
    tc = _merge_across(t.rows[0], 0, 5)
    _write_cell(tc, block["title"], align="left", valign="bottom", title_prefix=True)

    # ---- Row 1: header ----
    c1 = f"{names['c1_name']} ({block['c1_n']})" if block['c1_n'] else names['c1_name']
    c2 = f"{names['c2_name']} ({block['c2_n']})" if block['c2_n'] else names['c2_name']
    comparison = block.get("comparison", "")
    # NB: the template renders a trailing space after "rate" before the line break
    epr = "Event probability rate " + (f"\n{comparison}" if comparison else "")
    header = ["Outcomes %(n)", c1, c2, epr, "p-value", "HR [95%CI]"]
    hr = t.rows[1]
    _write_cell(hr.cells[0], header[0], italic=True, align="left", valign="bottom")
    for ci in range(1, 6):
        _write_cell(hr.cells[ci], header[ci], bold=True, align="center",
                    valign="center")

    # ---- Body ----
    for i, (label, vals) in enumerate(rows):
        ri = 2 + i
        row = t.rows[ri]
        _write_cell(row.cells[0], str(label), align="left", valign="center")
        for ci in range(1, 6):
            _write_cell(row.cells[ci], vals[ci - 1], align="center", valign="center")

    _apply_grouped_borders(t, n_header_rows=2, label_col=0, divider_after=None)
    return t


def render_codes_table(doc, parsed):
    """Render a Table-S1 codes table into doc."""
    rows = parsed.rows
    n_rows = 2 + len(rows)         # title + header + body
    t = _new_table(doc, n_rows, 2, CODES_COL_TW)

    tc = _merge_across(t.rows[0], 0, 1)
    _write_cell(tc, parsed.title, align="left", valign="bottom", title_prefix=True)

    header = parsed.meta.get("header", ("Characteristic", "Codes"))
    hr = t.rows[1]
    _write_cell(hr.cells[0], header[0] or "Characteristic", bold=True,
                align="left", valign="bottom")
    _write_cell(hr.cells[1], header[1] or "Codes", bold=True,
                align="left", valign="bottom")

    rename = parsed.meta.get("_rename", {})
    pair_rename = parsed.meta.get("_pair_rename", {})
    normalize = parsed.meta.get("_normalize", True)
    for i, (kind, label, code) in enumerate(rows):
        ri = 2 + i
        row = t.rows[ri]
        if kind == "section":
            _write_cell(row.cells[0], label, bold=True, italic=True,
                        align="left", valign="bottom")
            _write_cell(row.cells[1], "", align="left")
        else:
            lab_key = str(label).strip()
            code_key = str(code).strip()
            # 1) context-specific (label, code) pair edit wins
            if (lab_key, code_key) in pair_rename:
                disp_label = pair_rename[(lab_key, code_key)]
            else:
                disp_label = normalize_label(label, rename, normalize)
            # code text: apply rename map only (exact manuscript edits), no
            # general label normalizers (those are for row labels, not codes).
            disp_code = rename.get(code_key, code)
            _write_cell(row.cells[0], disp_label, align="left", valign="center")
            _write_cell(row.cells[1], disp_code, align="left", valign="center")

    _apply_grouped_borders(t, n_header_rows=2, label_col=0, divider_after=None)
    return t


def _apply_grouped_borders(table, n_header_rows, label_col=0, divider_after=None):
    """Apply the house border pattern.

    * Vertical rule between label_col and the next column (label_col Right,
      label_col+1 Left) on EVERY row except the title row (row 0).
    * Optional second vertical rule after column `divider_after`
      (that col Right, next col Left) -> used for Before/After PSM in Table 1.
    * Horizontal rule under the header block: header rows get Bottom,
      first body row gets Top, across all columns.
    """
    n_rows = len(table.rows)
    n_cols = len(table.columns)
    header_last = n_header_rows - 1
    first_body = n_header_rows

    for ri, row in enumerate(table.rows):
        if ri == 0:
            continue  # title row: no borders
        # gather distinct cells (merged group headers share a _tc)
        for ci in range(n_cols):
            cell = row.cells[ci]
            sides = {}
            # vertical label divider
            if ci == label_col:
                sides["right"] = True
            if ci == label_col + 1:
                sides["left"] = True
            # secondary divider (grouped blocks)
            if divider_after is not None:
                if ci == divider_after:
                    sides["right"] = True
                if ci == divider_after + 1:
                    sides["left"] = True
            # horizontal rule under header
            if ri == header_last:
                sides["bottom"] = True
            if ri == first_body:
                sides["top"] = True
            # merge with any existing borders already set on shared cells
            _merge_cell_borders(cell, sides)


def _merge_cell_borders(cell, sides):
    """Like _set_cell_borders but unions with borders already present."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcBorders"))
    have = {}
    if existing is not None:
        for side in ("top", "bottom", "left", "right"):
            e = existing.find(qn(f"w:{side}"))
            if e is not None and e.get(qn("w:val")) not in (None, "nil"):
                have[side] = True
    have.update({k: v for k, v in sides.items() if v})
    _set_cell_borders(cell, **have)


# --------------------------------------------------------------------------- #
# Orchestrator                                                                #
# --------------------------------------------------------------------------- #
def _safe_write_docx(doc, final_path):
    """Write a docx to a local temp file then copy to final_path (S3-FUSE safe)."""
    os.makedirs(os.path.dirname(os.path.abspath(final_path)), exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    doc.save(tmp)
    shutil.copy(tmp, final_path)
    os.remove(tmp)
    return final_path


def _setup_document():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    # Tabloid portrait (11 x 17 in) -- matches the example documents and gives
    # the wide 9-column baseline table room without shrinking the font.
    sec.page_width = Twips(PAGE_WIDTH_TWIPS)
    sec.page_height = Twips(PAGE_HEIGHT_TWIPS)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, attr, Twips(PAGE_MARGIN_TWIPS))
    # base style font
    normal = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(FONT_SIZE_PT)
    return doc


def _select_outcome_blocks(blocks, follow_up="longest"):
    """Pick which outcome blocks to render.

    follow_up: 'longest' (default) | 'all' | int (years) | list[int].
    """
    if not blocks:
        return []
    if follow_up == "all":
        return blocks
    if follow_up == "longest":
        return [max(blocks, key=lambda b: (b["follow_up_years"], blocks.index(b)))]
    if isinstance(follow_up, int):
        sel = [b for b in blocks if b["follow_up_years"] == follow_up]
        return sel or [max(blocks, key=lambda b: b["follow_up_years"])]
    if isinstance(follow_up, (list, tuple, set)):
        return [b for b in blocks if b["follow_up_years"] in set(follow_up)]
    return [max(blocks, key=lambda b: b["follow_up_years"])]


def build_docx(xlsx_path, out_dir, *, cohort_map=None,
               rename_map=None, pair_rename_map=None,
               normalize=True, blank_value_labels=None,
               file_grouping="match_uploads", follow_up="longest",
               consolidate_nicotine_rows=True,
               apply_general_renames=True,
               main_filename="Table_1_and_2.docx",
               supp_filename="Table_S1.docx",
               combined_filename="All_Tables.docx"):
    """Render all recognized sheets of an Excel workbook into .docx file(s).

    Parameters
    ----------
    xlsx_path : str           Path to the input .xlsx.
    out_dir   : str           Directory for output .docx files.
    cohort_map : dict|None    Override auto-derived cohort names
                              (keys: c1_name, c2_name, c1_name_raw, c2_name_raw).
    rename_map : dict|None    Explicit row-label substitutions. Defaults to
                              DEFAULT_RENAME_MAP (the IIH/AMD example edits).
                              Pass {} to disable custom renames.
    normalize : bool          Apply general label cleanups (strip ", Mean ± SD"
                              except exempted Age row; ", unspecified" -> "";
                              "(both cohorts)" -> ""; N+ -> >=N; title-case lone
                              drug names).
    consolidate_nicotine_rows : bool
                              Collapse 'Nicotine dependence, cigarettes' into the
                              broader 'Nicotine dependence' row (later renamed
                              'Smoking'); merge the dropped code in codes tables.
                              Default True (house rule). Set False to keep both.
    apply_general_renames : bool
                              Layer GENERAL_RENAME_MAP (Nicotine->Smoking, Age row)
                              on top of rename_map. Default True. Caller's explicit
                              rename_map entries override on key clash.
    file_grouping : str       'match_uploads' (Tables 1&2 together, S1 separate),
                              'combined' (one file), or 'per_table'.
    follow_up : str|int|list  Which Table-2 block(s): 'longest' (default), 'all',
                              a year int, or a list of year ints.

    Returns
    -------
    list[str] : the output file paths written.
    """
    if rename_map is None:
        rename_map = dict(DEFAULT_RENAME_MAP)
    if pair_rename_map is None:
        pair_rename_map = dict(DEFAULT_PAIR_RENAME_MAP)
    if blank_value_labels is None:
        blank_value_labels = set(DEFAULT_BLANK_VALUE_LABELS)

    # General house-wide renames (e.g. Nicotine dependence -> Smoking, Age row)
    # are layered on top of whatever rename_map the caller passed, so they apply
    # even when a per-study map is supplied (or {} to clear study-specific edits).
    if apply_general_renames:
        merged = dict(GENERAL_RENAME_MAP)
        merged.update(rename_map)   # caller's explicit entries win on key clash
        rename_map = merged

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    names = derive_cohort_names(wb, cohort_map=cohort_map)

    # classify and parse every sheet, preserving workbook order
    parsed_items = []  # list of ('baseline', ParsedSheet) | ('outcomes', block) | ('codes', ParsedSheet)
    for ws in wb.worksheets:
        kind = detect_table_type(ws)
        if kind == "baseline":
            ps = parse_baseline_sheet(ws)
            if consolidate_nicotine_rows:
                ps.rows = consolidate_nicotine(ps.rows, "baseline")
            ps.meta["_rename"] = rename_map
            ps.meta["_normalize"] = normalize
            ps.meta["_blank_labels"] = blank_value_labels
            parsed_items.append(("baseline", ps))
        elif kind == "outcomes":
            blocks = parse_outcomes_blocks(ws)
            for b in _select_outcome_blocks(blocks, follow_up):
                if consolidate_nicotine_rows:
                    b["rows"] = consolidate_nicotine(b["rows"], "outcomes")
                parsed_items.append(("outcomes", b))
        elif kind == "codes":
            ps = parse_codes_sheet(ws)
            if consolidate_nicotine_rows:
                ps.rows = consolidate_nicotine(ps.rows, "codes")
            ps.meta["_rename"] = rename_map
            ps.meta["_pair_rename"] = pair_rename_map
            ps.meta["_normalize"] = normalize
            parsed_items.append(("codes", ps))
        else:
            warnings.warn(f"Sheet '{ws.title}' not recognized (skipped).")

    if not parsed_items:
        raise ValueError("No recognizable tables found in the workbook.")

    os.makedirs(out_dir, exist_ok=True)
    written = []

    def _render_one(doc, item, first):
        kind, obj = item
        if not first:
            doc.add_paragraph("")  # spacer between stacked tables
        if kind == "baseline":
            render_baseline_table(doc, obj, names)
        elif kind == "outcomes":
            render_outcomes_table(doc, obj, names)
        elif kind == "codes":
            render_codes_table(doc, obj)

    if file_grouping == "combined":
        doc = _setup_document()
        for i, item in enumerate(parsed_items):
            _render_one(doc, item, first=(i == 0))
        written.append(_safe_write_docx(doc, os.path.join(out_dir, combined_filename)))

    elif file_grouping == "per_table":
        for i, item in enumerate(parsed_items):
            doc = _setup_document()
            _render_one(doc, item, first=True)
            kind = item[0]
            fname = f"{i+1:02d}_{kind}.docx"
            written.append(_safe_write_docx(doc, os.path.join(out_dir, fname)))

    else:  # match_uploads: (baseline + outcomes) together; codes separate
        main_items = [it for it in parsed_items if it[0] in ("baseline", "outcomes")]
        supp_items = [it for it in parsed_items if it[0] == "codes"]
        if main_items:
            doc = _setup_document()
            for i, item in enumerate(main_items):
                _render_one(doc, item, first=(i == 0))
            written.append(_safe_write_docx(doc, os.path.join(out_dir, main_filename)))
        for j, item in enumerate(supp_items):
            doc = _setup_document()
            _render_one(doc, item, first=True)
            fname = supp_filename if len(supp_items) == 1 else f"Table_S{j+1}.docx"
            written.append(_safe_write_docx(doc, os.path.join(out_dir, fname)))

    return written


# --------------------------------------------------------------------------- #
# CLI                                                                         #
# --------------------------------------------------------------------------- #
def _main(argv=None):
    p = argparse.ArgumentParser(
        description="Format a multi-sheet clinical-study Excel workbook into Word.")
    p.add_argument("xlsx", help="Input .xlsx path")
    p.add_argument("--out-dir", default="./formatted_tables", help="Output directory")
    p.add_argument("--grouping", default="match_uploads",
                   choices=["match_uploads", "combined", "per_table"])
    p.add_argument("--follow-up", default="longest",
                   help="'longest' (default), 'all', or a year integer")
    p.add_argument("--no-normalize", action="store_true",
                   help="Disable general label cleanups")
    p.add_argument("--no-rename", action="store_true",
                   help="Disable the default custom rename map")
    p.add_argument("--no-consolidate-nicotine", action="store_true",
                   help="Keep both nicotine rows instead of merging into 'Smoking'")
    p.add_argument("--no-general-renames", action="store_true",
                   help="Disable house-wide renames (Nicotine->Smoking, Age row)")
    args = p.parse_args(argv)

    fu = args.follow_up
    if fu.isdigit():
        fu = int(fu)
    out = build_docx(
        args.xlsx, args.out_dir,
        rename_map={} if args.no_rename else None,
        normalize=not args.no_normalize,
        consolidate_nicotine_rows=not args.no_consolidate_nicotine,
        apply_general_renames=not args.no_general_renames,
        file_grouping=args.grouping,
        follow_up=fu,
    )
    for f in out:
        print("Wrote:", f)


if __name__ == "__main__":
    _main()
