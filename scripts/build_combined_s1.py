"""
build_combined_s1.py

Builds a single deduplicated Supplementary Table S1 combining all uploaded reports
per Rule 15:
- Sections: Inclusion, Exclusions, Outcomes, Covariates
- Added 'Applies to' column indicating which study/comparison uses each code or row
- Produces Combined_Table_S1.xlsx and Combined_Table_S1.docx
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import docx
from docx.shared import Pt, Inches, Inches, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from format_tables_to_docx import (
    _new_table, _merge_across, _write_cell, _apply_grouped_borders,
    PAGE_MARGIN_TWIPS, PAGE_WIDTH_TWIPS, PAGE_HEIGHT_TWIPS
)

STUDY_MAPPING = [
    ("PLT__100__MM_vs_MMAE_202607261353", "MM vs MMAE"),
    ("PLT__100__Sx_vs_MMAE_202607261455", "Sx vs MMAE"),
    ("PLT__100__Sx_vs_Sx_+_MMAE_202607261402", "Sx vs Sx + MMAE"),
]

def build_combined_s1(output_dir: str = "output"):
    # Collect S1 rows from each study's clean/masked Excel
    # Structure: section -> dict of (char, codes) -> set of study tags
    sections_order = ["Inclusion", "Exclusions", "Outcomes", "Covariates"]
    collected = {sec: {} for sec in sections_order}

    for stem, tag in STUDY_MAPPING:
        xlsx_path = os.path.join(output_dir, stem, f"{stem}_Table1_2_S1.xlsx")
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb['Table S1']
        
        curr_sec = None
        for r in range(3, ws.max_row + 1):
            val_a = str(ws.cell(r, 1).value or '').strip()
            val_b = str(ws.cell(r, 2).value or '').strip()
            
            if val_a in sections_order and not val_b:
                curr_sec = val_a
                continue
                
            if curr_sec and val_a:
                key = (val_a, val_b)
                if key not in collected[curr_sec]:
                    collected[curr_sec][key] = []
                if tag not in collected[curr_sec][key]:
                    collected[curr_sec][key].append(tag)

    all_tags = [tag for _, tag in STUDY_MAPPING]

    # Flatten into rows: (kind, char, codes, applies_to)
    flat_rows = []
    for sec in sections_order:
        flat_rows.append(("section", sec, "", ""))
        for (char, codes), tags in collected[sec].items():
            if len(tags) == len(all_tags):
                applies_to = "All comparisons"
            else:
                applies_to = ", ".join(tags)
            flat_rows.append(("row", char, codes, applies_to))

    # 1. Build Excel
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Combined Table S1"

    f_title = Font(name="Arial", size=11, bold=True)
    f_header = Font(name="Arial", size=10, bold=True)
    f_section = Font(name="Arial", size=10, bold=True, italic=True)
    f_body = Font(name="Arial", size=9)

    title_text = "Supplementary Table S1. Definitions and Diagnostic/Procedural Codes for All Comparisons"
    ws_out.cell(1, 1, title_text).font = f_title
    ws_out.row_dimensions[1].height = 24.0

    headers = ["Characteristic / Element", "Codes & Definitions", "Applies to"]
    for ci, h in enumerate(headers, start=1):
        c = ws_out.cell(2, ci, h)
        c.font = f_header
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_out.row_dimensions[2].height = 20.0

    row_idx = 3
    for kind, char, codes, applies_to in flat_rows:
        if kind == "section":
            c = ws_out.cell(row_idx, 1, char)
            c.font = f_section
            ws_out.row_dimensions[row_idx].height = 18.0
        else:
            c1 = ws_out.cell(row_idx, 1, char)
            c2 = ws_out.cell(row_idx, 2, codes)
            c3 = ws_out.cell(row_idx, 3, applies_to)
            c1.font = f_body
            c2.font = f_body
            c3.font = f_body
            c1.alignment = Alignment(horizontal="left", vertical="center")
            c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c3.alignment = Alignment(horizontal="left", vertical="center")
            ws_out.row_dimensions[row_idx].height = 16.0
        row_idx += 1

    ws_out.column_dimensions['A'].width = 45.0
    ws_out.column_dimensions['B'].width = 90.0
    ws_out.column_dimensions['C'].width = 25.0

    out_xlsx = os.path.join(output_dir, "Combined_Table_S1.xlsx")
    wb_out.save(out_xlsx)
    print(f"Saved: {out_xlsx}")

    # 2. Build Word Document
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = docx.shared.Twips(PAGE_MARGIN_TWIPS)
    sec.bottom_margin = docx.shared.Twips(PAGE_MARGIN_TWIPS)
    sec.left_margin = docx.shared.Twips(PAGE_MARGIN_TWIPS)
    sec.right_margin = docx.shared.Twips(PAGE_MARGIN_TWIPS)
    sec.page_width = docx.shared.Twips(PAGE_WIDTH_TWIPS)
    sec.page_height = docx.shared.Twips(PAGE_HEIGHT_TWIPS)

    n_rows = 2 + len(flat_rows)
    col_widths = [4500, 7000, 2500]  # total 14000 twips (fits inside 15840 - 2880 = 12960)
    # Scale to page width
    available_twips = PAGE_WIDTH_TWIPS - (2 * PAGE_MARGIN_TWIPS)
    col_widths = [int(available_twips * 0.32), int(available_twips * 0.50), int(available_twips * 0.18)]
    
    t = _new_table(doc, n_rows, 3, col_widths)

    # Title row
    tc = _merge_across(t.rows[0], 0, 2)
    _write_cell(tc, title_text, align="left", valign="bottom", title_prefix=True)

    # Header row
    for ci, h in enumerate(headers):
        _write_cell(t.rows[1].cells[ci], h, bold=True, align="left", valign="bottom")

    # Body
    for i, (kind, char, codes, applies_to) in enumerate(flat_rows):
        ri = 2 + i
        row = t.rows[ri]
        if kind == "section":
            _write_cell(row.cells[0], char, bold=True, italic=True, align="left", valign="bottom")
            _write_cell(row.cells[1], "", align="left")
            _write_cell(row.cells[2], "", align="left")
        else:
            _write_cell(row.cells[0], char, align="left", valign="center")
            _write_cell(row.cells[1], codes, align="left", valign="center")
            _write_cell(row.cells[2], applies_to, align="left", valign="center")

    _apply_grouped_borders(t, n_header_rows=2, label_col=0, divider_after=None)

    out_docx = os.path.join(output_dir, "Combined_Table_S1.docx")
    doc.save(out_docx)
    print(f"Saved: {out_docx}")

if __name__ == '__main__':
    build_combined_s1("output")
