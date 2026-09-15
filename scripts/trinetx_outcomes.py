"""
Parse the Results / outcomes section of a TriNetX Compare Outcomes PDF.

For each outcome (e.g. "Mortality", "Surgery + Mortality", "Surgery") we capture,
per cohort:
  - patients in cohort, patients with outcome, Risk
  - KM survival probability at end of window
And per outcome (shared):
  - Log-Rank Test p-value
  - Hazard Ratio + 95% CI

Table 2 (per locked skill rules) uses:
  - cohort %(n): % = Risk*100 (2 dp), n = patients-with-outcome
  - event probability rate: 100 - KM survival %  (format "C1% vs C2%", 2 dp)
  - p-value: Log-Rank Test p (render p=0.000 -> <0.001)
  - HR [95% CI]: "x.xxx [low-high]" (3 dp)

Outcome-block detection is STRUCTURAL, not name-based: every outcome results block
begins with a numbered header line ("<n> <Name>") that sits immediately before a
"Risk analysis" line. This captures any outcome name -- the PLT<50 study uses
"Mortality" + "Surgery + Mortality"; the PLT<100 study uses "Mortality" + "Surgery".
"""
import re
import pdfplumber

FLOAT = r"[-+]?\d+(?:\.\d+)?"


def _full_text(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    return "\n".join((pg.extract_text() or "") for pg in pdf.pages)


def _find_outcome_starts(lines):
    """Return list of (line_index, outcome_name). An outcome header is the numbered
    line '<n> <Name>' immediately preceding a 'Risk analysis' line. We walk back over
    blank lines and a possible 'Jul 26, 2026' page-date artifact to find it."""
    starts = []
    hdr_re = re.compile(r"^\s*(\d+)\s+(.+?)\s*$")
    date_re = re.compile(r"^[A-Z][a-z]{2} \d{1,2}, \d{4}$")  # 'Jul 26, 2026'
    for i, ln in enumerate(lines):
        if ln.strip() != "Risk analysis":
            continue
        # walk backward to the nearest non-blank, non-date line
        j = i - 1
        while j >= 0 and (not lines[j].strip() or date_re.match(lines[j].strip())):
            j -= 1
        if j < 0:
            continue
        m = hdr_re.match(lines[j])
        if not m:
            continue
        num, name = m.group(1), m.group(2).strip()
        # sanity: header number is small (1..9) and name is short, non-numeric text
        if not num.isdigit() or int(num) > 9:
            continue
        if re.search(r"\d", name) or len(name) > 40:
            continue
        starts.append((j, name))
    return starts


def parse_outcomes(pdf_path):
    text = _full_text(pdf_path)
    lines = text.split("\n")
    starts = _find_outcome_starts(lines)

    outcomes = []
    for idx, (line_i, name) in enumerate(starts):
        # block spans from this header to the next header (or EOF)
        end_line = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block = "\n".join(lines[line_i:end_line])
        if "Risk analysis" not in block or "Kaplan" not in block:
            continue
        rec = {"outcome": name}

        # Risk analysis rows: "<marker> cSDH: <query name> <n_cohort> <n_outcome> <risk>"
        risk_sec = block.split("Kaplan", 1)[0]
        risk_rows = re.findall(
            r"(?m)^\s*([12])\s+cSDH:.*?\s(\d[\d,]*)\s+(\d[\d,]*)\s+(" + FLOAT + r")\s*$",
            risk_sec)

        # Kaplan-Meier survival: "<marker> cSDH: <query> <n_cohort> <n_outcome> <median|--> <surv>%"
        km_sec = block.split("Kaplan", 1)[1] if "Kaplan" in block else ""
        km_rows = re.findall(
            r"(?m)^\s*([12])\s+cSDH:.*?\s(\d[\d,]*)\s+(\d[\d,]*)\s+(--|\d[\d,]*)\s+(" + FLOAT + r")%",
            km_sec)

        # Log-Rank p  (chi2 df p)
        lr = re.search(r"Log-Rank Test\s+(" + FLOAT + r")\s+(\d+)\s+(" + FLOAT + r")", block)
        # Hazard ratio + CI. Layout: "Hazard Ratio and\n<HR> (lo, hi) chi2 df p\nProportionality"
        hr = re.search(r"Hazard Ratio and\s*\n\s*(" + FLOAT + r")\s*\(\s*(" +
                       FLOAT + r")\s*,\s*(" + FLOAT + r")\s*\)", block)
        if not hr:  # fallback: value may trail "Proportionality"
            hr = re.search(r"Proportionality\s+(" + FLOAT + r")\s*\(\s*(" + FLOAT +
                           r")\s*,\s*(" + FLOAT + r")\s*\)", block)

        # Risk difference / ratio / odds (captured for completeness/verification)
        rd = re.search(r"Risk Difference\s+(" + FLOAT + r")\s*\(\s*(" + FLOAT + r")\s*,\s*(" + FLOAT + r")\)", block)
        rr = re.search(r"Risk Ratio\s+(" + FLOAT + r")\s*\(\s*(" + FLOAT + r")\s*,\s*(" + FLOAT + r")\)", block)
        orr = re.search(r"Odds Ratio\s+(" + FLOAT + r")\s*\(\s*(" + FLOAT + r")\s*,\s*(" + FLOAT + r")\)", block)

        for (marker, n_cohort, n_out, risk) in risk_rows[:2]:
            rec[f"c{marker}_n_cohort"] = n_cohort.replace(",", "")
            rec[f"c{marker}_n_outcome"] = n_out.replace(",", "")
            rec[f"c{marker}_risk"] = risk
        for (marker, n_cohort, n_out, med, surv) in km_rows[:2]:
            rec[f"c{marker}_km_surv"] = surv
            rec[f"c{marker}_km_median"] = med
        if lr:
            rec["logrank_chi2"], rec["logrank_df"], rec["logrank_p"] = lr.group(1), lr.group(2), lr.group(3)
        if hr:
            rec["hr"], rec["hr_lo"], rec["hr_hi"] = hr.group(1), hr.group(2), hr.group(3)
        if rd:
            rec["risk_diff"] = (rd.group(1), rd.group(2), rd.group(3))
        if rr:
            rec["risk_ratio"] = (rr.group(1), rr.group(2), rr.group(3))
        if orr:
            rec["odds_ratio"] = (orr.group(1), orr.group(2), orr.group(3))
        outcomes.append(rec)
    return outcomes


if __name__ == "__main__":
    import sys, json
    outs = parse_outcomes(sys.argv[1])
    print(f"{len(outs)} outcome(s)")
    for o in outs:
        print(json.dumps(o, indent=2))
