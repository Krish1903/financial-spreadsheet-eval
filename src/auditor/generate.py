"""Build the clean base model: data/clean/fin-eval.xlsx

Simple 5-year operating model (Inputs, IS, CF, Summary). Rules that make it eval-friendly:
- Every typed-in number lives on Inputs. Every other sheet is formulas only.
- Each projection row uses the same formula across 2026E-2030E.
- Each sheet states its unit: ($ in thousands).

Usage: python -m src.auditor.generate
"""
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from src.auditor.recalc import recalc

OUT = Path("data/clean/fin-eval.xlsx")

# label, value, number format key, note
# v1 inputs: changed from v0 so no subtotal accidentally equals another cell
# (v0 had S&M+G&A = 25% and growth 8%, so 2029 S&M+G&A matched 2028 EBITDA exactly)
INPUTS = [
    ("rev0", "Base year revenue (2025A)", 10000, "cur", "Illustrative, $000s"),
    ("g", "Revenue growth %", 0.07, "pct", "Applied each projection year"),
    ("cogs", "COGS % of revenue", 0.42, "pct", ""),
    ("sm", "S&M % of revenue", 0.14, "pct", ""),
    ("ga", "G&A % of revenue", 0.09, "pct", ""),
    ("rd", "R&D % of revenue", 0.07, "pct", ""),
    ("da", "D&A % of revenue", 0.03, "pct", ""),
    ("tax", "Tax rate", 0.25, "pct", "Applied to positive EBIT only"),
    ("capex", "Capex % of revenue", 0.045, "pct", ""),
    ("nwc", "NWC % of change in revenue", 0.10, "pct", ""),
]

ARIAL = dict(name="Arial", size=10)
BLUE, GREEN = "0000FF", "008000"
FMT = {"cur": "#,##0;(#,##0);-", "pct": "0.0%;(0.0%);-"}
YEARS = ["2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
COLS, PROJ = "BCDEFG", "CDEFG"
THIN = Side(style="thin")


def font(color=None, bold=False, italic=False):
    return Font(**ARIAL, color=color, bold=bold, italic=italic)


def header(ws, title):
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws["A2"] = "($ in thousands)"
    ws["A2"].font = font("666666", italic=True)


def year_header(ws):
    for c, y in zip(COLS, YEARS):
        cell = ws[f"{c}4"]
        cell.value, cell.font = y, font(bold=True)
        cell.alignment, cell.border = Alignment(horizontal="right"), Border(bottom=THIN)
        ws.column_dimensions[c].width = 11
    ws.column_dimensions["A"].width = 26
    ws.freeze_panes = "B5"


def line(ws, r, label, fn, fmt="cur", color=None, bold=False, rule=False):
    """Write one projection row. fn(col, prev_col) returns the formula body."""
    ws[f"A{r}"] = label
    ws[f"A{r}"].font = font(bold=bold)
    for c in PROJ:
        p = chr(ord(c) - 1)
        cell = ws[f"{c}{r}"]
        cell.value, cell.number_format = "=" + fn(c, p), FMT[fmt]
        cell.font = font(color, bold)
        if rule:
            cell.border = Border(top=THIN)
    if rule:
        ws[f"A{r}"].border = Border(top=THIN)


def build(out: Path = OUT) -> None:
    wb = openpyxl.Workbook()

    # Inputs
    inp = wb.active
    inp.title = "Inputs"
    header(inp, "Inputs")
    for c, v in zip("ABC", ["Assumption", "Value", "Note"]):
        inp[f"{c}4"], inp[f"{c}4"].font, inp[f"{c}4"].border = v, font(bold=True), Border(bottom=THIN)
    ref = {}
    for r, (key, label, val, fmt, note) in enumerate(INPUTS, start=5):
        inp[f"A{r}"], inp[f"A{r}"].font = label, font()
        cell = inp[f"B{r}"]
        cell.value, cell.number_format, cell.font = val, FMT[fmt], font(BLUE)
        cell.fill = PatternFill("solid", fgColor="FFFF00")
        inp[f"C{r}"], inp[f"C{r}"].font = note, font("666666", italic=True)
        ref[key] = f"Inputs!$B${r}"
    inp["A17"] = "Legend: blue/yellow = input (only place hardcodes live), black = formula, green = link from another sheet"
    inp["A18"] = "All values are illustrative placeholders, not from a real company."
    for a in ("A17", "A18"):
        inp[a].font = font("666666", italic=True)
    for c, w in zip("ABC", (30, 12, 32)):
        inp.column_dimensions[c].width = w

    # Income statement
    s = wb.create_sheet("IS")
    header(s, "Income Statement")
    year_header(s)
    s["B5"], s["B5"].font, s["B5"].number_format = "=" + ref["rev0"], font(GREEN), FMT["cur"]
    line(s, 5, "Revenue", lambda c, p: f"{p}5*(1+{ref['g']})", bold=True)
    line(s, 6, "  Growth %", lambda c, p: f"IF({p}5=0,0,{c}5/{p}5-1)", "pct")
    line(s, 7, "COGS", lambda c, p: f"{c}5*{ref['cogs']}")
    line(s, 8, "Gross Profit", lambda c, p: f"{c}5-{c}7", bold=True, rule=True)
    line(s, 9, "  Gross margin %", lambda c, p: f"IF({c}5=0,0,{c}8/{c}5)", "pct")
    line(s, 11, "S&M", lambda c, p: f"{c}5*{ref['sm']}")
    line(s, 12, "G&A", lambda c, p: f"{c}5*{ref['ga']}")
    line(s, 13, "R&D", lambda c, p: f"{c}5*{ref['rd']}")
    line(s, 14, "Total Opex", lambda c, p: f"SUM({c}11:{c}13)", bold=True, rule=True)
    line(s, 16, "EBITDA", lambda c, p: f"{c}8-{c}14", bold=True)
    line(s, 17, "  EBITDA margin %", lambda c, p: f"IF({c}5=0,0,{c}16/{c}5)", "pct")
    line(s, 18, "D&A", lambda c, p: f"{c}5*{ref['da']}")
    line(s, 19, "EBIT", lambda c, p: f"{c}16-{c}18", bold=True, rule=True)
    line(s, 20, "Taxes", lambda c, p: f"MAX(0,{c}19*{ref['tax']})")
    line(s, 21, "Net Income", lambda c, p: f"{c}19-{c}20", bold=True, rule=True)

    # Cash flow
    cf = wb.create_sheet("CF")
    header(cf, "Cash Flow")
    year_header(cf)
    line(cf, 5, "Net Income", lambda c, p: f"IS!{c}21", color=GREEN)
    line(cf, 6, "(+) D&A", lambda c, p: f"IS!{c}18", color=GREEN)
    line(cf, 7, "(-) Capex", lambda c, p: f"-IS!{c}5*{ref['capex']}")
    line(cf, 8, "(-) Change in NWC", lambda c, p: f"-(IS!{c}5-IS!{p}5)*{ref['nwc']}")
    line(cf, 9, "Free Cash Flow", lambda c, p: f"SUM({c}5:{c}8)", bold=True, rule=True)
    line(cf, 10, "Cumulative FCF", lambda c, p: f"SUM($C$9:{c}9)")
    line(cf, 11, "  FCF margin %", lambda c, p: f"IF(IS!{c}5=0,0,{c}9/IS!{c}5)", "pct")

    # Summary
    sm = wb.create_sheet("Summary")
    header(sm, "5-Year Summary (2026E-2030E)")
    for c, v in zip("AB", ["Metric", "Value"]):
        sm[f"{c}4"], sm[f"{c}4"].font, sm[f"{c}4"].border = v, font(bold=True), Border(bottom=THIN)
    rows = [
        ("Total revenue", "SUM(IS!C5:G5)", "cur"),
        ("Revenue CAGR (2025A-2030E)", "IF(IS!B5=0,0,(IS!G5/IS!B5)^(1/COUNT(IS!C5:G5))-1)", "pct"),
        ("Average growth %", "AVERAGE(IS!C6:G6)", "pct"),
        ("Total gross profit", "SUM(IS!C8:G8)", "cur"),
        ("Total EBITDA", "SUM(IS!C16:G16)", "cur"),
        ("Average EBITDA margin %", "AVERAGE(IS!C17:G17)", "pct"),
        ("Total net income", "SUM(IS!C21:G21)", "cur"),
        ("Total free cash flow", "SUM(CF!C9:G9)", "cur"),
        ("Average FCF margin %", "AVERAGE(CF!C11:G11)", "pct"),
        ("FCF conversion (FCF / net income)", "IF(B11=0,0,B12/B11)", "pct"),
    ]
    for r, (label, f, fmt) in enumerate(rows, start=5):
        sm[f"A{r}"], sm[f"A{r}"].font = label, font()
        sm[f"B{r}"], sm[f"B{r}"].number_format, sm[f"B{r}"].font = "=" + f, FMT[fmt], font()
    sm.column_dimensions["A"].width, sm.column_dimensions["B"].width = 36, 14

    for ws in wb:
        ws.sheet_view.showGridLines = False
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    recalc(out)


if __name__ == "__main__":
    build()
    print(f"wrote {OUT}")
