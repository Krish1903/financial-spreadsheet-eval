"""Turn a workbook into plain text a model can audit.

One line per non-empty cell, with the real address, the formula (if any), and the value:

    === Sheet: IS ===
    A5: "Revenue"
    C5: =B5*(1+Inputs!$B$6)  ->  10700
    E7: 5038.848

A hardcode shows up as a bare number in a row of formulas. Cell addresses are explicit,
so the model never has to count rows. The file must have cached values (run recalc first).

Usage: python -m src.auditor.serialize data/injected/fin-eval-bugs.xlsx [out.txt]
"""
import sys
from pathlib import Path

import openpyxl


def fmt_value(v) -> str:
    if isinstance(v, str):
        return f'"{v}"'
    if isinstance(v, float):
        s = f"{v:.6f}".rstrip("0").rstrip(".")
        return s if s not in ("", "-0") else "0"
    return str(v)


def serialize(path: str | Path) -> str:
    formulas = openpyxl.load_workbook(path)
    values = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in formulas:
        vs = values[ws.title]
        out.append(f"=== Sheet: {ws.title} ===")
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith("="):
                    v = vs[c.coordinate].value
                    if v is None:
                        raise ValueError(f"{ws.title}!{c.coordinate} has no cached value. Run recalc first.")
                    out.append(f"{c.coordinate}: {c.value}  ->  {fmt_value(v)}")
                else:
                    out.append(f"{c.coordinate}: {fmt_value(c.value)}")
        out.append("")
    return "\n".join(out)


if __name__ == "__main__":
    text = serialize(sys.argv[1])
    if len(sys.argv) > 2:
        Path(sys.argv[2]).write_text(text)
        print(f"wrote {sys.argv[2]} ({len(text.splitlines())} lines)")
    else:
        print(text)
