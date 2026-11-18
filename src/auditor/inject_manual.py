"""Hand-picked 6-bug variant for quick manual testing (not the real dataset).

Reads  data/clean/fin-eval.xlsx
Writes data/injected/fin-eval-bugs.xlsx
       data/truth/fin-eval-bugs.truth.json

Formatting is never touched, so nothing visual gives a bug away.
After injecting, checks that no bugged cell accidentally equals some other cell,
since that kind of coincidence misled models in the v0 manual run.

Usage: python -m src.auditor.inject_manual
"""
import json
import shutil
from pathlib import Path

import openpyxl

from src.auditor.recalc import recalc

CLEAN = Path("data/clean/fin-eval.xlsx")
OUT = Path("data/injected/fin-eval-bugs.xlsx")
TRUTH = Path("data/truth/fin-eval-bugs.truth.json")

# (sheet, cell, bug_type, injected, note). "PASTE_VALUE" = replace formula with its current value.
BUGS = [
    ("IS", "E7", "hardcode", "PASTE_VALUE", "2028 COGS pasted as its current value"),
    ("CF", "D7", "hardcode", "PASTE_VALUE", "2027 capex pasted as its current value"),
    ("IS", "F14", "off_by_one", "=SUM(F11:F12)", "2029 total opex range leaves out R&D (row 13)"),
    ("Summary", "B12", "off_by_one", "=SUM(CF!C9:F9)", "5-yr total FCF range leaves out 2030 (column G)"),
    ("Inputs", "B12", "unit_mismatch", 25, "Tax rate entered as 25 instead of 0.25"),
    ("Summary", "B5", "unit_mismatch", "=SUM(IS!C5:G5)/1000", "Total revenue in $mm on a sheet labeled $ in thousands"),
]


# Equalities that follow from the model's structure, not coincidence.
# A total that drops the last year equals cumulative FCF through the prior year.
EXPECTED = {(("Summary", "B12"), ("CF", "F10"))}


def numeric_cells(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return {
        (ws.title, c.coordinate): c.value
        for ws in wb for row in ws.iter_rows() for c in row
        if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)
    }


def collisions(path, bug_cells):
    vals = numeric_cells(path)
    hits = []
    for key in bug_cells:
        v = vals.get(key)
        if v in (None, 0):
            continue
        for other, w in vals.items():
            if other != key and (key, other) not in EXPECTED and abs(v - w) <= 1e-9 * max(1, abs(v)):
                hits.append((key, other, v))
    return hits


def main():
    clean_vals = numeric_cells(CLEAN)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(CLEAN, OUT)
    wb = openpyxl.load_workbook(OUT)
    records = []
    for i, (sheet, cell, bug_type, injected, note) in enumerate(BUGS, start=1):
        c = wb[sheet][cell]
        original = c.value
        if injected == "PASTE_VALUE":
            injected = clean_vals[(sheet, cell)]  # full precision, like a real paste
        c.value = injected
        records.append(dict(id=i, sheet=sheet, cell=cell, bug_type=bug_type,
                            original=original, injected=injected, note=note))
    wb.save(OUT)
    recalc(OUT)

    hits = collisions(OUT, [(b["sheet"], b["cell"]) for b in records])
    for key, other, v in hits:
        print(f"WARNING: {key[0]}!{key[1]} = {other[0]}!{other[1]} = {v}")

    TRUTH.parent.mkdir(parents=True, exist_ok=True)
    TRUTH.write_text(json.dumps(dict(
        base=CLEAN.name, variant=OUT.name, bugs=records,
        notes=["Hardcodes equal their current values, so outputs do not change.",
               "Bugs 3 and 5 cascade downstream. Flags on dependent cells are ignored, not scored."],
    ), indent=2))
    print(f"wrote {OUT} and {TRUTH} ({len(records)} bugs, {len(hits)} collisions)")


if __name__ == "__main__":
    main()
