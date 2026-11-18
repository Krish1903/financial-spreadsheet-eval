"""Recalculate a workbook with headless LibreOffice so formula cells have cached values.

openpyxl writes formulas without computed values. This round-trips the file through
LibreOffice, which computes every formula and saves the results back in place.

Usage: python -m src.auditor.recalc path/to/file.xlsx
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        if shutil.which(name):
            return name
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(mac).exists():
        return mac
    raise RuntimeError("LibreOffice not found. Mac: brew install --cask libreoffice")


def recalc(path: str | Path) -> None:
    path = Path(path).resolve()
    with tempfile.TemporaryDirectory() as out, tempfile.TemporaryDirectory() as profile:
        subprocess.run(
            [find_soffice(), f"-env:UserInstallation=file://{profile}", "--headless",
             "--calc", "--convert-to", "xlsx", "--outdir", out, str(path)],
            check=True, capture_output=True, timeout=120,
        )
        shutil.copyfile(Path(out) / path.name, path)


if __name__ == "__main__":
    for p in sys.argv[1:]:
        recalc(p)
        print(f"recalculated {p}")
