"""Probe precision and coercion semantics, reading Excel's results from the file it saves.

    cd bench && uv run python ../capture/semantics_probe.py

openpyxl writes one formula per row in column A; Excel opens the workbook, calculates and saves
a copy (``save_copy.applescript``); the cached values and types are read from the copy's sheet
XML, which gives exact doubles and error codes. Output: ``capture/build/semantics_probe.out``.
"""

from __future__ import annotations

import html
import re
import subprocess
import zipfile
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
BUILD = HERE / "build"
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc"

FORMULAS = [
    "=0.1+0.2=0.3",
    "=0.1+0.2-0.3",
    "=(0.1+0.2)-0.3",
    "=0.1+0.2-0.3+0",
    "=1-0.9-0.1",
    "=(0.1+0.2-0.3)=0",
    "=0.1+0.2<>0.3",
    "=0.1+0.2>0.3",
    "=0^0",
    "=POWER(0,0)",
    '=0.1+0.2&""',
    '=1/3&""',
    '=10^15&""',
    '=123456789012345&""',
    '=1234567890123456&""',
    '=10^20&""',
    '=10^21&""',
    '=1E-5&""',
    '=0.0001&""',
    '=VALUE("0.12345678901234567")',
    '=VALUE("1,234")+0',
    '="$5"+1',
    '="50%"+0',
    '="(3)"+0',
    '="3/4/2025"+0',
    '="10:30 AM"+0',
    "=ROUND(2.675,2)",
    "=DATE(1900,2,29)",
    "=WEEKDAY(1)",
    '=TEXT(1234.5,"*-0")',
    '=TEXT(1234.5,"0_)")',
    '=TEXT(1234.5,"0.00e+00")',
    '=TEXT(1234567,"0")',
    "=FIXED(1234.567,1)",
    "=DOLLAR(-1234.567)",
    "=COUNTIF(B1:B3,0.3)",
]


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    BUILD.mkdir(exist_ok=True)
    src = CONTAINER / "xlrecalc_semantics_src.xlsx"
    copy = CONTAINER / "xlrecalc_semantics.xlsx"
    for p in (src, copy):
        if p.exists():
            p.unlink()
    wb = openpyxl.Workbook()
    ws = wb.active
    for row, formula in enumerate(FORMULAS, 1):
        ws.cell(row, 1, formula)
    ws["B1"] = "=0.1+0.2"  # for COUNTIF
    ws["B2"] = 0.3
    ws["B3"] = "=0.3"
    wb.save(src)
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        raise SystemExit(result.stderr)
    with zipfile.ZipFile(copy) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
        version = re.search(r"<AppVersion>([^<]*)", z.read("docProps/app.xml").decode())
    cells = {
        m.group(1): m
        for m in re.finditer(
            r'<c r="(A\d+)"([^>]*)>(?:<f>[^<]*</f>)?(?:<v[^>/]*>([^<]*)</v>|<v/>)?</c>', sheet
        )
    }
    lines = [f"app version {version.group(1) if version else '?'}"]
    for row, formula in enumerate(FORMULAS, 1):
        m = cells.get(f"A{row}")
        t = re.search(r't="(\w+)"', m.group(2)).group(1) if m and 't="' in m.group(2) else "n"
        v = html.unescape(m.group(3)) if m and m.group(3) is not None else ""
        lines.append(f"{formula}\t{t}\t{v}")
    text = "\n".join(lines)
    print(text)
    (BUILD / "semantics_probe.out").write_text(text + "\n")


if __name__ == "__main__":
    main()
