"""Defined names as references: one cell, a range, built from other names (ARCHITECTURE.md §7.3).

    uv run python -m capture.name_probe

Non-interactive. openpyxl writes formulas using defined names into column A of sheet ``S``; Excel
calculates and saves a copy, and each value and type is read from the copy. Writes
``tests/captures/names.tsv``.

``Top``, ``Bot`` and ``One`` are also column letters, so ``Top:Bot`` and ``One:One`` are whole
columns. A name defined as ``Top:Bot``, whole columns with no sheet, makes Excel refuse the file
(error -50).
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent

# Z1:Z5 hold 1 to 5, AA1:AA5 hold 10 to 50, Other!A1:A3 hold 100 to 300, Z9 is blank.
NAMES = {
    "One": "S!$Z$2",
    "Top": "S!$Z$1",
    "Bot": "S!$Z$4",
    "Block": "S!$Z$1:$Z$3",
    "Grown": "S!$Z$1:Bot",
    "Moving": "Top:OFFSET(Top,2,0)",
    "Alias": "One",
    "Elsewhere": "Other!$A$2",
    "Hole": "S!$Z$9",
    "Start": "S!$Z$1",
    "Finish": "S!$Z$4",
    "Joined": "Start:Finish",
    "EndOther": "Other!$A$3",
    "Number": "5",
    "Word": '"x"',
    "Total": "SUM(S!$Z$1:$Z$2)",
    "Broken": "#REF!",
}
FORMULAS = [
    "One",
    "One+1",
    "-One",
    "One*One",
    'One&""',
    "ROW(One)",
    "COLUMN(One)",
    "ROWS(One)",
    "COLUMNS(One)",
    "ISNUMBER(One)",
    "N(One)",
    "SUM(One)",
    "COUNT(One)",
    "OFFSET(One,1,0)",
    "OFFSET(One,0,1)",
    "SUM(OFFSET(One,0,0,2))",
    'SUMIF(OFFSET(One,1,0),">0")',
    "COUNTIF(One,2)",
    "INDEX(One,1,1)",
    'CELL("address",One)',
    "ISBLANK(Hole)",
    "Hole",
    "Hole+1",
    "ROW(Hole)",
    "Alias",
    "ROW(Alias)",
    "OFFSET(Alias,1,0)",
    "Elsewhere",
    "ROW(Elsewhere)",
    "OFFSET(Elsewhere,1,0)",
    "SUM(OFFSET(Elsewhere,-1,0,3))",
    "SUM(Top:Bot)",
    "ROWS(Top:Bot)",
    "SUM(Z1:Bot)",
    "SUM(Top:Z3)",
    "SUM(Grown)",
    "ROWS(Grown)",
    "MAX(Moving)",
    "ROWS(Moving)",
    "MIN(Top:OFFSET(Top,2,0))",
    "SUM(Block)",
    "ROWS(Block)",
    "SUM(One Block)",
    "ROW(One Block)",
    "Number",
    "Number+1",
    "ROW(Number)",
    "OFFSET(Number,1,0)",
    "Word",
    "ROW(Word)",
    "Total",
    "Total+1",
    "ROW(Total)",
    "OFFSET(Total,1,0)",
    "Broken",
    "ROW(Broken)",
    "OFFSET(Broken,1,0)",
    "SUM(Top:Broken)",
    "IF(TRUE,One)",
    "ROW(IF(TRUE,One))",
    "CHOOSE(1,One,Top)",
    "ROW(CHOOSE(1,One,Top))",
    "One:One",
    "ROW(One:One)",
    "ROW((One))",
    "OFFSET((One),1,0)",
    "SUMPRODUCT(OFFSET(One,0,0,2)*2)",
    "SUM(Start:Finish)",
    "ROWS(Start:Finish)",
    "SUM(Z2:Finish)",
    "SUM(Joined)",
    "ROWS(Joined)",
    "SUM(Other!$A$1:EndOther)",
    "ROWS(Other!$A$1:EndOther)",
    "SUM(Start:EndOther)",
]


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 6):
        ws.cell(r, 26, float(r))
        ws.cell(r, 27, float(10 * r))
    other = wb.create_sheet("Other")
    for r in range(1, 4):
        other.cell(r, 1, float(100 * r))
    for name, text in NAMES.items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    for k, formula in enumerate(FORMULAS, 1):
        ws.cell(k, 1, "=" + formula)
    wb.save(path)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_names.xlsx"
    copy = CONTAINER / "xlrecalc_names_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    try:
        result = subprocess.run(
            ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
            capture_output=True,
            text=True,
            timeout=90,
        )
    except subprocess.TimeoutExpired:
        sys.exit("Excel did not finish within 90 s. Excel is left as it is.")
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    with Book(copy) as book:
        cells = {c.row: c for c in book.cells("S") if c.column == 1}
    header = [
        "# suite: names",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/name_probe.py",
        "# method: legacy formulas written by openpyxl into column A; Excel calculates and saves a",
        "#   copy; value and type from the copy",
        f"# names: {NAMES!r}",
        "# columns: row, formula, type, value (Python repr)",
    ]
    body = []
    for k, formula in enumerate(FORMULAS, 1):
        c = cells.get(k)
        value = c.value if c else None
        body.append("\t".join([str(k), repr(formula), c.type if c else "none", repr(value)]))
        print(f"{k:3} {formula:34} {c.type if c else '':4} {value!r}")
    (CAPTURES / "names.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
