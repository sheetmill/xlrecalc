"""Dynamic references: OFFSET, INDIRECT, CELL, INDEX as a reference (ARCHITECTURE.md §6.3).

    uv run python -m capture.dynref_probe

Non-interactive. Sheet ``S`` holds data (Z1:Z5 = 1..5, AA1:AA5 = 10..50, AB1 = "text"),
``Sheet2!A1`` is 100, ``'My Sheet'!A1`` is 200, and the name ``Rate`` is ``S!$Z$2``. Column A
holds one legacy formula per row (``FORMULAS``); column C holds formulas whose dependencies only
evaluation finds (``PLACED``: cycles through INDIRECT and OFFSET, references in branches not
taken, a precedent found through INDIRECT). Excel opens the workbook, calculates and saves a
copy; every formula cell's value and type are read from the copy. (Excel refuses a file
holding ``OFFSET(5,0,0)``, a constant where the reference goes: it fails to open with -50.)
``CELL("filename")`` gives the file's path: the directory is written as ``<dir>``, the file's
name as ``<book>``. Writes ``tests/captures/dynrefs.tsv``.
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
FORMULAS = [
    "OFFSET(Z1,1,0)",
    "OFFSET(Z1,0,1)",
    "SUM(OFFSET(Z1,0,0,3,1))",
    "SUM(OFFSET(Z1,1,0,2,2))",
    "OFFSET(Z1,-1,0)",
    "OFFSET(Z5,0,0,-2,1)",
    "SUM(OFFSET(Z5,0,0,-2,1))",
    "SUM(OFFSET(AA5,0,0,-2,-2))",
    "OFFSET(Z1,0,0,0,1)",
    "OFFSET(Z1,1.9,0)",
    "OFFSET(Z1,-0.5,1)",
    'OFFSET(Z1,"1",0)',
    "ROWS(OFFSET(Z1,0,0,3,2))",
    "COLUMNS(OFFSET(Z1,0,0,3,2))",
    "SUM(OFFSET(Z1:Z3,1,0))",
    "SUM(OFFSET(Z1:AA2,1,0,,1))",
    "OFFSET(Z1,1048576,0)",
    "OFFSET(Sheet2!A1,0,0)",
    'INDIRECT("Z2")',
    'INDIRECT("Z"&2)',
    'INDIRECT("z2")',
    'INDIRECT("$Z$2")',
    'INDIRECT(" Z2")',
    'INDIRECT("Z2 ")',
    'SUM(INDIRECT("Z1:Z3"))',
    'INDIRECT("R2C26",FALSE)',
    'INDIRECT("R[1]C[25]",FALSE)',
    'INDIRECT("r2c26",FALSE)',
    'INDIRECT("Sheet2!A1")',
    "INDIRECT(\"'My Sheet'!A1\")",
    'INDIRECT("My Sheet!A1")',
    'INDIRECT("Rate")',
    'INDIRECT("nosuch")',
    'INDIRECT("")',
    'INDIRECT("[other.xlsx]S!A1")',
    'SUM(INDIRECT("Z:Z"))',
    "INDIRECT(5)",
    'INDIRECT("Z2",0)',
    'ROWS(INDIRECT("Z1:Z3"))',
    'INDIRECT("ZZZZ1")',
    'INDIRECT("Z2",TRUE)',
    'CELL("row",Z3)',
    'CELL("col",Z3)',
    'CELL("address",Z3)',
    'CELL("address",Sheet2!A1)',
    "CELL(\"address\",'My Sheet'!A1)",
    'CELL("address",Z1:Z3)',
    'CELL("contents",Z3)',
    'CELL("contents",AB1)',
    'CELL("contents",Y1)',
    'CELL("type",Z3)',
    'CELL("type",AB1)',
    'CELL("type",Y1)',
    'CELL("ROW",Z3)',
    'CELL("nosuch",Z3)',
    'CELL("filename",Z3)',
    'CELL("filename",Sheet2!A1)',
    'CELL("format",Z3)',
    'CELL("protect",Z3)',
    'CELL("width",Z3)',
    'CELL("prefix",AB1)',
    'CELL("color",Z3)',
    'CELL("parentheses",Z3)',
    'MID(CELL("filename",Z3),FIND("]",CELL("filename",Z3))+1,255)',
    "SUM(Z1:INDEX(Z1:Z5,3))",
    "ROWS(Z1:INDEX(Z1:Z5,3))",
    "SUM(INDEX(Z1:Z5,2):INDEX(Z1:Z5,4))",
    'CELL("address",INDEX(Z1:AA5,2,2))',
    "INDEX(Z1:Z5,3)",
    "SUM(INDEX(Z1:AA5,0,2))",
    "SUM(OFFSET(INDEX(Z1:Z5,2),0,0,2,1))",
]
# (cell, formula): formulas whose dependencies only evaluation finds
PLACED = [
    ("C1", 'INDIRECT("C1")'),
    ("C3", "OFFSET(C4,-1,0)"),
    ("C5", "IF(FALSE,C5,5)"),
    ("C7", 'INDIRECT("C8")+1'),
    ("C8", "2+3"),
    ("C10", "SUM(OFFSET(C11,0,0,2,1))"),
    ("C11", "1+1"),
    ("C12", "C11*2"),
    ("C14", "IF(TRUE,5,C14)"),
    ("C16", "CHOOSE(2,C16,7)"),
    ("C18", "IFERROR(1,C18)"),
    ("C20", "C21"),
    ("C21", "IF(C22>0,1,C20)"),
    ("C22", "1+0"),
    ("C24", 'INDIRECT("C25")*2'),
    ("C25", 'INDIRECT("C26")+1'),
    ("C26", "3+0"),
    ("C28", "IF(C29,C28,4)"),
    ("C29", "FALSE"),
]


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 6):
        ws.cell(r, 26, float(r))
        ws.cell(r, 27, float(10 * r))
    ws["AB1"] = "text"
    wb.create_sheet("Sheet2")["A1"] = 100.0
    wb.create_sheet("My Sheet")["A1"] = 200.0
    wb.defined_names["Rate"] = DefinedName("Rate", attr_text="S!$Z$2")
    for k, formula in enumerate(FORMULAS, 1):
        ws.cell(k, 1, "=" + formula)
    for cell, formula in PLACED:
        ws[cell] = "=" + formula
    wb.save(path)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_dynrefs.xlsx"
    copy = CONTAINER / "xlrecalc_dynrefs_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    folder = str(CONTAINER)
    with Book(copy) as book:
        cells = {(c.row, c.column): c for c in book.cells("S")}
    header = [
        "# suite: dynrefs",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/dynref_probe.py",
        "# method: legacy formulas written by openpyxl (column A, and placed cells in column C);",
        "#   Excel calculates and saves a copy; value and type from the copy; the file's",
        '#   directory in CELL("filename") results is written as <dir>',
        "# columns: cell, formula, type, value (Python repr)",
    ]
    body = []
    placed = [(f"A{k}", f) for k, f in enumerate(FORMULAS, 1)] + PLACED
    for where, formula in placed:
        row = int(where[1:])
        col = 1 if where[0] == "A" else 3
        c = cells.get((row, col))
        value = c.value if c else None
        if isinstance(value, str):
            value = value.replace(folder, "<dir>").replace(copy.name, "<book>")
            value = value.replace(src.name, "<book>")
        body.append("\t".join([where, repr(formula), c.type if c else "none", repr(value)]))
        print(f"{where:5} {formula:44} {c.type if c else '':4} {value!r}")
    (CAPTURES / "dynrefs.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
