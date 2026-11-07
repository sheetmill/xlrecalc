"""Ranges built with ``INDEX`` over a formula's own column: does Excel call them circular?

    uv run python -m capture.index_range_probe

Non-interactive. Formulas of the shape ``INDEX(C:C,r):INDEX(C:C,s)``, lookups and aggregates, each
in a column of its own over the numbers 1 to 5, where what they read holds the formula or does not.
Excel with iteration off leaves a circular formula uncalculated, and saves it
as the empty text or 0 (capture ``circular``), so each formula's cached value in the copy Excel
saves shows what it did. Writes ``tests/captures/index_range.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
FORMULAS = [
    # The shape: a whole column in INDEX's array, the range read does not reach the formula.
    "SUM(INDEX(C:C,1):INDEX(C:C,3))",
    "PRODUCT(INDEX(C:C,2):INDEX(C:C,4))",
    'IF(C1+C2>100,"",PRODUCT(INDEX(C:C,1):INDEX(C:C,1+C2)))',
    "SUM(INDEX($C:$C,ROW()-9):INDEX($C:$C,ROW()-8))",
    "SUM(INDEX(C1:C100,1):INDEX(C1:C100,3))",
    "INDEX(C:C,2)",
    "INDEX(C:C,2)+1",
    "ROWS(INDEX(C:C,1):INDEX(C:C,3))",
    # The range read reaches the formula itself: circular if Excel reads it.
    "SUM(INDEX(C:C,1):INDEX(C:C,ROW()))",
    "SUM(INDEX(C:C,ROW()):INDEX(C:C,ROW()+1))",
    "INDEX(C:C,ROW())",
    "ROWS(INDEX(C:C,1):INDEX(C:C,ROW()))",
    # Controls: static ranges and OFFSET.
    "SUM(C1:C3)",
    "SUM(OFFSET(C1,0,0,3))",
    "SUM(C:C)",
    "SUM(C1:C100)",
    "COUNT(C:C)",
    "ROWS(C:C)",
    "ROW(C:C)",
    "VLOOKUP(2,C:D,1,FALSE)",
    "MATCH(2,C:C,0)",
    "SUMIF(C:C,2)",
    # Lookups over their own column: a match above the formula, none, approximate, and in a table.
    "MATCH(99,C:C,0)",
    "MATCH(2,C:C,1)",
    "MATCH(99,C:C,1)",
    "MATCH(2,C1:C100,0)",
    "MATCH(5,C:C,0)",
    "VLOOKUP(99,C:D,1,FALSE)",
    "VLOOKUP(2,C:D,1,TRUE)",
    "VLOOKUP(2,C:D,2,FALSE)",
    "LOOKUP(2,C:C)",
    "HLOOKUP(1,C:D,2,FALSE)",
    "_xlfn.XLOOKUP(2,C:C,C:C)",
    "_xlfn.XMATCH(2,C:C)",
    "_xlfn.XMATCH(99,C:C)",
    "INDEX(C:C,MATCH(2,C:C,0))",
    "COUNTIF(C:C,2)",
    "SUMPRODUCT(C:C)",
]
ROW_OF_FORMULA = 10


def placed(formula: str, k: int) -> tuple[str, str]:
    """The formula's column (each its own, so a cycle in one reaches no other), and its text there.

    The formulas are written for column C and its neighbor D.
    """
    column, neighbor = get_column_letter(3 + 3 * k), get_column_letter(4 + 3 * k)
    text = re.sub(r"(\$?)\bC(?=[:\d]|\b)", lambda m: m[1] + column, formula)
    return column, re.sub(r"\bD\b", neighbor, text)


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for k, formula in enumerate(FORMULAS):
        column, text = placed(formula, k)
        for r, v in enumerate((1.0, 2.0, 3.0, 4.0, 5.0), 1):
            ws[f"{column}{r}"] = v
        ws[f"{column}{ROW_OF_FORMULA}"] = "=" + text
    wb.save(path)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_index_range.xlsx"
    copy = CONTAINER / "xlrecalc_index_range_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    result = osascript_or_exit(
        str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=180
    )
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    rows = [
        "# suite: index_range",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/index_range_probe.py",
        "# columns: formula, cell, type, value (repr), in the copy Excel saved",
    ]
    with Book(copy) as book:
        cells = {(c.row, c.column): c for c in book.cells("probe")}
    for k, formula in enumerate(FORMULAS):
        column, text = placed(formula, k)
        cell = cells[ROW_OF_FORMULA, 3 + 3 * k]
        rows.append(f"{text}\t{column}{ROW_OF_FORMULA}\t{cell.type}\t{cell.value!r}")
    (CAPTURES / "index_range.tsv").write_text("\n".join(rows) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)
    print(f"{len(FORMULAS)} rows -> index_range.tsv")


if __name__ == "__main__":
    main()
