"""SUBTOTAL over rows a filter hid and rows hidden by hand (ARCHITECTURE.md §8.2).

    uv run python -m capture.hidden_probe

Non-interactive. Sheet ``Filtered`` has a filter that hides some rows; ``Hand`` has the same values
with the same rows hidden by hand; ``Both`` has a filter hiding one row and another hidden by hand.
``R`` holds ``SUBTOTAL`` of every function number over each, and other functions as controls. Excel
calculates and saves a copy, from which each value is read. Writes
``tests/captures/hidden_rows.tsv``.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl

from capture.edition import edition_line
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
RANGES = {"Filtered": "A2:A7", "Hand": "A2:A7", "Both": "A2:A7"}
NUMBERS = [*range(1, 12), *range(101, 112)]
FORMULAS = [
    *(f"SUBTOTAL({n},{s}!{r})" for s, r in RANGES.items() for n in NUMBERS),
    *(f"SUM({s}!{r})" for s in RANGES for r in [RANGES[s]]),
    *(f"COUNT({s}!{r})" for s in RANGES for r in [RANGES[s]]),
    *(f"AVERAGE({s}!{r})" for s in RANGES for r in [RANGES[s]]),
    "SUBTOTAL(9,Filtered!A2:A7,Hand!A2:A7)",
    "SUBTOTAL(9,Filtered!A1:A10)",
    "SUBTOTAL(9,Both!A2:A9)",
    "SUBTOTAL(109,Both!A2:A9)",
    "SUBTOTAL(9,Both!A8:A10)",
    "SUBTOTAL(9,Plain!A1:A3)",
    "SUBTOTAL(109,Plain!A1:A3)",
    "SUBTOTAL(9,Wide!A1:C1)",
    "SUBTOTAL(109,Wide!A1:C1)",
    "SUBTOTAL(9,Filtered!A3)",
    "SUBTOTAL(9,Filtered!A4)",
    "ROW(Filtered!A3)",
    'SUMIF(Filtered!A2:A7,">0")',
    "SUBTOTAL(9,Hand!A2:A7,Filtered!A2:A7)",
    "SUBTOTAL(9,Filtered!A2:A7,Other!A2:A7)",
    "SUBTOTAL(9,Other!A2:A7,Filtered!A2:A7)",
    "SUBTOTAL(9,Hand!A2:A7,Other!A2:A7)",
    "SUBTOTAL(109,Hand!A2:A7,Other!A2:A7)",
    "SUBTOTAL(9,Filtered!A2:A7,Filtered!A2:A7)",
    "SUBTOTAL(9,Shown!A2:A7,Filtered!A2:A7)",
    "SUBTOTAL(9,Filtered!A2:A7,Shown!A2:A7)",
    "SUBTOTAL(9,Filtered!A2:A7,Hand!A2)",
    "SUBTOTAL(9,Filtered!A2:A7,Hand!A3)",
]


def _values(ws: object) -> None:
    ws["A1"] = "v"  # type: ignore[index]
    ws["B1"] = "k"  # type: ignore[index]
    for r, (value, key) in enumerate(
        ((1.0, "x"), (2.0, "y"), (4.0, "x"), (8.0, "y"), (16.0, "x"), (32.0, "z")), 2
    ):
        ws[f"A{r}"] = value  # type: ignore[index]
        ws[f"B{r}"] = key  # type: ignore[index]


def _filter(ws: object, keep: list[str]) -> None:
    ws.auto_filter.ref = "A1:B7"  # type: ignore[attr-defined]
    ws.auto_filter.add_filter_column(1, keep)  # type: ignore[attr-defined]
    ws.sheet_properties.filterMode = True  # type: ignore[attr-defined]


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    filtered = wb.active
    filtered.title = "Filtered"
    _values(filtered)
    _filter(filtered, ["x"])
    for r in (3, 5, 7):
        filtered.row_dimensions[r].hidden = True
    hand = wb.create_sheet("Hand")
    _values(hand)
    for r in (3, 5, 7):
        hand.row_dimensions[r].hidden = True
    both = wb.create_sheet("Both")
    _values(both)
    _filter(both, ["x", "y"])
    both.row_dimensions[7].hidden = True  # the filter hid the z row
    both.row_dimensions[3].hidden = True  # by hand: a row the filter would show
    both["A9"] = 64.0
    both.row_dimensions[9].hidden = True  # by hand: outside the filter's range
    both["A10"] = 128.0
    plain = wb.create_sheet("Plain")  # a filter's range, with no column filtering: no filterMode
    for r in (1, 2, 3):
        plain[f"A{r}"] = float(r)
    plain.auto_filter.ref = "A1:A3"
    plain.row_dimensions[2].hidden = True
    wide = wb.create_sheet("Wide")  # a hidden column, which neither number leaves out
    for c, value in zip("ABC", (1.0, 2.0, 4.0), strict=True):
        wide[f"{c}1"] = value
    wide.column_dimensions["B"].hidden = True
    other = wb.create_sheet("Other")  # hidden by hand: rows that Filtered shows, and one it hides
    _values(other)
    for r in (2, 4, 7):
        other.row_dimensions[r].hidden = True
    _values(wb.create_sheet("Shown"))  # nothing hidden
    results = wb.create_sheet("R")
    for k, formula in enumerate(FORMULAS, 1):
        results.cell(k, 1, "=" + formula)
    wb.save(path)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_hidden.xlsx"
    copy = CONTAINER / "xlrecalc_hidden_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    result = osascript_or_exit(str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=90)
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    with Book(copy) as book:
        cells = {c.row: c for c in book.cells("R") if c.column == 1}
    header = [
        "# suite: hidden_rows",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/hidden_probe.py",
        "# method: rows hidden by a filter (filterMode, a filter column) or by hand; Excel",
        "#   calculates and saves a copy; value and type from the copy",
        "# columns: row, formula, type, value (Python repr)",
    ]
    body = []
    for k, formula in enumerate(FORMULAS, 1):
        c = cells.get(k)
        value = c.value if c else None
        body.append("\t".join([str(k), repr(formula), c.type if c else "none", repr(value)]))
    (CAPTURES / "hidden_rows.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
