"""CELL without a reference: which cell Excel takes in a recalculation on open (§6.3).

    uv run python -m capture.cell_probe

Non-interactive. Three workbooks of three sheets (S1, S2, S3) differ in the active sheet
(``workbookView activeTab``) and in each sheet's active cell (``sheetView`` ``selection
activeCell``). Each sheet holds ``CELL("address")``, ``CELL("filename")``, ``CELL("row")`` and
``CELL("col")`` without a reference in A1:A4. Excel opens each (openpyxl's files ask for a full
calculation on load), calculates and saves a copy; the results are read from the copy, the
file's directory and name written as ``<dir>`` and ``<book>``. Writes
``tests/captures/cell_noref.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl

from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
FORMULAS = ['CELL("address")', 'CELL("filename")', 'CELL("row")', 'CELL("col")']
# (active sheet index, active cell of S1, S2, S3)
VARIANTS = [(0, ("C5", "D7", "E9")), (1, ("C5", "D7", "E9")), (2, ("B2", "A1", "F3"))]


def build(path: Path, active: int, cells: tuple[str, str, str]) -> None:
    wb = openpyxl.Workbook()
    sheets = [wb.active, wb.create_sheet(), wb.create_sheet()]
    for k, (ws, cell) in enumerate(zip(sheets, cells, strict=True), 1):
        ws.title = f"S{k}"
        for r, formula in enumerate(FORMULAS, 1):
            ws.cell(r, 1, "=" + formula)
        ws.sheet_view.selection[0].activeCell = cell
        ws.sheet_view.selection[0].sqref = cell
        ws.sheet_view.tabSelected = k - 1 == active
    wb.active = active
    wb.save(path)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    version = excel_version()
    body = []
    for v, (active, cells) in enumerate(VARIANTS):
        src = CONTAINER / f"xlrecalc_cell_{v}.xlsx"
        copy = CONTAINER / f"xlrecalc_cell_{v}_saved.xlsx"
        for p in (src, copy):
            p.unlink(missing_ok=True)
        build(src, active, cells)
        result = subprocess.run(
            ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
            capture_output=True,
            text=True,
        )
        if result.returncode:
            sys.exit(f"Excel did not open {src.name}: {result.stderr.strip()}")
        with Book(copy) as book:
            for k in range(3):
                values = {c.row: c.value for c in book.cells(f"S{k + 1}") if c.column == 1}
                for r, formula in enumerate(FORMULAS, 1):
                    value = values.get(r)
                    if isinstance(value, str):
                        value = value.replace(str(CONTAINER), "<dir>")
                        value = value.replace(copy.name, "<book>").replace(src.name, "<book>")
                    body.append(
                        "\t".join(
                            [
                                str(v),
                                str(active),
                                ",".join(cells),
                                f"S{k + 1}",
                                repr(formula),
                                repr(value),
                            ]
                        )
                    )
        for p in (src, copy):
            p.unlink(missing_ok=True)
    header = [
        "# suite: cell_noref",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/cell_probe.py",
        "# method: three workbooks differing in the active sheet and the sheets' active cells;",
        "#   CELL without a reference in A1:A4 of each sheet; Excel calculates on open and saves",
        "#   a copy; values from the copy, directory and name as <dir> and <book>",
        "# columns: variant, active sheet index, active cells of S1..S3, sheet, formula, value",
    ]
    (CAPTURES / "cell_noref.tsv").write_text("\n".join([*header, *body]) + "\n", "utf-8")
    for line in body:
        print(line)


if __name__ == "__main__":
    main()
