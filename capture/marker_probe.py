"""Which typed formulas Excel marks as dynamic arrays in a file (ARCHITECTURE.md §7.7).

    uv run python -m capture.marker_probe [value|formula_local]
    uv run python -m capture.marker_probe --typed-workbook PATH
    uv run python -m capture.marker_probe --typed PATH

Non-interactive. Formulas are entered into cells of a workbook openpyxl wrote (data in
Z1:Z3), through AppleScript's ``value`` of a cell or its ``formula local`` (``marker_probe
.applescript``), and Excel saves a copy; for each formula the copy shows its file form, whether
the cell carries the dynamic-array marker (``cm``) and the formula's ``t`` and ``ref``.
Writes ``tests/captures/marker_<way>.tsv``. Both ways enter legacy formulas (no marker, implicit
intersection), so AppleScript cannot enter a formula as Excel 365's formula bar does.

``--typed PATH`` reads instead a workbook in which a person typed ``TYPED[k]`` into cell
``A<4k+1>`` in Excel's formula bar and saved it (``xlrecalc_marker_typed.xlsx``, built by
``typed_workbook``), and writes ``tests/captures/marker_typed.tsv``.
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
RS = "\x1e"
FORMULAS = [
    "=Z1:Z3*2",
    "=Z1:Z3",
    "={1,2,3}",
    "=IF(Z1>0,Z1:Z3,0)",
    "=ABS(Z1:Z3)",
    "=SUM(Z1:Z3)",
    "=Z1*2",
    "=SUM(Z1:Z3*2)",
    "=MAX(Z1:Z3*2)",
    "=@Z1:Z3",
    "=INDEX(Z1:Z3,2)",
    "=SUMPRODUCT(Z1:Z3*2)",
    "=Z:Z",
    '=Z1&"x"',
]


TYPED = [
    "=Z1:Z3*2",
    "=SUM(Z1:Z3)",
    "=@Z1:Z3",
    "=SUM(Z1:Z3*2)",
    "=Z1*2",
    "={1,2,3}",
    "=MAX(Z1:Z3*2)",
    "=INDEX(Z1:Z3,2)",
    "=ABS(Z1:Z3)",
    "=Z1:Z3",
    # functions that make arrays, each with a one-cell result, and wrapped in another function
    "=SEQUENCE(1)",
    "=SEQUENCE(1,1,5)",
    "=MMULT(Z1,Z2)",
    "=UNIQUE(Z1)",
    "=SORT(Z2)",
    "=SORTBY(Z1,Z2)",
    "=FILTER(Z1:Z3,Z1:Z3=2)",
    "=SUM(SEQUENCE(3))",
    "=ROWS(UNIQUE(Z1:Z3))",
]


def typed_workbook(path: Path) -> None:
    """The workbook a person types ``TYPED`` into: data in Z1:Z3, each formula's text beside."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r in (1, 2, 3):
        ws.cell(r, 26, float(r))
    for k, formula in enumerate(TYPED):
        ws.cell(4 * k + 1, 6, f"type into A{4 * k + 1}:  {formula}")
    wb.save(path)


def read_typed(path: Path) -> None:
    with Book(path) as book:
        cells = {(c.row, c.column): c for c in book.cells("probe")}
    header = [
        "# suite: marker_typed",
        f"# excel: {excel_version()} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/marker_probe.py --typed (a person typed each formula into the",
        "#   formula bar of Excel and saved; data 1, 2, 3 in Z1:Z3; formula k in A<4k+1>)",
        "# columns: typed, file form, cm, f_t, f_ref, value (Python repr)",
    ]
    body = []
    for k, typed in enumerate(TYPED):
        c = cells.get((4 * k + 1, 1))
        f = c.formula if c else None
        row = [
            typed,
            f.text if f else None,
            c.cm if c else None,
            f.kind if f else None,
            f.ref if f else None,
            c.value if c else None,
        ]
        body.append("\t".join(map(repr, row)))
    (CAPTURES / "marker_typed.tsv").write_text("\n".join([*header, *body]) + "\n", "utf-8")


def main() -> None:
    if sys.argv[1:2] == ["--typed"]:
        read_typed(Path(sys.argv[2]))
        return
    if sys.argv[1:2] == ["--typed-workbook"]:
        typed_workbook(Path(sys.argv[2]))
        return
    way = sys.argv[1] if len(sys.argv) > 1 else "value"
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / f"xlrecalc_marker_{way}.xlsx"
    copy = CONTAINER / f"xlrecalc_marker_{way}_saved.xlsx"
    listing = CONTAINER / "xlrecalc_marker.txt"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r in (1, 2, 3):
        ws.cell(r, 26, float(r))
    wb.save(src)
    listing.write_text(RS.join(FORMULAS), encoding="utf-8")
    version = excel_version()
    result = subprocess.run(
        [
            "osascript",
            str(HERE / "marker_probe.applescript"),
            str(src),
            str(listing),
            str(copy),
            way,
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    with Book(copy) as book:
        cells = {(c.row, c.column): c for c in book.cells("probe")}
    header = [
        f"# suite: marker_{way}",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/marker_probe.py",
        f"# method: each formula entered through AppleScript's {way} of a cell, as typed; Excel",
        "#   saves a copy; the cell's file form, cm, f t and f ref are read from the copy",
        "# columns: typed, file form, cm, f_t, f_ref, value (Python repr)",
    ]
    body = []
    for k, typed in enumerate(FORMULAS):
        c = cells.get((2 * k + 1, 1))
        f = c.formula if c else None
        row = [
            typed,
            f.text if f else None,
            c.cm if c else None,
            f.kind if f else None,
            f.ref if f else None,
            c.value if c else None,
        ]
        body.append("\t".join(map(repr, row)))
        print(row)
    (CAPTURES / f"marker_{way}.tsv").write_text("\n".join([*header, *body]) + "\n", "utf-8")
    for p in (src, copy, listing):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
