"""Which typed formulas Excel marks as dynamic arrays in a file (ARCHITECTURE.md §7.7).

    uv run python -m capture.marker_probe [value|formula_local]
    uv run python -m capture.marker_probe --typed-workbook PATH [START]
    uv run python -m capture.marker_probe --typed PATH [PATH ...]

Non-interactive. Formulas are entered into cells of a workbook openpyxl wrote (data in
Z1:Z3), through AppleScript's ``value`` of a cell or its ``formula local`` (``marker_probe
.applescript``), and Excel saves a copy; for each formula the copy shows its file form, whether
the cell carries the dynamic-array marker (``cm``) and the formula's ``t`` and ``ref``.
Writes ``tests/captures/marker_<way>.tsv``. Both ways enter legacy formulas (no marker, implicit
intersection), so AppleScript cannot enter a formula as Excel 365's formula bar does.

``--typed PATH ...`` reads instead the workbooks in which a person typed ``TYPED[k]`` into cell
``A<4k+1>`` in Excel's formula bar and saved them (``xlrecalc_marker_typed.xlsx`` for the first
``FIRST_BATCH`` formulas, ``xlrecalc_marker_typed2.xlsx`` up to ``THIRD_BATCH``,
``xlrecalc_marker_typed3.xlsx`` up to ``FOURTH_BATCH``, ``xlrecalc_marker_typed4.xlsx`` for the
rest, each built by ``typed_workbook`` from its first formula on), and writes
``tests/captures/marker_typed.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path
from typing import Any

from capture.strip import strip
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
CAPTURES = HERE.parent / "tests" / "captures"  # as capture.pipeline, which needs openpyxl
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
FIRST_BATCH = len(TYPED)
"""Formulas before this index went into the first of the typed workbooks (``KEPT``)."""
TYPED += [
    # a function that makes an array under a branch, an explicit @ and an operator
    "=IF(TRUE,SEQUENCE(1))",
    "=@SEQUENCE(3)",
    "=-SEQUENCE(1)",
    # other functions that can give an array, with a one-cell result, and one that spills
    "=TRANSPOSE(Z1)",
    "=TRANSPOSE(Z1:Z3)",
    '=LINEST("a")',
    "=LINEST(Z1:Z3)",
    "=XLOOKUP(2,Z1:Z3,Z1:Z3)",
    "=ROW(Z1)",
    "=COLUMN(Z1)",
    "=ISFORMULA(Z1)",
    "=FORMULATEXT(Z1)",
    "=INDEX(Z1:Z1,0,1)",
    "=OFFSET(Z1,0,0)",
    '=INDIRECT("Z1")',
    # the reshaping functions, each with a one-cell result
    "=CHOOSECOLS(Z1,1)",
    "=CHOOSEROWS(Z1,1)",
    "=TAKE({1,2},1,1)",
    "=DROP({1,2},,1)",
    "=EXPAND(1,1)",
    "=VSTACK(1)",
    "=HSTACK(1)",
    "=TOCOL(1)",
    "=TOROW(1)",
    "=WRAPROWS(1,1)",
    "=WRAPCOLS(1,1)",
]
THIRD_BATCH = len(TYPED)
"""Formulas from here up to ``FOURTH_BATCH`` went into the third workbook."""
TYPED += [
    # an array carried through a lifted function, an operator, a branch; consumed by SUM
    "=ABS(SEQUENCE(1))",
    "=IFERROR(SEQUENCE(1),0)",
    '=SEQUENCE(1)&""',
    "=SUM(SEQUENCE(1))+0",
    # a reference whose size is known from the text, or not
    "=INDEX(Z1:Z1,1,1)",
    "=INDEX(Z1:Z1,0)",
    "=INDEX(Z1:Z1,Z1)",
    "=OFFSET(Z1,0,0,1,1)",
    "=OFFSET(Z1,0,0,Z1,1)",
    '=INDIRECT("Z1:Z1")',
]

FOURTH_BATCH = len(TYPED)
"""Formulas from here on went into the fourth workbook."""
TYPED += [
    # array constants, written and lifted into OFFSET; a defined name holding SEQUENCE(1)
    "={1}",
    "=OFFSET(Z1,{0},0)",
    "=one_seq",
    # a function that makes an array under CHOOSE, IFS, SWITCH and N, given to INDEX, and
    # consumed by SUMPRODUCT and ROWS
    "=CHOOSE(1,SEQUENCE(1))",
    "=IFS(TRUE,SEQUENCE(1))",
    "=SWITCH(1,1,SEQUENCE(1))",
    "=N(SEQUENCE(1))",
    "=INDEX(SEQUENCE(3),1)",
    "=SUMPRODUCT(SEQUENCE(1))",
    "=ROWS(SEQUENCE(1))",
    # INDIRECT under an operator and an explicit @; a one-cell range where one value belongs
    '=INDIRECT("Z1")+0',
    '=@INDIRECT("Z1")',
    "=ABS(Z1:Z1)",
    # the functions captured since the third batch
    "=FREQUENCY(Z1:Z3,Z2)",
    "=MODE.MULT(Z1,Z1)",
    '=TEXTSPLIT("a",",")',
    "=RANDARRAY()",
    "=MUNIT(1)",
    "=MINVERSE(Z1)",
    "=MDETERM(Z1)",
]
NEXT_BATCH: list[str] = []
"""Formulas for the next typed batch, joined to ``TYPED`` once captured."""
NAMES = {"one_seq": "_xlfn.SEQUENCE(1)"}
"""Defined names the typed workbooks read (``=one_seq`` in the fourth); each is built with them."""


def typed_workbook(path: Path, start: int = 0) -> None:
    """The workbook a person types into: data in Z1:Z3, each formula's text beside its cell.

    Only formulas from ``start`` on are listed, each still at ``A<4k+1>``, so that a later batch
    goes into a workbook of its own and the one typed before is left as Excel saved it.
    """
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r in (1, 2, 3):
        ws.cell(r, 26, float(r))
    from openpyxl.workbook.defined_name import DefinedName

    for name, text in NAMES.items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    for k, formula in enumerate([*TYPED, *NEXT_BATCH][start:], start):
        ws.cell(4 * k + 1, 6, f"type into A{4 * k + 1}:  {formula}")
    wb.save(path)


KEPT = tuple(CAPTURES / f"marker_typed{n}.xlsx" for n in ("", "2", "3", "4"))
"""The typed workbooks as committed, without who saved them or where: the capture's evidence."""


def read_typed(*paths: Path) -> None:
    """Keep the typed workbooks (``KEPT``) and write ``marker_typed.tsv`` from the kept copies."""
    from capture.pipeline import excel_version

    for path, kept in zip(paths, KEPT, strict=True):
        strip(path, kept)
    header = [
        "# suite: marker_typed",
        f"# excel: {excel_version()} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/marker_probe.py --typed (a person typed each formula into the",
        "#   formula bar of Excel and saved; data 1, 2, 3 in Z1:Z3; formula k in A<4k+1>;",
        "#   the workbooks are kept as marker_typed.xlsx and marker_typed2.xlsx to 4.xlsx)",
        "# columns: typed, file form, cm, f_t, f_ref, value (Python repr)",
    ]
    body = typed_rows(*KEPT)
    (CAPTURES / "marker_typed.tsv").write_text("\n".join([*header, *body]) + "\n", "utf-8")


def typed_rows(*paths: Path) -> list[str]:
    """The capture's rows from the typed workbooks: each cell from the first that has it."""
    cells: dict[tuple[int, int], Any] = {}
    for path in paths:
        with Book(path) as book:
            for cell in book.cells("probe"):
                if cell.formula is not None:
                    cells.setdefault((cell.row, cell.column), cell)
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
    return body


def main() -> None:
    if sys.argv[1:2] == ["--typed"]:
        read_typed(*map(Path, sys.argv[2:]))
        return
    if sys.argv[1:2] == ["--typed-workbook"]:
        typed_workbook(Path(sys.argv[2]), int(sys.argv[3]) if len(sys.argv) > 3 else 0)
        return
    way = sys.argv[1] if len(sys.argv) > 1 else "value"
    from capture.pipeline import CONTAINER, excel_version

    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / f"xlrecalc_marker_{way}.xlsx"
    copy = CONTAINER / f"xlrecalc_marker_{way}_saved.xlsx"
    listing = CONTAINER / "xlrecalc_marker.txt"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    import openpyxl

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
