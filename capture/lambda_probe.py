"""How Excel writes ``LET``, ``LAMBDA`` and its helpers into a file, typed in its formula bar.

    uv run python -m capture.lambda_probe --typed-workbook PATH
    uv run python -m capture.lambda_probe --typed PATH

AppleScript enters only legacy formulas and does not know ``LAMBDA`` calls (capture
``formula_bar``), and openpyxl's plain formulas holding ``LAMBDA`` show ``#NAME?`` in Excel
(capture ``prefixes``), so a person types each formula of ``TYPED`` into cell ``A<5k+1>`` of the
workbook ``--typed-workbook`` builds (data 1, 2, 3 in Z1:Z3 and 10, 20, 30 in AA1:AA3), after
defining the names of ``NAMES`` in Excel's Name Manager, and saves it. ``--typed`` reads every cell
of each formula's block (rows ``5k+1`` to ``5k+4``, columns A to D) with its value, the formula's
file form, ``t`` and ``ref``, ``cm`` (the dynamic-array marker) and the error a rich value holds,
and each defined name as Excel stored it. A formula Excel refuses is typed as text, after an
apostrophe, and recorded as such.

The ``LET`` formulas (``LET_COUNT``) are written to ``tests/captures/let_typed.tsv``. Office Home &
Business 2021 for Mac lacks ``LAMBDA`` and its helpers, so their rows, the defined names and the
workbook are kept in ``capture/build/`` (not committed) as a record of what that edition does,
until an edition that has them is captured; its formula bar refuses ``r`` and ``c`` as parameter
names (R1C1 shorthand), which a recapture renames.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

from capture.edition import edition_line
from capture.strip import strip
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
CAPTURES = HERE.parent / "tests" / "captures"
HELD = HERE / "build" / "lambda_typed"
"""The typed workbook, stripped, and the LAMBDA rows: local, not committed."""
LET_COUNT = 6
"""The first formulas of ``TYPED``, the LET formulas, which go into the committed capture."""

NAMES = {
    "ADDONE": "=LAMBDA(x,x+1)",
    "FACTR": "=LAMBDA(n,IF(n<=1,1,n*FACTR(n-1)))",
}
"""Names the person defines in Name Manager before typing, with what they refer to."""

TYPED = [
    # LET
    "=LET(x,1,x+1)",
    "=LET(x,Z1:Z3,x*2)",
    "=LET(x,Z1:Z3,SUM(x))",
    "=LET(x,SEQUENCE(3),SUM(x))",
    "=LET(x,2,y,x*3,x+y)",
    "=LET(x,Z1,x)",
    # LAMBDA called in place, uncalled, nested, with an omitted argument
    "=LAMBDA(x,x+1)(2)",
    "=LAMBDA(x,y,x*y)(2,3)",
    "=LAMBDA(x,x+1)",
    "=LAMBDA(x,x*2)(Z1:Z3)",
    "=LAMBDA(x,LAMBDA(y,x+y))(1)(2)",
    "=LAMBDA(x,y,IF(ISOMITTED(y),x,x+y))(1,)",
    "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1)",
    "=LAMBDA(x,x)(Z1)",
    # LET holding a lambda, and a closure over a LET name
    "=LET(f,LAMBDA(x,x*10),f(4))",
    "=LET(a,2,f,LAMBDA(x,x*a),f(3))",
    # the helpers
    "=MAP(Z1:Z3,LAMBDA(c,c*2))",
    "=MAP(Z1:Z3,LAMBDA(c,ISREF(c)))",
    "=REDUCE(0,Z1:Z3,LAMBDA(a,b,a+b))",
    "=SCAN(0,Z1:Z3,LAMBDA(a,b,a+b))",
    "=BYROW(Z1:AA3,LAMBDA(r,SUM(r)))",
    "=BYCOL(Z1:AA3,LAMBDA(c,SUM(c)))",
    "=MAKEARRAY(2,2,LAMBDA(r,c,r*c))",
    "=SUM(MAP(Z1:Z3,LAMBDA(c,c*2)))",
    "=REDUCE(0,SEQUENCE(3),LAMBDA(a,b,a+b))",
    # names defined as lambdas (NAMES)
    "=ADDONE(5)",
    "=FACTR(5)",
    "=ADDONE",
    "=ADDONE(Z1:Z3)",
]


def typed_workbook(path: Path) -> None:
    """The workbook a person types into: the data, and each formula's text beside its cell."""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r, (z, aa) in enumerate(((1.0, 10.0), (2.0, 20.0), (3.0, 30.0)), 1):
        ws.cell(r, 26, z)
        ws.cell(r, 27, aa)
    for k, formula in enumerate(TYPED):
        ws.cell(5 * k + 1, 5, f"type into A{5 * k + 1}:  {formula}")
    wb.save(path)


def typed_rows(path: Path, formulas: range, names: bool) -> list[str]:
    """The capture's rows: the defined names, then every cell of each formula's block."""
    body = []
    with Book(path) as book:
        for name in book.defined_names if names else ():
            body.append("\t".join(map(repr, ["name", name.name, name.text, name.sheet])))
        cells = {(c.row, c.column): c for c in book.cells("probe")}
        errors = book.value_errors
    for k in formulas:
        typed = TYPED[k]
        for row in range(5 * k + 1, 5 * k + 5):
            for col in range(1, 5):
                c = cells.get((row, col))
                if c is None:
                    continue
                f = c.formula
                fields = [
                    "cell",
                    k,
                    typed,
                    f"{'ABCD'[col - 1]}{row}",
                    c.type,
                    c.value,
                    f.text if f else None,
                    f.kind if f else None,
                    f.ref if f else None,
                    c.cm,
                    errors.get(c.vm) if c.vm is not None else None,
                ]
                body.append("\t".join(map(repr, fields)))
    return body


def read_typed(path: Path) -> None:
    """Write ``let_typed.tsv``; keep the workbook, stripped, and the LAMBDA rows in ``HELD``."""
    from capture.pipeline import excel_version

    HELD.mkdir(parents=True, exist_ok=True)
    kept = HELD / "lambda_typed.xlsx"
    strip(path, kept)
    for suite, target, formulas, names in (
        ("let_typed", CAPTURES, range(LET_COUNT), False),
        ("lambda_typed", HELD, range(LET_COUNT, len(TYPED)), True),
    ):
        header = [
            f"# suite: {suite}",
            f"# excel: {excel_version()} for Mac, locale en_US",
            edition_line(),
            f"# captured: {dt.date.today().isoformat()}",
            "# generator: capture/lambda_probe.py --typed (a person typed each formula into the",
            "#   formula bar of Excel and saved; data 1, 2, 3 in Z1:Z3 and 10, 20, 30 in AA1:AA3;",
            "#   formula k in A<5k+1>)",
            *(["# name rows: name, name, text, sheet scope (Python repr)"] if names else []),
            "# cell rows: cell, k, typed, address, type, value, file form, f_t, f_ref, cm, "
            "rich error",
        ]
        rows = typed_rows(kept, formulas, names)
        (target / f"{suite}.tsv").write_text("\n".join([*header, *rows]) + "\n", "utf-8")


def main() -> None:
    if sys.argv[1:2] == ["--typed-workbook"]:
        typed_workbook(Path(sys.argv[2]))
    elif sys.argv[1:2] == ["--typed"]:
        read_typed(Path(sys.argv[2]))
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
