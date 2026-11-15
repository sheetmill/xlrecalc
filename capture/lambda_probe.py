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
import subprocess
import sys
import tempfile
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


def typed_rows(
    path: Path, formulas: range, names: bool, typed_list: list[str] | None = None
) -> list[str]:
    """The capture's rows: the defined names, then every cell of each formula's block."""
    body = []
    with Book(path) as book:
        for name in book.defined_names if names else ():
            body.append("\t".join(map(repr, ["name", name.name, name.text, name.sheet])))
        cells = {(c.row, c.column): c for c in book.cells("probe")}
        errors = book.value_errors
    for k in formulas:
        typed = (TYPED if typed_list is None else typed_list)[k]
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


# The LAMBDA captures, entered by pasting (``paste_formulas.applescript``), which
# enters formulas as typing them does (CAPTURING.md §2). No formula here is one Excel refuses.
LAMBDA_NAMES = {
    "ADDONE": "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)",
    "FACTR": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=1,1,_xlpm.n*FACTR(_xlpm.n-1)))",
    "CD": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,1+CD(_xlpm.n-1)))",
    "ISEV": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n=0,TRUE,ISOD(_xlpm.n-1)))",
    "ISOD": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n=0,FALSE,ISEV(_xlpm.n-1)))",
    "TWICE": "_xlfn.LAMBDA(_xlpm.f,_xlpm.x,_xlpm.f(_xlpm.f(_xlpm.x)))",
}
PASTED: dict[str, tuple[list[str], dict[str, str]]] = {
    "lambda_calls": (
        [
            "=LAMBDA(x,x+1)(1,2)",
            "=LAMBDA(x,y,x+y)(1)",
            "=LAMBDA(x,y,x+y)()",
            "=LAMBDA(x)(1)",
            "=LAMBDA(1)()",
            "=LAMBDA(x,ISOMITTED(x))()",
            "=LAMBDA(x,ISOMITTED(x))(Z9)",
            "=LAMBDA(x,y,ISBLANK(y))(1,)",
            "=LAMBDA(x,[y],y)(1)",
            "=LAMBDA(x,[y],ISOMITTED(y))(1,Z9)",
            "=LAMBDA(x,x)(1/0)",
            "=LAMBDA(x,ISERROR(x))(1/0)",
            "=SUM(LAMBDA(x,x))",
            '="a"&LAMBDA(x,x)',
            "=LAMBDA(x,x)+1",
            "=IF(TRUE,LAMBDA(x,x))",
            "=IF(TRUE,LAMBDA(x,x),0)(5)",
            "=ISREF(LAMBDA(x,x)(Z1))",
            "=ROW(LAMBDA(x,x)(Z2))",
            "=LAMBDA(x,x)(Z1:Z3)",
            "=CHOOSE(2,LAMBDA(x,x+1),LAMBDA(x,x*10))(3)",
            '=LAMBDA(x,y,x&y)("a",1)',
            "=ISERROR(LAMBDA(x,x))",
            "=TYPE(LAMBDA(x,x))",
            "=LET(f,LAMBDA(x,x*2),f)",
            "=LET(f,LAMBDA(x,x*2),MAP(Z1:Z3,f))",
            "=LAMBDA(x,LAMBDA(y,x*y))(3)",
            "=LAMBDA(X,x+1)(2)",
            "=LAMBDA(x,x*2)({1,2;3,4})",
            "=ISOMITTED(Z1)",
        ],
        {},
    ),
    "lambda_helpers": (
        [
            "=MAP(Z1:Z3,AA1:AA3,LAMBDA(a,b,a+b))",
            "=MAP(Z1:Z3,AA1:AA2,LAMBDA(a,b,a+b))",
            "=MAP({1,2},LAMBDA(a,ISREF(a)))",
            "=MAP(5,LAMBDA(a,a*2))",
            "=REDUCE(,Z1:Z3,LAMBDA(a,b,a+b))",
            "=SCAN(,Z1:Z3,LAMBDA(a,b,a+b))",
            '=REDUCE("",Z1:AA2,LAMBDA(a,b,a&b&","))',
            '=SCAN("",Z1:AA2,LAMBDA(a,b,a&b))',
            "=BYROW(Z1:AA3,LAMBDA(r,r))",
            "=BYCOL(Z1:AA3,LAMBDA(c,ROWS(c)))",
            "=BYROW(Z1:AA3,LAMBDA(r,ISREF(r)))",
            "=MAKEARRAY(0,2,LAMBDA(r,c,1))",
            "=MAKEARRAY(-1,2,LAMBDA(r,c,1))",
            "=MAKEARRAY(2.7,1.2,LAMBDA(r,c,r))",
            "=MAP(Z1:Z3,5)",
            "=MAP(Z1:Z3,LAMBDA(a,1/0))",
            "=REDUCE(0,{1,2;3,4},LAMBDA(a,b,a*10+b))",
            "=SCAN(0,{1,2;3,4},LAMBDA(a,b,a*10+b))",
            "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
            "=MAP(Z1:Z3,LAMBDA(a,ROW(a)))",
            "=BYCOL(Z1:AA3,LAMBDA(c,SUM(c)*2))",
            "=REDUCE(0,Z1:Z3,LAMBDA(a,b,Z1:Z2))",
            "=MAP(Z1:Z3,LAMBDA(a,{1,2}))",
            "=SCAN(0,5,LAMBDA(a,b,a+b))",
            "=MAP(Z9:Z10,LAMBDA(a,ISBLANK(a)))",
        ],
        {},
    ),
    "lambda_names": (
        [
            "=ADDONE(1)",
            "=FACTR(170)",
            "=FACTR(171)",
            "=ISEV(10)",
            "=ISOD(7)",
            "=TWICE(ADDONE,1)",
            "=TWICE(LAMBDA(x,x*3),2)",
            "=MAP(Z1:Z3,ADDONE)",
            "=ISREF(ADDONE)",
            "=CD(10)",
            "=CD(100)",
            "=CD(1000)",
            "=CD(2000)",
            "=CD(5000)",
            "=CD(10000)",
        ],
        LAMBDA_NAMES,
    ),
}

# Choices the engine made without a capture, and the budget's optional parameters. Not here:
# SUM(1)(2) and REDUCE of two arguments, which Excel refuses in a file (refusal_probe).
_MORE_NAMES = {
    "COPT": "_xlfn.LAMBDA(_xlpm.n,_xlop.a,IF(_xlpm.n<=0,0,1+COPT(_xlpm.n-1)))",
    "COPTG": "_xlfn.LAMBDA(_xlpm.n,_xlop.a,IF(_xlpm.n<=0,0,1+COPTG(_xlpm.n-1,1)))",
    "FIVE": "5",
    "BLETTWO": "_xlfn.LAMBDA(_xlpm.n,_xlfn.LET(_xlpm.m,_xlpm.n-1,_xlpm.k,1,"
    "IF(_xlpm.n<=0,0,_xlpm.k+BLETTWO(_xlpm.m))))",
    # A LET name only in the branch taken at the bottom, and in a branch never taken.
    "BLETBASE": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,_xlfn.LET(_xlpm.z,0,_xlpm.z),"
    "1+BLETBASE(_xlpm.n-1)))",
    "BLETNEVER": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,"
    "IF(FALSE,_xlfn.LET(_xlpm.z,1,_xlpm.z),1)+BLETNEVER(_xlpm.n-1)))",
    "VIAMAP": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,"
    "SUM(_xlfn.MAP({1},_xlfn.LAMBDA(_xlpm.x,1+VIAMAP(_xlpm.n-1))))))",
}
PASTED["lambda_more"] = (
    [
        *(f"=COPT({n})" for n in (4094, 4095, 5460, 5461)),
        *(f"=COPTG({n},1)" for n in (4094, 4095, 5460, 5461)),
        *(
            f"={fn}(LAMBDA(x,x))"
            for fn in ("ISERR", "ISNA", "ISBLANK", "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISNONTEXT")
        ),
        "=N(LAMBDA(x,x))",
        "=T(LAMBDA(x,x))",
        "=FIVE(1)",
        "=MAP({1;2;3},{10},LAMBDA(a,b,a+b))",
        "=MAP(Z1:Z3,AA1,LAMBDA(a,b,a+b))",
        "=REDUCE(0,Z1:Z3,LAMBDA(a,b,a+ISREF(b)))",
        "=SCAN(0,Z1:Z3,LAMBDA(a,b,ISREF(b)))",
        "=VIAMAP(20)",
        "=VIAMAP(1000)",
        "=LAMBDA(x,x)=1",
        "=(LAMBDA(x,x))(2)",
        # Where recursion through MAP stops (20 computes, 1,000 is #VALUE!).
        *(f"=VIAMAP({n})" for n in (30, 40, 50, 60, 63, 64, 65, 80, 100, 150, 200, 300, 500)),
        *(f"=VIAMAP({n})" for n in range(21, 30)),
        # A LET binding in a recursive body: two of them (one stops where two parameters do).
        *(f"=BLETTWO({n})" for n in (3275, 3276, 4094, 4095)),
        # MAKEARRAY with an empty size, every element shown; LET names in untaken branches.
        "=MAKEARRAY(,2,LAMBDA(r,c,r*10+c))",
        "=MAKEARRAY(2,,LAMBDA(r,c,r*10+c))",
        "=ROWS(MAKEARRAY(,2,LAMBDA(r,c,r*10+c)))",
        "=COLUMNS(MAKEARRAY(2,,LAMBDA(r,c,r*10+c)))",
        *(f"=BLETBASE({n})" for n in (4094, 4095, 5460, 5461)),
        *(f"=BLETNEVER({n})" for n in (4094, 4095, 5460, 5461)),
    ],
    _MORE_NAMES,
)

_DEPTH_NAMES = {
    "CD": LAMBDA_NAMES["CD"],
    # Not CD3: a name that is a cell's address (column CD, row 3) is not one, and Excel did not
    # save the workbook that held it.
    "CDTHREE": "_xlfn.LAMBDA(_xlpm.n,_xlpm.a,_xlpm.b,"
    "IF(_xlpm.n<=0,0,1+CDTHREE(_xlpm.n-1,_xlpm.a,_xlpm.b)))",
    "CDTWO": "_xlfn.LAMBDA(_xlpm.n,_xlpm.a,IF(_xlpm.n<=0,0,1+CDTWO(_xlpm.n-1,_xlpm.a)))",
    "CDBIG": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,1+0*SUM(_xlpm.n,1,2,3)+CDBIG(_xlpm.n-1)))",
}
_BODIES = {  # recursive bodies of one parameter, differing in what they evaluate around the call
    "BSUMTWO": "IF(_xlpm.n<=0,0,1+0*SUM(_xlpm.n,1)+BSUMTWO(_xlpm.n-1))",
    "BSUMEIGHT": "IF(_xlpm.n<=0,0,1+0*SUM(_xlpm.n,1,2,3,4,5,6,7)+BSUMEIGHT(_xlpm.n-1))",
    "BNEST": "IF(_xlpm.n<=0,0,1+ABS(ABS(ABS(BNEST(_xlpm.n-1)))))",
    "BLETX": "_xlfn.LET(_xlpm.m,_xlpm.n-1,IF(_xlpm.n<=0,0,1+BLETX(_xlpm.m)))",
    "BPLAIN": "IF(_xlpm.n<=0,0,BPLAIN(_xlpm.n-1))",
}
_DEPTH_NAMES.update({k: f"_xlfn.LAMBDA(_xlpm.n,{v})" for k, v in _BODIES.items()})
_DEPTH_ARGS = {
    "CD": "",
    "CDTHREE": ",1,2",
    "CDTWO": ",1",
    "CDBIG": "",
    **dict.fromkeys(_BODIES, ""),
}


def depth_suite(
    values: list[int], names: tuple[str, ...] = ("CD", "CDTHREE")
) -> tuple[list[str], dict[str, str]]:
    """Countdowns to each depth through each name: where Excel's recursion stops."""
    formulas = [f"={name}({n}{_DEPTH_ARGS[name]})" for name in names for n in values]
    return formulas, {name: _DEPTH_NAMES[name] for name in names}


PASTED["lambda_depth"] = depth_suite([6000, 8000])  # --depths N,N,... NAME,NAME for others


def paste(suite: str) -> None:
    """Build the suite's workbook, have Excel paste its formulas, write ``<suite>.tsv``."""
    import openpyxl
    from openpyxl.workbook.defined_name import DefinedName

    from capture.excel_watch import osascript_or_exit
    from capture.pipeline import CONTAINER, excel_version

    formulas, names = PASTED[suite]
    CONTAINER.mkdir(parents=True, exist_ok=True)
    build = CONTAINER / f"xlrecalc_{suite}.xlsx"
    copy = CONTAINER / f"xlrecalc_{suite}_saved.xlsx"
    for p in (build, copy):
        p.unlink(missing_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r, (z, aa) in enumerate(((1.0, 10.0), (2.0, 20.0), (3.0, 30.0)), 1):
        ws.cell(r, 26, z)
        ws.cell(r, 27, aa)
    for name, text in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    wb.save(build)
    lines = [""] * (5 * len(formulas))
    for k, f in enumerate(formulas):
        lines[5 * k] = f
    saved = subprocess.run(["pbpaste"], capture_output=True, check=False).stdout
    with tempfile.NamedTemporaryFile("w", suffix=".txt", encoding="utf-8", delete=False) as f:
        f.write("\n".join(lines))
        listing = f.name
    try:
        osascript_or_exit(
            str(HERE / "paste_formulas.applescript"), str(build), listing, str(copy), seconds=1800
        )
        rows = typed_rows(copy, range(len(formulas)), bool(names), formulas)
    finally:
        subprocess.run(["pbcopy"], input=saved, check=False)
        Path(listing).unlink(missing_ok=True)
        for p in (build, copy):
            p.unlink(missing_ok=True)
    header = [
        f"# suite: {suite}",
        f"# excel: {excel_version()} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        f"# generator: capture/lambda_probe.py --paste {suite} (each formula pasted as text, which",
        "#   enters it as typing does; data 1, 2, 3 in Z1:Z3 and 10, 20, 30 in AA1:AA3; formula k",
        "#   in A<5k+1>, its block A to D of rows 5k+1 to 5k+4)",
        *(["# name rows: name, name, text, sheet scope (Python repr)"] if names else []),
        "# cell rows: cell, k, typed, address, type, value, file form, f_t, f_ref, cm, rich error",
    ]
    (CAPTURES / f"{suite}.tsv").write_text("\n".join([*header, *rows]) + "\n", "utf-8")
    print(f"{suite}: {len(formulas)} formulas -> {suite}.tsv", flush=True)


def main() -> None:
    if sys.argv[1:2] == ["--typed-workbook"]:
        typed_workbook(Path(sys.argv[2]))
    elif sys.argv[1:2] == ["--typed"]:
        read_typed(Path(sys.argv[2]))
    elif sys.argv[1:2] == ["--paste"]:
        if "--depths" in sys.argv:  # a bisection of the recursion limit, a few depths a run
            at = sys.argv.index("--depths")
            values = [int(n) for n in sys.argv[at + 1].split(",")]
            names = tuple(sys.argv[at + 2].split(",")) if len(sys.argv) > at + 2 else ("CD",)
            PASTED["lambda_depth"] = depth_suite(values, names)
            del sys.argv[at : at + 3]
        for suite in sys.argv[2:]:
            paste(suite)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
