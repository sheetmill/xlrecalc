"""Which file-form prefix a function needs, by what Excel computes (ARCHITECTURE.md §5.2).

    uv run python -m capture.prefix_probe

Non-interactive. For functions the formula-bar capture could not settle (``formula_bar.py``:
the formula bar refused every call tried, or its legacy formula does not know the function),
openpyxl writes each call into a file twice, with and without the ``_xlfn.`` prefix, over the
inputs 1, 2, 3 in Z1:Z3; Excel calculates and saves a copy. The form Excel computes is the file
form; the other gives ``#NAME?``. Writes ``tests/captures/prefixes.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl

from capture.edition import edition_line
from capture.pipeline import CAPTURES, CONTAINER, excel_version, read_results, result_repr

HERE = Path(__file__).resolve().parent
L = "_xlfn.LAMBDA"
# (function, formula with {f} where the function's name goes)
CASES = [
    ("MAXIFS", '{f}(Z1:Z3,Z1:Z3,">1")'),
    ("MINIFS", '{f}(Z1:Z3,Z1:Z3,">1")'),
    ("MAP", f"SUM({{f}}(Z1:Z3,{L}(_xlpm.x,_xlpm.x*2)))"),
    ("REDUCE", f"{{f}}(0,Z1:Z3,{L}(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b))"),
    ("SCAN", f"SUM({{f}}(0,Z1:Z3,{L}(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)))"),
    ("BYROW", f"SUM({{f}}(Z1:Z3,{L}(_xlpm.r,SUM(_xlpm.r))))"),
    ("BYCOL", f"SUM({{f}}(Z1:Z3,{L}(_xlpm.c,SUM(_xlpm.c))))"),
    ("MAKEARRAY", f"SUM({{f}}(2,2,{L}(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c)))"),
    ("ISOMITTED", f"{L}(_xlpm.x,_xlpm.y,{{f}}(_xlpm.y))(1)"),
    ("LAMBDA", "{f}(_xlpm.x,_xlpm.x+1)(2)"),
    ("LET", "{f}(_xlpm.x,1,_xlpm.x+1)"),
    ("REGEXTEST", '{f}("abc","b")'),
    ("REGEXEXTRACT", '{f}("abc","b")'),
    ("REGEXREPLACE", '{f}("abc","b","x")'),
    ("SUMIFS", '{f}(Z1:Z3,Z1:Z3,">1")'),
    ("AVERAGEIFS", '{f}(Z1:Z3,Z1:Z3,">1")'),
    ("AVERAGEIF", '{f}(Z1:Z3,">1")'),
    ("LARGE", "{f}(Z1:Z3,1)"),
    ("SMALL", "{f}(Z1:Z3,1)"),
    ("PERCENTILE", "{f}(Z1:Z3,0.5)"),
    ("PERCENTILE.INC", "{f}(Z1:Z3,0.5)"),
    ("QUARTILE", "{f}(Z1:Z3,1)"),
    ("QUARTILE.INC", "{f}(Z1:Z3,1)"),
    ("MODE", "{f}(1,1,2)"),
    ("MODE.SNGL", "{f}(1,1,2)"),
    ("IFS", "{f}(TRUE,1)"),
    ("SWITCH", '{f}(1,1,"a")'),
    ("TEXTJOIN", '{f}(",",TRUE,"a","b")'),
    ("CONCAT", '{f}("a","b")'),
    ("XLOOKUP", "{f}(1,Z1:Z3,Z1:Z3)"),
    ("XMATCH", "{f}(1,Z1:Z3)"),
]


def formulas() -> list[tuple[str, str, str]]:
    out = []
    for name, template in CASES:
        for prefix in ("_xlfn.", ""):
            out.append((name, prefix, template.format(f=prefix + name)))
    return out


def main() -> None:
    rows = formulas()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_prefixes.xlsx"
    copy = CONTAINER / "xlrecalc_prefixes_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in (1, 2, 3):
        ws.cell(r, 26, float(r))
    for r, (_, _, formula) in enumerate(rows, 1):
        ws.cell(r, 1, "=" + formula)
    wb.save(src)
    version = excel_version()
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    results = read_results(copy, len(rows))
    header = [
        "# suite: prefixes",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/prefix_probe.py",
        "# method: each call written by openpyxl with and without _xlfn.; Excel calculates and",
        "#   saves a copy; the value is its cached value (#NAME? for a form it does not know)",
        "# columns: function prefix formula type value",
    ]
    body = []
    for (name, prefix, formula), res in zip(rows, results, strict=True):
        kind, text = result_repr(res)
        body.append("\t".join([name, repr(prefix), repr(formula), kind, text]))
        print(f"{name:13} {prefix!r:9} {kind} {text}")
    (CAPTURES / "prefixes.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
