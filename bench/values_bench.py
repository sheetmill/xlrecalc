"""Time and memory of reading a workbook's values, compared across xlrecalc versions.

    uv run python bench/values_bench.py gen N
    uv run python bench/values_bench.py run [--rounds R] [--src LABEL=DIR ...] [CASE ...]

Runs in the package's own environment from the repository root. ``gen`` writes two files to
``bench/data/``, without openpyxl or Excel, so anyone can regenerate them:

- ``excel_form_N.xlsx``: the table of ``bench/gen.py`` (the same seed, so the same numbers as
  ``bench_N.xlsx``) in the form Excel saves: formulas filled down as shared formulas, every
  formula with its cached value, ``<fileVersion appName="xl">`` and Excel 16.115's own
  ``calcId`` (191029). It imitates Excel's saved form; Excel never wrote it. Opening it, Excel
  trusts every cached value (xlrecalc's opening rule answers ``"none"``).
- ``whole_columns.xlsx``: 1,000 rows (``k1``... in column A, 1... in column B) and, in column F,
  one array formula over whole columns per row (``COLUMN_FORMS``). No cached values: every
  formula is computed.

``run`` measures each case in a fresh process per version, the versions alternating, ``R``
rounds (2 by default). A version is a source directory put first on ``PYTHONPATH``
(``--src 0.6.0=<checkout of v0.6.0>/src``); ``HEAD=src`` is always measured.
Cases:

- ``load:FILE``: ``xlrecalc.load`` of ``bench/data/FILE``, then every value of ``A1:F<N>`` read
  through ``Workbook.values``;
- ``loader:FILE``: the same through a ``Loader`` fed the sheet (``Loader.workbook()``, or
  ``finish()`` where a version has no ``workbook()``);
- ``column:I``: ``whole_columns.xlsx`` loaded, then ``F<I+1>``'s value read (the time of the
  read alone).

Each child prints seconds, peak memory (``bench/rss.py``: on macOS the physical footprint, the
interpreter and imports included) and the templates compiled (the engine's own count, internal).
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import subprocess
import sys
import time
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
ROOT = HERE.parent

COLUMN_FORMS = [
    'SUMPRODUCT((A:A="k500")*B:B)',
    'SUMPRODUCT(--(A:A<>""))',
    'SUM(--(A:A<>""))',
    "MIN(B:B*1)",
    "MAX(-B:B)",
    "COUNT(B:B*1)",
    'COUNTA(A:A&"")',
    'AND(A:A<>"x")',
    'OR(A:A="k5")',
    'LOOKUP(2,1/(A:A<>""),A:A)',
    'MATCH(TRUE,A:A="",0)',
    "INDEX(B:B*1,1048576)",
    "ROWS(Z:Z*1)",
    'COLUMNS(Z:AA&"")',
    "SUMPRODUCT(--ISBLANK(A:A))",
    "SUMPRODUCT(LEN(A:A))",
    "MATCH(TRUE,ISBLANK(A:A),0)",
    "SUM(ABS(B:B))",
    'SUM(IF(A:A="k500",B:B))',
    'MAX(IF(A:A<>"",B:B))',
    "COUNT(IF(B:B>500,B:B))",
    'SUM(IF(A:A="",1))',
]
COLUMN_ROWS = 1000

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OOXML = "application/vnd.openxmlformats-officedocument"
PARTS = {
    "/xl/workbook.xml": "spreadsheetml.sheet.main+xml",
    "/xl/worksheets/sheet1.xml": "spreadsheetml.worksheet+xml",
    "/xl/styles.xml": "spreadsheetml.styles+xml",
    "/docProps/app.xml": "extended-properties+xml",
}
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels"'
    ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    + "".join(
        f'<Override PartName="{part}" ContentType="{OOXML}.{kind}"/>'
        for part, kind in PARTS.items()
    )
    + "</Types>"
)
ROOT_RELS = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
    f'<Relationship Id="rId2" Type="{REL}/extended-properties" Target="docProps/app.xml"/>'
    "</Relationships>"
)
WORKBOOK_RELS = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
    f'<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/>'
    "</Relationships>"
)
STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="{MAIN}">'
    '<fonts count="1"><font><sz val="12"/><name val="Aptos Narrow"/><family val="2"/></font>'
    "</fonts>"
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)
APP = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns='
    '"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    "<Application>Microsoft Macintosh Excel</Application></Properties>"
)


def _workbook(excel_form: bool) -> str:
    version = (
        '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="29328"/>'
        if excel_form
        else ""
    )
    calc = '<calcPr calcId="191029"/>' if excel_form else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<workbook xmlns="{MAIN}" xmlns:r="{REL}">{version}<workbookPr/>'
        f'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{calc}</workbook>'
    )


def _package(path: Path, rows: object, dimension: str, excel_form: bool) -> None:
    """Write a one-sheet package, the sheet's rows streamed from ``rows`` (strings)."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("docProps/app.xml", APP)
        z.writestr("xl/workbook.xml", _workbook(excel_form))
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", STYLES)
        with z.open("xl/worksheets/sheet1.xml", "w", force_zip64=True) as f:
            head = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<worksheet xmlns="{MAIN}"><dimension ref="{dimension}"/><sheetData>'
            )
            f.write(head.encode())
            chunk: list[str] = []
            for row in rows:  # type: ignore[attr-defined]
                chunk.append(row)
                if len(chunk) == 10_000:
                    f.write("".join(chunk).encode())
                    chunk.clear()
            f.write(("".join(chunk) + "</sheetData></worksheet>").encode())


def _excel_form_rows(n: int, seed: int = 1) -> object:
    """``bench/gen.py``'s table, as Excel writes it: shared formulas and cached values."""
    random.seed(seed)  # gen.py's own draws: one per row, rounded to 3 places
    a = [round(random.random() * 100, 3) for _ in range(n)]
    b = [x * 2 for x in a]
    c: list[float] = []
    total = 0.0
    for r in range(n):
        total = b[0] if r == 0 else total + b[r]
        c.append(total)
    d = [sum(a[r : min(r + 10, n)]) for r in range(n)]
    summary = {0: (f"SUM(B1:B{n})", sum(b)), 1: (f"SUM(C1:C{n})", sum(c))}
    summary[2] = (f"SUM(D1:D{n})", sum(d))
    last_shared_d = n - 9  # D's window is ten rows down to here, shorter below
    for r in range(n):
        x = r + 1
        if r == 0:
            fb = f'<f t="shared" ref="B1:B{n}" si="0">A1*2</f>'
            fc = "<f>B1</f>"
        else:
            fb = '<f t="shared" si="0"/>'
            fc = (
                f'<f t="shared" ref="C2:C{n}" si="1">C1+B2</f>'
                if r == 1
                else '<f t="shared" si="1"/>'
            )
        if x <= last_shared_d:
            fd = (
                f'<f t="shared" ref="D1:D{last_shared_d}" si="2">SUM(A1:A10)</f>'
                if r == 0
                else '<f t="shared" si="2"/>'
            )
        else:
            fd = f"<f>SUM(A{x}:A{n})</f>"
        cells = (
            f'<c r="A{x}"><v>{a[r]!r}</v></c><c r="B{x}">{fb}<v>{b[r]!r}</v></c>'
            f'<c r="C{x}">{fc}<v>{c[r]!r}</v></c><c r="D{x}">{fd}<v>{d[r]!r}</v></c>'
        )
        if r in summary:
            text, value = summary[r]
            cells += f'<c r="F{x}"><f>{text}</f><v>{value!r}</v></c>'
        yield f'<row r="{x}">{cells}</row>'


def _column_rows() -> object:
    for r in range(COLUMN_ROWS):
        x = r + 1
        cells = f'<c r="A{x}" t="inlineStr"><is><t>k{x}</t></is></c><c r="B{x}"><v>{x}</v></c>'
        if r < len(COLUMN_FORMS):
            cells += f'<c r="F{x}"><f t="array" ref="F{x}">{_escape(COLUMN_FORMS[r])}</f></c>'
        yield f'<row r="{x}">{cells}</row>'


def _escape(text: str) -> str:
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    )


def gen(n: int) -> None:
    """Write ``excel_form_N.xlsx`` and ``whole_columns.xlsx``."""
    DATA.mkdir(exist_ok=True)
    _package(DATA / f"excel_form_{n}.xlsx", _excel_form_rows(n), f"A1:F{n}", excel_form=True)
    _package(DATA / "whole_columns.xlsx", _column_rows(), f"A1:F{COLUMN_ROWS}", excel_form=False)
    print(f"wrote excel_form_{n}.xlsx and whole_columns.xlsx")


def _rows_of(name: str) -> int:
    match = re.search(r"_(\d+)\.xlsx$", name)
    if match is None:
        raise SystemExit(f"{name}: no row count in the name")
    return int(match.group(1))


def child(case: str) -> None:
    """Measure one case in this process and print it as JSON."""
    from rss import peak_mb

    import xlrecalc

    kind, _, arg = case.partition(":")
    out: dict[str, object] = {"case": case, "xlrecalc": xlrecalc.__version__}
    t = time.perf_counter()
    if kind == "column":
        wb = xlrecalc.load(DATA / "whole_columns.xlsx")
        t = time.perf_counter()
        out["value"] = repr(wb.value(f"Sheet1!F{int(arg) + 1}"))
    elif kind in ("load", "loader"):
        path = DATA / arg
        if kind == "load":
            wb = xlrecalc.load(path)
        else:
            import xlio

            book = xlio.open(path)
            loader = xlrecalc.Loader(book)
            for sheet in (s for s in book.sheets if s.kind == "worksheet"):
                for _ in loader.cells(sheet):
                    pass
            wb = loader.workbook() if hasattr(loader, "workbook") else loader.finish()
        out["read_s"] = time.perf_counter() - t
        grid = wb.values(f"Sheet1!A1:F{_rows_of(arg)}")
        out["values"] = [repr(grid[k][5]) for k in range(3)]
    else:
        raise SystemExit(f"unknown case {case}")
    out["seconds"] = time.perf_counter() - t
    out["peak_mb"] = peak_mb()
    engine = getattr(wb, "_engine", None)
    out["templates"] = len(engine.templates) if engine is not None else None
    print(json.dumps(out), flush=True)


def run(cases: list[str], sources: dict[str, str], rounds: int) -> None:
    """Measure each case in a child process per version, the versions alternating."""
    for case in cases:
        for k in range(rounds):
            for label, src in sources.items():
                env = dict(os.environ, PYTHONPATH=str(Path(src).resolve()))
                done = subprocess.run(
                    [sys.executable, str(Path(__file__).resolve()), "child", case],
                    env=env,
                    cwd=ROOT,
                    capture_output=True,
                    text=True,
                    check=False,
                )
                if done.returncode != 0:
                    print(f"{case} {label}: rc {done.returncode}\n{done.stderr[-2000:]}")
                    continue
                r = json.loads(done.stdout.strip().splitlines()[-1])
                print(
                    f"{case:<36} {label:<6} round {k + 1}: {r['seconds']:8.3f} s "
                    f"{r['peak_mb']:7.0f} MB  templates {r['templates']}  "
                    f"{r.get('values') or r.get('value')}",
                    flush=True,
                )


def main() -> None:
    """The command line."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="command", required=True)
    g = sub.add_parser("gen")
    g.add_argument("n", type=int)
    r = sub.add_parser("run")
    r.add_argument("--rounds", type=int, default=2)
    r.add_argument("--src", action="append", default=[], metavar="LABEL=DIR")
    r.add_argument("cases", nargs="*")
    c = sub.add_parser("child")
    c.add_argument("case")
    args = parser.parse_args()
    if args.command == "gen":
        gen(args.n)
    elif args.command == "child":
        child(args.case)
    else:
        sources = {"HEAD": str(ROOT / "src")}
        sources.update(item.split("=", 1) for item in args.src)
        cases = args.cases or [
            "load:bench_1000000.xlsx",
            "loader:bench_1000000.xlsx",
            "load:excel_form_1000000.xlsx",
            "loader:excel_form_1000000.xlsx",
        ] + [f"column:{i}" for i in range(len(COLUMN_FORMS))]
        run(cases, sources, args.rounds)


if __name__ == "__main__":
    main()
