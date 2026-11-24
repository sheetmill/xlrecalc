"""Reading a workbook sheet by sheet, values then statuses, when many formulas keep their value.

    uv run python bench/kept_bench.py gen [ROWS] [COLUMNS]
    uv run python bench/kept_bench.py run [--rounds R] [--src LABEL=DIR ...] [FILE]

Runs in the package's own environment from the repository root. ``gen`` writes
``bench/data/kept_sheets_ROWS_COLUMNS.xlsx`` (1,000 rows and 26 columns by default), without
openpyxl or Excel. Its six worksheets, each formula with a cached value:

- ``Linked1``, ``Linked2`` and ``Linked3``: a grid of ``ROWS`` x ``COLUMNS`` formulas reading a
  linked workbook (``=[1]data!B2``, through an ``externalLink`` part holding the linked values),
  which xlrecalc does not compute: they keep the file's value. ``3 x ROWS x COLUMNS`` such
  formulas in all (``Linked3`` negates what it reads).
- ``Numbers``: the same grid as numbers.
- each of those four: a row of dates over the grid (row 2, each column the one before plus row
  1's step; on the later sheets the earlier sheet's row 2), one ``SUMPRODUCT`` of each grid row
  with the defined name ``Rate``, and, under the grid, each column's ``AVERAGE`` and its
  ``STDEV`` over that average.
- ``Total``: the four grids added cell by cell, the same per-row ``SUMPRODUCT`` and column
  summaries, a row repeating ``Linked1``'s averages, the row the name ``Rate`` holds (each
  column's discount factor from its date in row 2 and the rate in ``Rates!B1``), and 150 rows
  each adding the four sheets' averages, each row its own template: a summary block of that
  size makes each step from a cell to its dependents weigh what it weighs in a workbook of
  this kind.
- ``Rates``: that rate.

Every grid sheet's ``SUMPRODUCT`` reads ``Rate`` on ``Total``, which reads every grid sheet:
reading the first sheet's values reaches every sheet.

With no ``<fileVersion>``, xlrecalc's opening rule answers ``"all"``: every formula it can
compute is computed.

``run`` reads the file as a reader showing each sheet does (sheetpeek's way): a ``Loader`` fed
one worksheet at a time, cells with their runs; after each sheet is fed, every formula cell's
value, then every formula cell's status. One fresh process per version and round, the versions
alternating; ``HEAD=src`` is always measured, others with ``--src LABEL=DIR``. Each child prints
each sheet's seconds and the peak memory after it (``bench/rss.py``), the total, and the statuses
counted by state, which must agree across versions.
"""

from __future__ import annotations

import argparse
import json
import os
import random
import statistics
import subprocess
import sys
import time
import zipfile
from collections import Counter
from pathlib import Path

from values_bench import APP, MAIN, OOXML, PKG_REL, REL, STYLES

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
ROOT = HERE.parent

KEPT = ("Linked1", "Linked2", "Linked3")
GRIDS = ("Linked1", "Linked2", "Numbers", "Linked3")
SHEETS = (*GRIDS, "Total", "Rates")
FIRST = 3  # the grid's first row; row 1 holds steps, row 2 dates
SUMMARIES = 150  # rows of summaries under Total's grid, one template each
YEARLY = 0.15  # Rates!B1, the rate Total's discount factors use


def _col(c: int) -> str:
    """The column letters of column ``c`` (1-based)."""
    text = ""
    while c:
        c, r = divmod(c - 1, 26)
        text = chr(65 + r) + text
    return text


def _cell(ref: str, value: float, formula: str | None = None) -> str:
    f = "" if formula is None else f"<f>{formula}</f>"
    return f'<c r="{ref}">{f}<v>{value!r}</v></c>'


def _sheet_xml(rows: dict[int, list[str]], dimension: str) -> str:
    body = "".join(f'<row r="{r}">{"".join(cells)}</row>' for r, cells in sorted(rows.items()))
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="{MAIN}">'
        f'<dimension ref="{dimension}"/><sheetData>{body}</sheetData></worksheet>'
    )


def _grid_sheet(
    grid: list[list[float]],
    formulas: list[list[str]] | None,
    rates: list[float],
    dates_from: str | None,
) -> tuple[dict[int, list[str]], list[float]]:
    """A grid sheet's rows; returns them and the sheet's column averages."""
    n, k = len(grid), len(grid[0])
    last = _col(k + 1)
    rows: dict[int, list[str]] = {1: [_cell(f"{_col(c + 2)}1", 30.0) for c in range(k)]}
    dates = [37772.0 + 30.0 * c for c in range(k)]
    if dates_from is None:
        rows[2] = [_cell("B2", dates[0])] + [
            _cell(f"{_col(c + 2)}2", dates[c], f"{_col(c + 1)}2+{_col(c + 2)}1")
            for c in range(1, k)
        ]
    else:
        rows[2] = [
            _cell(f"{_col(c + 2)}2", dates[c], f"{dates_from}!{_col(c + 2)}2") for c in range(k)
        ]
    for r in range(n):
        x = FIRST + r
        cells = [
            _cell(f"{_col(c + 2)}{x}", grid[r][c], formulas[r][c] if formulas else None)
            for c in range(k)
        ]
        total = sum(v * w for v, w in zip(grid[r], rates, strict=True))
        cells.append(_cell(f"{_col(k + 3)}{x}", total, f"SUMPRODUCT(B{x}:{last}{x},Rate)"))
        rows[x] = cells
    end = FIRST + n - 1
    averages, ratios = [], []
    for c in range(k):
        column = [grid[r][c] for r in range(n)]
        mean = statistics.fmean(column)
        averages.append(mean)
        ratios.append(statistics.stdev(column) / mean if mean else 0.0)
    rows[end + 2] = [
        _cell(
            f"{_col(c + 2)}{end + 2}",
            averages[c],
            f"IF(ISERR(AVERAGE({_col(c + 2)}{FIRST}:{_col(c + 2)}{end})),0,"
            f"AVERAGE({_col(c + 2)}{FIRST}:{_col(c + 2)}{end}))",
        )
        for c in range(k)
    ]
    rows[end + 3] = [
        _cell(
            f"{_col(c + 2)}{end + 3}",
            ratios[c],
            f"IF(ISERR(STDEV({_col(c + 2)}{FIRST}:{_col(c + 2)}{end})/{_col(c + 2)}{end + 2}),0,"
            f"STDEV({_col(c + 2)}{FIRST}:{_col(c + 2)}{end})/{_col(c + 2)}{end + 2})",
        )
        for c in range(k)
    ]
    return rows, averages


def gen(n: int, k: int) -> Path:
    """Write ``kept_sheets_N_K.xlsx``; returns its path."""
    rng = random.Random(1)
    rates = [1 / (1 + YEARLY) ** (30.0 * c / 365) for c in range(k)]  # Total's Rate row
    linked = [[round(rng.uniform(-500, 500), 2) for _ in range(3 * k)] for _ in range(n)]
    grids: dict[str, list[list[float]]] = {}
    formulas: dict[str, list[list[str]]] = {}
    for i, name in enumerate(KEPT):  # each reads its own third of the linked sheet
        sign = -1.0 if name == "Linked3" else 1.0
        grids[name] = [[sign * row[i * k + c] for c in range(k)] for row in linked]
        times = "*-1" if sign < 0 else ""
        formulas[name] = [
            [f"[1]data!{_col(i * k + c + 1)}{r + 2}{times}" for c in range(k)] for r in range(n)
        ]
    grids["Numbers"] = [[round(rng.uniform(-50, 50), 2) for _ in range(k)] for _ in range(n)]
    grids["Total"] = [[sum(grids[g][r][c] for g in GRIDS) for c in range(k)] for r in range(n)]
    formulas["Total"] = [
        ["+".join(f"{g}!{_col(c + 2)}{FIRST + r}" for g in GRIDS) for c in range(k)]
        for r in range(n)
    ]
    parts: dict[str, str] = {}
    before: str | None = None
    averages: dict[str, list[float]] = {}
    end = FIRST + n - 1
    for i, name in enumerate((*GRIDS, "Total")):
        rows, averages[name] = _grid_sheet(grids[name], formulas.get(name), rates, before)
        if name == "Total":
            rows[end + 5] = [
                _cell(
                    f"{_col(c + 2)}{end + 5}",
                    averages["Linked1"][c],
                    f"Linked1!{_col(c + 2)}{end + 2}",
                )
                for c in range(k)
            ]
            rows[end + 7] = [
                _cell(
                    f"{_col(c + 2)}{end + 7}", w, f"1/(1+Rates!$B$1)^(({_col(c + 2)}$2-$B$2)/365)"
                )
                for c, w in enumerate(rates)
            ]
            for j in range(SUMMARIES):  # each row its own template: it reads rows j apart
                x = end + 9 + j
                rows[x] = [
                    _cell(
                        f"{_col(c + 2)}{x}",
                        sum(averages[g][c] for g in GRIDS),
                        "+".join(f"{g}!{_col(c + 2)}{end + 2}" for g in GRIDS),
                    )
                    for c in range(k)
                ]
        last_row = end + 8 + SUMMARIES if name == "Total" else end + 3
        parts[f"xl/worksheets/sheet{i + 1}.xml"] = _sheet_xml(rows, f"B1:{_col(k + 3)}{last_row}")
        before = name
    parts["xl/worksheets/sheet6.xml"] = _sheet_xml({1: [_cell("B1", YEARLY)]}, "B1")
    parts["xl/externalLinks/externalLink1.xml"] = _external(linked)
    path = DATA / f"kept_sheets_{n}_{k}.xlsx"
    DATA.mkdir(exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types())
        z.writestr(
            "_rels/.rels",
            _rels(
                ("officeDocument", "xl/workbook.xml"), ("extended-properties", "docProps/app.xml")
            ),
        )
        z.writestr("docProps/app.xml", APP)
        z.writestr("xl/workbook.xml", _workbook(k, end + 7))
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            _rels(
                *(("worksheet", f"worksheets/sheet{i + 1}.xml") for i in range(len(SHEETS))),
                ("styles", "styles.xml"),
                ("externalLink", "externalLinks/externalLink1.xml"),
            ),
        )
        z.writestr("xl/styles.xml", STYLES)
        z.writestr(
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" '
            f'Type="{REL}/externalLinkPath" Target="data.xlsx" TargetMode="External"/>'
            "</Relationships>",
        )
        for part, xml in parts.items():
            z.writestr(part, xml)
    return path


def _rels(*targets: tuple[str, str]) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{PKG_REL}">'
        + "".join(
            f'<Relationship Id="rId{i + 1}" Type="{REL}/{kind}" Target="{target}"/>'
            for i, (kind, target) in enumerate(targets)
        )
        + "</Relationships>"
    )


def _content_types() -> str:
    parts = {
        "/xl/workbook.xml": "spreadsheetml.sheet.main+xml",
        "/xl/styles.xml": "spreadsheetml.styles+xml",
        "/docProps/app.xml": "extended-properties+xml",
        "/xl/externalLinks/externalLink1.xml": "spreadsheetml.externalLink+xml",
    }
    parts.update(
        {
            f"/xl/worksheets/sheet{i + 1}.xml": "spreadsheetml.worksheet+xml"
            for i in range(len(SHEETS))
        }
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels"'
        ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(f'<Override PartName="{p}" ContentType="{OOXML}.{t}"/>' for p, t in parts.items())
        + "</Types>"
    )


def _workbook(k: int, rate_row: int) -> str:
    sheets = "".join(
        f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
        for i, name in enumerate(SHEETS)
    )
    link = len(SHEETS) + 2
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><workbookPr/><sheets>{sheets}</sheets>'
        f'<externalReferences><externalReference r:id="rId{link}"/></externalReferences>'
        f'<definedNames><definedName name="Rate">Total!$B${rate_row}:${_col(k + 1)}${rate_row}'
        "</definedName>"
        "</definedNames></workbook>"
    )


def _external(linked: list[list[float]]) -> str:
    rows = "".join(
        f'<row r="{r + 2}">'
        + "".join(f'<cell r="{_col(c + 1)}{r + 2}"><v>{v!r}</v></cell>' for c, v in enumerate(row))
        + "</row>"
        for r, row in enumerate(linked)
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<externalLink xmlns="{MAIN}" xmlns:r="{REL}"><externalBook r:id="rId1">'
        '<sheetNames><sheetName val="data"/></sheetNames>'
        f'<sheetDataSet><sheetData sheetId="0">{rows}</sheetData></sheetDataSet>'
        "</externalBook></externalLink>"
    )


def child(path: Path) -> None:
    """Read ``path`` sheet by sheet in this process and print the figures as JSON."""
    import xlio
    from rss import peak_mb

    import xlrecalc

    out: dict[str, object] = {"xlrecalc": xlrecalc.__version__, "sheets": {}, "peaks": {}}
    states: Counter[str] = Counter()
    start = time.perf_counter()
    with xlio.open(path) as book:
        loader = xlrecalc.Loader(book, recalc_on_load="excel")
        workbook = None
        for sheet in (s for s in book.sheets if s.kind == "worksheet"):
            t = time.perf_counter()
            refs = [
                xlrecalc.CellRef(sheet.name, cell.row, cell.column)
                for cell in loader.cells(sheet, rich=True)
                if cell.formula is not None
            ]
            if workbook is None:
                out["opening"] = loader.opening
                workbook = loader.workbook()
            for ref in refs:
                workbook.value(ref)
            for ref in refs:
                states[workbook.status(ref).state] += 1
            out["sheets"][sheet.name] = time.perf_counter() - t  # type: ignore[index]
            out["peaks"][sheet.name] = peak_mb()  # type: ignore[index]
    out["seconds"] = time.perf_counter() - start
    out["peak_mb"] = peak_mb()
    out["states"] = dict(sorted(states.items()))
    print(json.dumps(out), flush=True)


def run(path: Path, sources: dict[str, str], rounds: int) -> None:
    """Measure ``path`` in a child process per version and round, the versions alternating."""
    for k in range(rounds):
        for label, src in sources.items():
            env = dict(os.environ, PYTHONPATH=str(Path(src).resolve()))
            done = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "child", str(path)],
                env=env,
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )
            if done.returncode != 0:
                print(f"{label}: rc {done.returncode}\n{done.stderr[-2000:]}")
                continue
            r = json.loads(done.stdout.strip().splitlines()[-1])
            sheets = "  ".join(f"{name} {s:.2f}" for name, s in r["sheets"].items())
            print(
                f"{label:<8} round {k + 1}: {r['seconds']:7.2f} s {r['peak_mb']:6.0f} MB  "
                f"[{sheets}]  {r['opening']} {r['states']}",
                flush=True,
            )


def main() -> None:
    """The command line."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="command", required=True)
    g = sub.add_parser("gen")
    g.add_argument("rows", type=int, nargs="?", default=1000)
    g.add_argument("columns", type=int, nargs="?", default=26)
    r = sub.add_parser("run")
    r.add_argument("--rounds", type=int, default=2)
    r.add_argument("--src", action="append", default=[], metavar="LABEL=DIR")
    r.add_argument("file", nargs="?", default=str(DATA / "kept_sheets_1000_26.xlsx"))
    c = sub.add_parser("child")
    c.add_argument("file")
    args = parser.parse_args()
    if args.command == "gen":
        print(f"wrote {gen(args.rows, args.columns)}")
    elif args.command == "child":
        child(Path(args.file))
    else:
        sources = {"HEAD": str(ROOT / "src")}
        sources.update(item.split("=", 1) for item in args.src)
        run(Path(args.file), sources, args.rounds)


if __name__ == "__main__":
    main()
