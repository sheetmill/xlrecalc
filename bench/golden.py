"""Every value the engine computes and every file it writes, for byte-for-byte comparison.

    uv run python bench/golden.py [--large] [OUT]

Runs in the package's own environment, from the repository root. Evaluates:

- every capture suite of ``tests/test_captures.py`` laid out as that test lays it out, and the
  captures that ``tests/test_arrays.py``, ``tests/test_array_refs.py``, ``tests/test_dynrefs.py``
  and ``tests/test_circular.py`` evaluate as workbooks;
- the bench table (``bench/data/bench_<N>.pkl``, and ``bench_<N>.xlsx`` through the package's
  reader) at 1,000, 10,000 and 100,000 rows, after a full recalculation, after changing A1 and
  after changing the last row's A, as ``bench/engine_bench.py`` does; with ``--large``, at
  1,000,000 rows too.

Every cell value goes on one line with its exact type and bits (``float.hex`` for numbers),
cells whose own formula draws a random number or reads the clock excepted, together with each
engine's report. The bench workbooks at 100,000 rows and more contribute one digest per column
and stage instead of their lines. Prints the SHA-256 of the whole output; two runs that print
the same digest computed every value identically. Writes OUT (default
``capture/build/golden.txt``, or ``golden_large.txt`` with ``--large``) for a diff when they do
not. An optimization must leave the digest unchanged, including for rows a test accepts within
a declared tolerance.

A second digest covers the files the writer produces: ``recalculate`` on the bench workbooks
(xlsx, the same sizes) and on a workbook per capture suite of ``capture/suites.py`` (built as
the capture pipeline builds it, rows drawing a random number or reading the clock or the file's
path left out); ``recalculate`` again on the written file, where nothing changes; and edits of
every kind through ``load`` and ``Workbook.save``. Each part of each written file goes on one
line with its SHA-256 (a file's new parts carry the time of writing in the zip, and openpyxl
stamps ``docProps/core.xml`` of the workbooks it builds, so the parts are compared, not the
files). The lines go to ``golden_files.txt`` (``golden_files_large.txt``) beside OUT.
"""

from __future__ import annotations

import gc
import hashlib
import io
import pickle
import re
import shutil
import sys
import zipfile
from collections.abc import Iterator
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tests import test_array_refs, test_arrays, test_circular, test_dynrefs  # noqa: E402
from tests.captures import load as load_suite  # noqa: E402
from tests.test_captures import SUITES, run  # noqa: E402

from xlrecalc._engine import Engine  # noqa: E402
from xlrecalc._load import load  # noqa: E402
from xlrecalc._store import COLS  # noqa: E402
from xlrecalc._values import ErrorValue  # noqa: E402
from xlrecalc._workbook import load as load_book  # noqa: E402

DATA = ROOT / "bench" / "data"
_UNSTABLE = frozenset({"RAND", "RANDBETWEEN", "NOW", "_xlfn.RANDARRAY", "_XLFN.RANDARRAY"})


def encode(value: object) -> str:
    """A value as exact, typed text."""
    if isinstance(value, bool):
        return f"b:{value}"
    if isinstance(value, float):
        return f"n:{value.hex()}"
    if isinstance(value, ErrorValue):
        return f"e:{value.code}"
    if isinstance(value, str):
        return f"s:{value!r}"
    return f"{type(value).__name__}:{value!r}"


def cells(engine: Engine, only_formulas: bool = False) -> Iterator[str]:
    """Every stored value, sheet by sheet in key order, one line each."""
    for s, sheet in enumerate(engine.sheets):
        keys = sheet.templates if only_formulas else sheet.values.keys() | sheet.templates.keys()
        for key in sorted(keys):
            row, col = divmod(key, COLS)
            tid = sheet.templates.get(key)
            if tid is not None and engine.templates[tid].calls & _UNSTABLE:
                yield f"{s}!{row},{col} <unstable>"
                continue
            yield f"{s}!{row},{col} {encode(engine.get(s, row, col))}"


def report(engine: Engine) -> Iterator[str]:
    """The engine's report, one line per kind."""
    r = engine.report
    yield f"circular {sorted(map(str, r.circular))}"
    yield f"uncomputed {sorted(map(str, r.uncomputed))}"
    yield f"from_kept {sorted(map(str, r.from_kept))}"
    yield f"refused {sorted((str(k), v) for k, v in r.refused.items())}"
    yield f"exceptions {sorted((str(k), v) for k, v in r.exceptions.items())}"
    yield f"unsupported {sorted(r.unsupported.items())}"


def captures() -> Iterator[str]:
    """Every capture suite and capture workbook, evaluated."""
    for name in SUITES:
        engine, _ = run(load_suite(name))
        yield f"# suite {name}"
        yield from cells(engine)
        yield from report(engine)
    workbooks = {
        "arrays": lambda: load(io.BytesIO(test_arrays.workbook())),
        "array_refs": lambda: load(io.BytesIO(test_array_refs.workbook())),
        "dynrefs": lambda: _with_path(load(io.BytesIO(test_dynrefs.workbook()))),
        "circular": _circular,
    }
    for name, make in workbooks.items():
        engine = make()
        engine.recalculate()
        yield f"# workbook {name}"
        yield from cells(engine)
        yield from report(engine)


def _with_path(engine: Engine) -> Engine:
    engine.path = "<dir>/<book>"
    return engine


def _circular() -> Engine:
    import ast

    sheet: dict[str, object] = {}
    for number, _tag, formula, constant, *_ in test_circular.CIRCULAR:
        sheet[f"A{number}"] = "=" + formula
        if constant != "None":
            sheet[f"B{number}"] = ast.literal_eval(constant)
    return Engine.from_cells({"probe": sheet})


def _table(n: int) -> Engine:
    with (DATA / f"bench_{n}.pkl").open("rb") as f:
        table = pickle.load(f)
    engine = Engine()
    s = engine.add_sheet("Sheet1")
    sheet = engine.sheets[s]
    for r, row in enumerate(table):
        for c, v in enumerate(row):
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                engine.set_formula(s, r, c, v[1:])
            else:
                sheet.set(r, c, float(v))
    return engine


def _stage(engine: Engine, label: str, digest_columns: bool) -> Iterator[str]:
    yield f"## {label}"
    lines = cells(engine, only_formulas=True)
    if not digest_columns:
        yield from lines
    else:
        by_column: dict[int, hashlib._Hash] = {}
        counts: dict[int, int] = {}
        for line in lines:
            col = int(line.split(",", 1)[1].split(" ", 1)[0])
            by_column.setdefault(col, hashlib.sha256()).update(line.encode() + b"\n")
            counts[col] = counts.get(col, 0) + 1
        for col in sorted(by_column):
            yield f"column {col}: {counts[col]} cells, sha256 {by_column[col].hexdigest()}"
    yield from report(engine)


def bench(n: int, source: str) -> Iterator[str]:
    """The bench table of n rows from the pickle or the xlsx, at three stages."""
    path = DATA / f"bench_{n}.{'pkl' if source == 'pickle' else 'xlsx'}"
    if not path.exists():
        sys.exit(f"missing {path}: cd bench && uv run python gen.py {n}")
    digest_columns = n >= 100_000
    engine = _table(n) if source == "pickle" else load(path)
    yield f"# bench {n} {source}"
    engine.recalculate(full=True)
    yield from _stage(engine, "full", digest_columns)
    engine.set_value(0, 0, 0, 1000.0)
    engine.recalculate()
    yield from _stage(engine, "A1 = 1000", digest_columns)
    engine.set_value(0, n - 1, 0, 5.0)
    engine.recalculate()
    yield from _stage(engine, "last A = 5", digest_columns)


_VOLATILE_CASE = re.compile(r"\b(RAND|RANDBETWEEN|RANDARRAY|NOW|TODAY|CELL)\(", re.IGNORECASE)
FILES = ROOT / "capture" / "build" / "golden_files"  # fixed: CELL-free, but paths stay stable


def parts(path: Path, skip: frozenset[str] = frozenset()) -> Iterator[str]:
    """Each part of a written package: its name, compression and the SHA-256 of its bytes."""
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            if info.filename in skip:
                yield f"{info.filename} <skipped>"
                continue
            digest = hashlib.sha256(z.read(info)).hexdigest()
            yield f"{info.filename} {info.compress_type} {info.file_size} {digest}"


def _edit(path: Path, out: Path, last_row: int) -> None:
    """Edits of every kind through the public API, saved to ``out``."""
    book = load_book(path)
    book.set("A1", 1000)  # a constant changed
    book.set("A3", " text ")  # a number made text with outer whitespace
    book.set("A4", True)  # a logical
    book.set("A5", ErrorValue.NA)  # an error
    book.set("B2", None)  # a formula cleared
    book.set("C6", 7)  # a formula made a constant
    book.set_formula("E2", '=IFNA(B3,0)&"!"')  # a new cell in a row
    book.set_formula(f"A{last_row + 2}", "=SUM(A1:A10)")  # a new row at the end
    book.set("Z1", 1.5)
    book.set_formula("Z2", "=Z1*2")
    book.save(out)


def written(source: Path, label: str, skip: frozenset[str] = frozenset()) -> Iterator[str]:
    """recalculate, recalculate again (nothing changes) and edits, each written file's parts."""
    from xlrecalc import recalculate

    folder = FILES / label
    shutil.rmtree(folder, ignore_errors=True)
    folder.mkdir(parents=True)
    first, again, edited = (
        folder / "recalculated.xlsx",
        folder / "again.xlsx",
        folder / "edited.xlsx",
    )
    report = recalculate(source, first)
    yield f"# written {label}: recalculate"
    yield f"report circular={len(report.circular)} uncomputed={len(report.uncomputed)}"
    yield from parts(first, skip)
    recalculate(first, again)
    yield f"# written {label}: recalculate again"
    yield from parts(again, skip)
    with zipfile.ZipFile(first) as z:
        rows = len(re.findall(rb"<row ", z.read("xl/worksheets/sheet1.xml")))
    _edit(first, edited, rows)
    yield f"# written {label}: edited"
    yield from parts(edited, skip)
    shutil.rmtree(folder)


def written_files(sizes: list[int]) -> Iterator[str]:
    """The writer's output for the bench workbooks and a workbook per capture suite."""
    from capture import suites
    from capture.pipeline import build

    for n in sizes:
        yield from written(DATA / f"bench_{n}.xlsx", f"bench_{n}")
    for name, suite in suites.SUITES.items():
        cases = [c for c in suite.cases() if not _VOLATILE_CASE.search(c.formula)]
        FILES.mkdir(parents=True, exist_ok=True)
        source = FILES / f"suite_{name}.xlsx"
        build(source, cases, suite.date1904, suite.fixed)
        yield from written(source, f"suite_{name}", frozenset({"docProps/core.xml"}))
        source.unlink()


def _write(lines: Iterator[str], out: Path) -> str:
    total = hashlib.sha256()
    with out.open("w", encoding="utf-8") as f:
        for line in lines:
            f.write(line + "\n")
            total.update(line.encode() + b"\n")
    return total.hexdigest()


def _values(sizes: list[int]) -> Iterator[str]:
    parts = [captures()] + [bench(n, src) for n in sizes for src in ("pickle", "xlsx")]
    for part in parts:
        yield from part
        gc.collect()


def main() -> None:
    """Write the snapshot and print its digest."""
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    large = "--large" in sys.argv
    out = (
        Path(args[0])
        if args
        else ROOT / "capture" / "build" / ("golden_large.txt" if large else "golden.txt")
    )
    sizes = [1000, 10_000, 100_000] + ([1_000_000] if large else [])
    out.parent.mkdir(parents=True, exist_ok=True)
    files_out = out.with_name(out.stem.replace("golden", "golden_files") + out.suffix)
    for lines, path in ((_values(sizes), out), (written_files(sizes), files_out)):
        digest = _write(lines, path)
        print(f"{digest}  {path.relative_to(ROOT) if path.is_relative_to(ROOT) else path}")


if __name__ == "__main__":
    main()
