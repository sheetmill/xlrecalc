"""Compare what xlrecalc's reader and openpyxl read from the same workbooks, cell by cell.

    uv run python bench/read_compare.py FILE_OR_DIRECTORY ... [--sample N] [--seed S]

Runs in the package's own environment (it needs xlrecalc and openpyxl). For every worksheet,
openpyxl is read twice in read-only mode with its date conversion off (formulas, then cached
values); xlrecalc's reader once. Differences are classified and counted, with examples. A
shared formula's dependent cells hold no text in the file; openpyxl writes the translated
formula there, the reader reports the reference to the master, so those cells are compared by
value only. Only reads files; prints a report.
"""

from __future__ import annotations

import argparse
import collections
import datetime as dt
import random
import sys
import warnings
from pathlib import Path

import openpyxl

from xlrecalc._xlsx import Book, PackageError


def _openpyxl(path: Path, data_only: bool) -> dict[str, dict[tuple[int, int], tuple[object, str]]]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=data_only, keep_links=False)
    wb._date_formats = set()  # noqa: SLF001 - keep serials, as sheetpeek does
    wb._timedelta_formats = set()  # noqa: SLF001
    out = {}
    try:
        for ws in wb.worksheets:
            cells = {}
            for row in ws.iter_rows():
                for c in row:
                    if getattr(c, "value", None) is not None:
                        cells[(c.row, c.column)] = (c.value, c.data_type)
            out[ws.title] = cells
    finally:
        wb.close()
    return out


def _ours(path: Path) -> dict[str, dict[tuple[int, int], object]]:
    out = {}
    with Book(path) as book:
        for sheet in book.sheets:
            if sheet.kind == "worksheet":
                out[sheet.name] = {(c.row, c.column): c for c in book.cells(sheet)}
    return out


def _value(cell: object) -> object:
    return cell.value  # type: ignore[attr-defined]


def compare(path: Path, diffs: collections.Counter[str], examples: dict[str, list[str]],
            totals: collections.Counter[str]) -> None:
    ours = _ours(path)
    formulas = _openpyxl(path, data_only=False)
    values = _openpyxl(path, data_only=True)

    def note(kind: str, text: str) -> None:
        diffs[kind] += 1
        if len(examples[kind]) < 4:
            examples[kind].append(f"{path.name}: {text}")

    for name, cells in ours.items():
        theirs_f, theirs_v = formulas.get(name, {}), values.get(name, {})
        for pos, cell in cells.items():
            totals["cells"] += 1
            f = cell.formula  # type: ignore[attr-defined]
            # Formula text.
            if f is not None:
                totals["formulas"] += 1
                other = theirs_f.get(pos, (None, None))[0]
                if f.text is None:
                    totals["shared dependents"] += 1
                elif isinstance(other, str) and other.startswith("="):
                    if other[1:] != f.text:
                        note("formula text", f"{pos} ours {f.text!r} openpyxl {other!r}")
                elif other is None:
                    note("formula openpyxl lacks", f"{pos} {f.kind} {f.text!r}")
                else:
                    note(f"formula as {type(other).__name__}", f"{pos} {f.kind} {f.text!r}: {other!r}"[:160])
            # Cached value.
            mine = cell.value  # type: ignore[attr-defined]
            kind = cell.type  # type: ignore[attr-defined]
            other_v, other_t = theirs_v.get(pos, (None, None))
            if mine is None and other_v is None:
                continue
            if kind == "e":
                if other_v != mine or other_t != "e":
                    note("error", f"{pos} ours {mine!r} openpyxl {other_v!r} ({other_t})")
            elif kind == "b":
                if other_v is not mine:
                    note("logical", f"{pos} ours {mine!r} openpyxl {other_v!r}")
            elif kind == "n":
                if isinstance(other_v, bool) or not isinstance(other_v, int | float):
                    note(f"number as {type(other_v).__name__}", f"{pos} ours {mine!r} openpyxl {other_v!r}")
                elif float(other_v) != mine:
                    note("number value", f"{pos} ours {mine!r} openpyxl {other_v!r}")
                elif isinstance(other_v, int):
                    totals["integral numbers openpyxl gives as int"] += 1
            elif kind in ("s", "str", "inlineStr"):
                if other_v is None and mine == "":
                    note("empty text openpyxl reads as empty cell", f"{pos} {kind}")
                elif other_v != mine:
                    note(f"text ({kind})", f"{pos} ours {mine!r} openpyxl {other_v!r}"[:200])
            elif kind == "d":
                if not isinstance(other_v, dt.datetime | dt.date | dt.time):
                    note("ISO date", f"{pos} ours {mine!r} openpyxl {other_v!r}")
            elif kind == "none":
                note("value openpyxl has, ours none", f"{pos} openpyxl {other_v!r}")
        for pos in set(theirs_v) - set(cells):
            note("cell only openpyxl has", f"{name}!{pos} {theirs_v[pos]!r}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--sample", type=int, default=0)
    parser.add_argument("--seed", type=int, default=1)
    args = parser.parse_args()
    files: list[Path] = []
    for p in args.paths:
        if p.exists():
            files.extend(sorted(p.rglob("*.xlsx")) if p.is_dir() else [p])
    if not files:
        print("no workbooks found (local corpora are optional)")
        return 0
    if args.sample:
        files = random.Random(args.seed).sample(files, min(args.sample, len(files)))
    diffs: collections.Counter[str] = collections.Counter()
    examples: dict[str, list[str]] = collections.defaultdict(list)
    totals: collections.Counter[str] = collections.Counter()
    failed = collections.Counter[str]()
    warnings.simplefilter("ignore")
    for path in files:
        try:
            compare(path, diffs, examples, totals)
        except (PackageError, KeyError, ValueError, OSError) as exc:
            failed[type(exc).__name__] += 1
            if failed[type(exc).__name__] <= 3:
                print(f"failed: {path.name}: {type(exc).__name__}: {exc}"[:200], file=sys.stderr)
    print(f"{len(files)} workbooks ({sum(failed.values())} failed: {dict(failed)}); "
          + ", ".join(f"{k} {v}" for k, v in totals.items()))
    for kind, n in diffs.most_common():
        print(f"{n:9d}  {kind}")
        for e in examples[kind]:
            print(f"             {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
