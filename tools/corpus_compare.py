"""Compare the engine with the values Excel cached in a corpus of workbooks.

    uv run python tools/corpus_compare.py DIR [DIR...] --out OUT [--sample N] [--seed S]
        [--jobs 4] [--batch 150] [--limit 600]

Run from the repository root. Every workbook under the directories is evaluated in batches, each
in a child under the memory watchdog (``tests/watchdog.py``), with a time limit per workbook; one
JSON line per workbook goes to ``OUT/rows``, a finished workbook is not evaluated again, and the
run stops at the first batch killed, naming the workbook it was reading. A workbook named in
``OUT/skip.txt`` is left out. ``OUT`` then gets ``summary.txt`` and ``mismatches.tsv``. They
describe someone else's workbooks, so they stay local, as the corpus does.

Each formula is evaluated twice:

- **alone**: against Excel's cached values for everything it reads. A difference here is the
  formula's own, or a stale cache; no other formula can have caused it.
- **in a full recalculation**, against the engine's values. A cell that matches alone and
  differs here only inherits a difference.

Categories, the first that applies to a formula cell:

- ``no_cached``: the file holds no value to compare with.
- ``volatile``, ``volatile_dependent``, ``external``, ``external_dependent``: set aside. Volatile
  here means a result that depends on when or where it is computed (``NOW``, ``TODAY``, ``RAND``,
  ``RANDBETWEEN``, ``RANDARRAY``, ``INFO``, ``CELL``), not ``OFFSET`` or ``INDIRECT``, which
  are compared.
- ``iterating``: on or after a cycle with iteration on; it keeps its cached value.
- ``circular_empty``, ``circular_alone_matches``, ``circular_alone_differs``: on or after a cycle
  with iteration off. The first is what Excel saves for such a cell. In the others the cache holds
  a value, so the cache is older than the cycle or the engine calls circular what Excel does not.
- ``stopped_own``, by label: an absent function or a form not implemented.
- ``exception_own``: an internal error, alone.
- ``stopped_inherited``: computes alone, but reads a stopped cell in the full recalculation.
- ``differ_own``: differs alone.
- ``differ_inherited``: matches alone, differs in the full recalculation.
- ``match_exact``, ``match_15_digits`` (relative 1E-14), ``match_close`` (relative 1E-9),
  ``match_tiny`` (both below an absolute 1E-9 of each other): the worse of the two evaluations.

A workbook where at least half of the formulas compared differ alone, and at least 20 do, is
``suspect``: its cache may be stale. Its cells are counted apart from the others.
"""

from __future__ import annotations

import argparse
import collections
import json
import os
import random
import re
import signal
import subprocess
import sys
import time
from collections import Counter
from collections.abc import Iterable, Iterator
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

from xlrecalc._arrays import ArrayLimitError, NotImplementedFormError, element, is_array, shape
from xlrecalc._engine import COLS, Engine, node, unpack
from xlrecalc._load import _value, load
from xlrecalc._report import CellRef
from xlrecalc._values import BLANK, ErrorValue, Ref
from xlrecalc._xlsx import Book

UNTIMELY = frozenset({"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "INFO", "CELL"})
_STRING = re.compile(r'"(?:[^"]|"")*"')
_BOOK = re.compile(r"\[[^\]]*\]")
COMPARED = (
    "match_exact",
    "match_15_digits",
    "match_close",
    "match_tiny",
    "differ_own",
    "differ_inherited",
    "stopped_own",
    "stopped_inherited",
    "exception_own",
)
SET_ASIDE = (
    "no_cached",
    "volatile",
    "volatile_dependent",
    "external",
    "external_dependent",
    "iterating",
    "circular_empty",
    "circular_alone_matches",
    "circular_alone_differs",
)
_RANK = {"exact": 0, "15_digits": 1, "close": 2, "tiny": 3, "differ": 4}


class _Timeout(BaseException):
    """Through the engine's per-cell ``except Exception``."""


def compare(cached: Any, got: Any) -> str:
    """How close two values are: ``exact``, ``15_digits``, ``close``, ``tiny`` or ``differ``."""
    if type(cached) is not type(got):
        return "differ"
    if isinstance(cached, float):
        if cached == got:
            return "exact"
        gap = abs(cached - got)
        size = max(abs(cached), abs(got))
        if gap <= 1e-14 * size:
            return "15_digits"
        if gap <= 1e-9 * size:
            return "close"
        return "tiny" if gap <= 1e-9 else "differ"
    return "exact" if cached == got else "differ"


def kind_of(value: Any) -> str:
    """A value's type for the report: an error by its code."""
    if isinstance(value, ErrorValue):
        return str(value)
    if isinstance(value, bool):
        return "logical"
    if isinstance(value, float):
        return "number"
    return "text" if isinstance(value, str) else "blank"


def _worse(a: str, b: str) -> str:
    return a if _RANK[a] >= _RANK[b] else b


def _external(text: str, names: frozenset[str]) -> bool:
    """A formula's text refers to another workbook, directly or through a name."""
    bare = _STRING.sub('""', text)
    if _BOOK.search(bare):
        return True
    if not names:
        return False
    return any(word.casefold() in names for word in re.findall(r"[A-Za-z_\\][\w.\\]*", bare))


def _alone(engine: Engine, n: int, template: Any) -> tuple[list[str], str | None, Any]:
    """Evaluate one formula against the store: labels stopping it, an exception, its result."""
    sheet, row, col = unpack(n)
    if template.unsupported:
        return sorted(template.unsupported), None, None
    engine.sheet, engine.row, engine.col = sheet, row, col
    engine.reference_array = None
    labels = engine.report.unsupported
    labels.clear()
    try:
        value = template.fn(engine, row, col)
        if engine.reference_array is not None and value is not engine.reference_array:
            raise NotImplementedFormError("an array of references read by a function")
    except ArrayLimitError as exc:
        return [f"an array of more than {exc.limit:,} cells"], None, None
    except NotImplementedFormError as exc:
        return [exc.label], None, None
    except Exception as exc:
        return [], type(exc).__name__, None
    if labels:
        return sorted(labels), None, None
    return [], None, value


def _members(engine: Engine, n: int) -> list[tuple[int, int]]:
    """An anchor's cells other than its own, 0-based."""
    _, row, col = unpack(n)
    r1, c1, r2, c2 = engine.anchors[n].box
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) != (row, col)]


def _cached_values(engine: Engine, path: str) -> dict[int, Any]:
    """Excel's cached value of every formula cell and every cell an array formula fills."""
    out: dict[int, Any] = {}
    with Book(path) as book:
        for index, info in enumerate(book.sheets):
            if info.kind != "worksheet":
                continue
            for cell in book.cells(info):
                n = node(index, cell.row - 1, cell.column - 1)
                if cell.formula is None and engine.member_of(n) is None:
                    continue
                value = _value(cell, engine.date1904, engine.quirks, book.value_errors)
                if value is not None:
                    out[n] = value
    return out


def compare_workbook(path: str) -> dict[str, Any]:
    """One workbook's counts, labels and distinct mismatches."""
    engine = load(path, trust_cached=False)
    cached = _cached_values(engine, path)
    formulas = [(index << 34) | key for index, s in enumerate(engine.sheets) for key in s.templates]
    engine.recalculate(full=True)
    report = engine.report
    circular = {node(_index(engine, c), c.row - 1, c.col - 1) for c in report.circular}
    uncomputed = {node(_index(engine, c), c.row - 1, c.col - 1) for c in report.uncomputed}
    kept = set(engine._kept) if engine.iterate else set()
    full: dict[int, Any] = {}
    for n in formulas:
        sheet, row, col = unpack(n)
        values = engine.sheets[sheet].values
        full[n] = values.get(row * COLS + col, BLANK)
        if n in engine.anchors:
            for r, c in _members(engine, n):
                m = node(sheet, r, c)
                full[m] = values.get(r * COLS + c, BLANK)

    # From here the store holds what Excel saved, and no cell is stale.
    for n in full:
        sheet, row, col = unpack(n)
        store = engine.sheets[sheet]
        store.set(row, col, cached.get(n, BLANK))
    engine.stale.clear()
    engine._unclosed.clear()
    engine._kept.clear()
    engine.iterate = False
    engine._changed = False

    names = frozenset(
        name.casefold() for (_, name), text in engine.names.items() if _BOOK.search(text)
    )
    untimely: list[int] = []
    outside: list[int] = []
    reaches_out: dict[int, tuple[bool, bool]] = {}
    for n in formulas:
        sheet, row, col = unpack(n)
        tid = engine.sheets[sheet].templates.at(row, col)
        if tid not in reaches_out:
            template = engine.templates[tid]
            reaches_out[tid] = (
                bool(template.calls & UNTIMELY),
                _external(template.origin[0], names),
            )
        when, where = reaches_out[tid]
        if when:
            untimely.append(n)
        elif where:
            outside.append(n)
    untimely_set, outside_set = set(untimely), set(outside)
    untimely_after = set(engine.dependents(untimely)) if untimely else set()
    outside_after = set(engine.dependents(outside)) if outside else set()

    counts: Counter[str] = Counter()
    labels: Counter[str] = Counter()
    label_seeds: dict[str, list[int]] = collections.defaultdict(list)
    exceptions: Counter[str] = Counter()
    mismatches: dict[tuple[str, str, str], dict[str, Any]] = {}
    alone_differs: list[int] = []
    for n in formulas:
        sheet, row, col = unpack(n)
        store = engine.sheets[sheet]
        template = engine.templates[store.templates.at(row, col)]
        value = cached.get(n)
        if value is None:
            counts["no_cached"] += 1
        elif n in untimely_set:
            counts["volatile"] += 1
        elif n in outside_set:
            counts["external"] += 1
        elif n in untimely_after:
            counts["volatile_dependent"] += 1
        elif n in outside_after:
            counts["external_dependent"] += 1
        elif n in kept:
            counts["iterating"] += 1
        elif n in circular and value == "":
            counts["circular_empty"] += 1
        else:
            stopped, exception, result = _alone(engine, n, template)
            if n in circular:
                same = not stopped and not exception and _same_alone(engine, n, result, cached)
                counts["circular_alone_matches" if same else "circular_alone_differs"] += 1
                continue
            if stopped:
                counts["stopped_own"] += 1
                for label in stopped:
                    labels[label] += 1
                    label_seeds[label].append(n)
                continue
            if exception:
                counts["exception_own"] += 1
                exceptions[exception] += 1
                continue
            alone = _worst_alone(engine, n, result, cached)
            if alone[0] == "differ":
                counts["differ_own"] += 1
                alone_differs.append(n)
                _note(mismatches, "own", template, engine, n, value, alone[1], store)
            elif n in uncomputed:
                counts["stopped_inherited"] += 1
            else:
                whole = compare(value, full[n])
                for r, c in _members(engine, n) if n in engine.anchors else ():
                    m = node(sheet, r, c)
                    whole = _worse(whole, compare(cached.get(m, BLANK), full[m]))
                if whole == "differ":
                    counts["differ_inherited"] += 1
                else:
                    counts["match_" + _worse(alone[0], whole)] += 1

    downstream = {label: len(set(engine.dependents(seeds))) for label, seeds in label_seeds.items()}
    after_own = len(set(engine.dependents(alone_differs))) if alone_differs else 0
    with Book(path) as book:
        calc = book.calc
        tables = 0
        for info in book.sheets:
            if info.kind == "worksheet":
                tables += sum(
                    1 for c in book.cells(info) if c.formula and c.formula.kind == "dataTable"
                )
    return {
        "counts": dict(counts),
        "labels": dict(labels),
        "downstream": downstream,
        "exceptions": dict(exceptions),
        "after_own": after_own,
        "refused": len(report.refused),
        "data_tables": tables,
        "calc_mode": calc.calc_mode,
        "flag": calc.full_calc_on_load,
        "iterate": calc.iterate,
        "mismatches": list(mismatches.values()),
    }


def _index(engine: Engine, cell: CellRef) -> int:
    index = engine.sheet_index(cell.sheet)
    assert index is not None
    return index


def _same_alone(engine: Engine, n: int, result: Any, cached: dict[int, Any]) -> bool:
    return _worst_alone(engine, n, result, cached)[0] != "differ"


def _worst_alone(engine: Engine, n: int, result: Any, cached: dict[int, Any]) -> tuple[str, Any]:
    """How a result compares with the cache, an array over its range, with the value differing."""
    sheet = unpack(n)[0]
    if n not in engine.anchors:
        got = engine._result(result)
        return compare(cached.get(n, BLANK), got), got
    r1, c1, r2, c2 = engine.anchors[n].box
    grid = result if is_array(result) else ((result,),)
    rows, cols = shape(grid)
    worst, shown = "exact", None
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            v = element(grid, rows, cols, r - r1, c - c1)
            if isinstance(v, Ref) and v.source is engine.checked:
                v = ErrorValue.VALUE
            got = engine._result(v)
            how = compare(cached.get(node(sheet, r, c), BLANK), got)
            if _RANK[how] > _RANK[worst] or shown is None:
                shown = got
            worst = _worse(worst, how)
    return worst, shown


def _note(
    mismatches: dict[tuple[str, str, str], dict[str, Any]],
    scope: str,
    template: Any,
    engine: Engine,
    n: int,
    cached: Any,
    got: Any,
    store: Any,
) -> None:
    kind = f"{kind_of(cached)} -> {kind_of(got)}"
    key = (scope, template.key, kind)
    row = mismatches.get(key)
    if row is not None:
        row["cells"] += 1
        return
    sheet, r, c = unpack(n)
    mismatches[key] = {
        "key": template.key,
        "kind": kind,
        "functions": sorted(template.calls),
        "cells": 1,
        "cell": str(CellRef(store.name, r + 1, c + 1)),
        "formula": (engine.formula_text(sheet, r, c) or "")[:300],
        "cached": _short(cached),
        "got": _short(got),
    }


def _short(value: Any) -> str:
    text = str(value) if isinstance(value, ErrorValue) else repr(value)
    return text[:80]


# Worker -----------------------------------------------------------------------------------------


def worker(listing: str, limit: int, out: str) -> None:
    """One batch: one JSON line per workbook, flushed, so that a killed batch names the next."""

    def on_alarm(*_: Any) -> None:
        raise _Timeout

    signal.signal(signal.SIGALRM, on_alarm)
    with open(out, "a", encoding="utf-8") as sink:
        for path in Path(listing).read_text(encoding="utf-8").splitlines():
            row: dict[str, Any] = {"path": path}
            start = time.perf_counter()
            signal.alarm(limit)
            try:
                row.update(compare_workbook(path))
            except _Timeout:
                row["timeout"] = True
            except Exception as exc:
                row["error"] = f"{type(exc).__name__}: {exc}"[:200]
            finally:
                signal.alarm(0)
            row["seconds"] = round(time.perf_counter() - start, 2)
            sink.write(json.dumps(row) + "\n")
            sink.flush()


# Driver -----------------------------------------------------------------------------------------


def _workbooks(directories: Iterable[str]) -> list[str]:
    found = []
    for directory in directories:
        for folder, _, names in os.walk(directory):
            found += [
                os.path.abspath(os.path.join(folder, name))
                for name in names
                if name.lower().endswith(".xlsx")  # never .xlsm (CAPTURING.md §4)
            ]
    return sorted(found)


def _rows(out: Path) -> Iterator[dict[str, Any]]:
    for part in sorted((out / "rows").glob("*.jsonl")):
        for line in part.read_text(encoding="utf-8").splitlines():
            yield json.loads(line)


def _run_batch(out: Path, number: int, paths: list[str], limit: int) -> tuple[int, int, str]:
    listing = out / "batches" / f"{number:05}.txt"
    listing.write_text("\n".join(paths) + "\n", encoding="utf-8")
    sink = out / "rows" / f"{number:05}.jsonl"
    command = [
        sys.executable,
        "-m",
        "tests.watchdog",
        "--",
        sys.executable,
        "-I",
        __file__,
        "--worker",
        str(listing),
        "--limit",
        str(limit),
        "--sink",
        str(sink),
    ]
    log = out / "batches" / f"{number:05}.log"
    with log.open("w", encoding="utf-8") as stderr:
        code = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=stderr).returncode
    done = len(sink.read_text(encoding="utf-8").splitlines()) if sink.exists() else 0
    return code, number, paths[done] if code and done < len(paths) else ""


def drive(args: argparse.Namespace) -> int:
    out = Path(args.out)
    for folder in ("rows", "batches"):
        (out / folder).mkdir(parents=True, exist_ok=True)
    skip_file = out / "skip.txt"
    skip = set(skip_file.read_text(encoding="utf-8").split("\n")) if skip_file.exists() else set()
    paths = _workbooks(args.directories)
    if args.sample:
        paths = random.Random(args.seed).sample(paths, min(args.sample, len(paths)))
    done = {row["path"] for row in _rows(out)}
    todo = [p for p in paths if p not in done and p not in skip]
    print(f"{len(paths)} workbooks, {len(done)} done, {len(todo)} to do", flush=True)
    first = 1 + max((int(p.stem) for p in (out / "rows").glob("*.jsonl")), default=0)
    batches = [todo[i : i + args.batch] for i in range(0, len(todo), args.batch)]
    killed = False
    with ThreadPoolExecutor(args.jobs) as pool:
        pending = iter(enumerate(batches, first))

        def run_next() -> None:
            nonlocal killed
            for number, batch in pending:
                if killed:
                    return
                code, _, culprit = _run_batch(out, number, batch, args.limit)
                print(f"batch {number}: return code {code}", flush=True)
                if code:
                    killed = True
                    print(f"KILLED: batch {number} while reading {culprit}", flush=True)
                    return

        for future in [pool.submit(run_next) for _ in range(args.jobs)]:
            future.result()
    summarize(out, set(paths))
    return 1 if killed else 0


# Report -----------------------------------------------------------------------------------------


def _suspect(counts: dict[str, int]) -> bool:
    compared = sum(counts.get(c, 0) for c in COMPARED if not c.startswith("stopped_own"))
    own = counts.get("differ_own", 0)
    return own >= 20 and own * 2 >= compared


def summarize(out: Path, paths: set[str] | None = None) -> None:
    rows = [r for r in _rows(out) if paths is None or r["path"] in paths]
    lines: list[str] = []
    timeouts = [r for r in rows if r.get("timeout")]
    errors = [r for r in rows if r.get("error")]
    good = [r for r in rows if "counts" in r]
    seconds = [r["seconds"] for r in rows]
    lines.append(
        f"workbooks {len(rows)}, compared {len(good)}, timed out {len(timeouts)}, "
        f"unreadable {len(errors)}; {sum(seconds):,.0f} s in all, "
        f"{sum(s > 10 for s in seconds)} past 10 s, slowest {max(seconds, default=0):.0f} s"
    )
    suspects = [r for r in good if _suspect(r["counts"])]
    plain = [r for r in good if not _suspect(r["counts"])]
    lines.append(f"suspect workbooks (half or more differ alone): {len(suspects)}")
    lines.append(
        f"iteration on: {sum(r['iterate'] for r in good)}; manual calculation: "
        f"{sum(r['calc_mode'] == 'manual' for r in good)}; fullCalcOnLoad: "
        f"{sum(r['flag'] for r in good)}"
    )
    lines.append("")
    total = Counter[str]()
    apart = Counter[str]()
    books = Counter[str]()
    for row in plain:
        total.update(row["counts"])
        books.update(row["counts"].keys())
    for row in suspects:
        apart.update(row["counts"])
    lines.append(
        f"{'category':26} {'cells':>12} {'share':>7} {'workbooks':>10} {'in suspects':>12}"
    )
    compared = sum(total[c] for c in COMPARED)
    for name in COMPARED:
        share = f"{total[name] / compared:.2%}" if compared else ""
        lines.append(
            f"{name:26} {total[name]:>12,} {share:>7} {books[name]:>10,} {apart[name]:>12,}"
        )
    lines.append(f"{'compared, in all':26} {compared:>12,}")
    lines.append("")
    lines += [
        f"{name:26} {total[name]:>12,} {'':>7} {books[name]:>10,} {apart[name]:>12,}"
        for name in SET_ASIDE
    ]
    lines.append(f"{'refused formulas':26} {sum(r['refused'] for r in good):>12,}")
    lines.append(f"{'data tables':26} {sum(r['data_tables'] for r in good):>12,}")
    lines.append("")

    own = Counter[str]()
    after = Counter[str]()
    label_books = Counter[str]()
    for row in good:
        own.update(row["labels"])
        after.update(row["downstream"])
        label_books.update(row["labels"].keys())
    lines.append("what stops cells: own cells, cells downstream of them (may overlap), workbooks")
    for label, cells in sorted(own.items(), key=lambda kv: (-kv[1], kv[0])):
        lines.append(f"{label[:40]:40} {cells:>10,} {after[label]:>10,} {label_books[label]:>8,}")
    lines.append("")
    lines.append("the same, by workbooks")
    for label, count in label_books.most_common(40):
        lines.append(f"{label[:40]:40} {count:>8,} {own[label]:>10,}")
    lines.append("")
    exceptions = Counter[str]()
    for row in good:
        exceptions.update(row["exceptions"])
    lines.append(f"exceptions, alone: {dict(exceptions)}")
    lines.append(f"timed out: {[Path(r['path']).name for r in timeouts]}")
    lines.append(f"unreadable: {[(Path(r['path']).name, r['error']) for r in errors]}")
    lines.append("")

    # Distinct mismatches, outside suspect workbooks first.
    groups: dict[tuple[str, str], dict[str, Any]] = {}
    for row in good:
        suspect = _suspect(row["counts"])
        for m in row["mismatches"]:
            group = groups.setdefault(
                (m["key"], m["kind"]),
                {**m, "cells": 0, "plain": 0, "workbooks": 0, "path": Path(row["path"]).name},
            )
            group["cells"] += m["cells"]
            group["workbooks"] += 1
            if not suspect:
                group["plain"] += m["cells"]
    by_functions: Counter[str] = Counter()
    by_kind: Counter[str] = Counter()
    distinct_plain = 0
    for group in groups.values():
        if group["plain"]:
            distinct_plain += 1
            by_functions[" ".join(group["functions"]) or "(operators)"] += group["plain"]
            by_kind[group["kind"]] += group["plain"]
    lines.append(f"distinct formulas differing alone, outside suspects: {distinct_plain:,}")
    lines.append("by kind (cached -> computed), cells")
    for kind, cells in by_kind.most_common(40):
        lines.append(f"  {kind:30} {cells:>10,}")
    lines.append("by the functions in the formula, cells")
    for functions, cells in by_functions.most_common(60):
        lines.append(f"  {functions[:60]:60} {cells:>10,}")
    lines.append("")
    lines.append("suspect workbooks: differ alone, compared, calculation mode")
    for row in sorted(suspects, key=lambda r: -r["counts"].get("differ_own", 0))[:200]:
        counts = row["counts"]
        lines.append(
            f"  {Path(row['path']).name[:60]:60} {counts.get('differ_own', 0):>8,} "
            f"{sum(counts.get(c, 0) for c in COMPARED):>8,} {row['calc_mode']}"
        )
    (out / "summary.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    header = (
        "cells\tplain_cells\tworkbooks\tkind\tfunctions\tformula\tcached\tcomputed\tcell\tworkbook"
    )
    table = [header]
    for group in sorted(groups.values(), key=lambda g: (-g["plain"], -g["cells"], g["key"])):
        fields = [
            group["cells"],
            group["plain"],
            group["workbooks"],
            group["kind"],
            " ".join(group["functions"]),
            group["formula"],
            group["cached"],
            group["got"],
            group["cell"],
            group["path"],
        ]
        table.append("\t".join(str(f).replace("\t", " ").replace("\n", " ") for f in fields))
    (out / "mismatches.tsv").write_text("\n".join(table) + "\n", encoding="utf-8")
    print("\n".join(lines[:60]))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("directories", nargs="*")
    parser.add_argument("--out")
    parser.add_argument("--sample", type=int)
    parser.add_argument("--seed", type=int, default=1)
    parser.add_argument("--jobs", type=int, default=4)
    parser.add_argument("--batch", type=int, default=150)
    parser.add_argument("--limit", type=int, default=600, help="seconds per workbook")
    parser.add_argument("--worker", help=argparse.SUPPRESS)
    parser.add_argument("--sink", help=argparse.SUPPRESS)
    parser.add_argument("--summarize", action="store_true", help="only rewrite the report")
    args = parser.parse_args()
    if args.worker:
        worker(args.worker, args.limit, args.sink)
        return 0
    if not args.out:
        parser.error("--out is required")
    if args.summarize:
        summarize(Path(args.out))
        return 0
    if not args.directories:
        parser.error("name at least one directory")
    return drive(args)


if __name__ == "__main__":
    sys.exit(main())
