"""Where the engine's memory goes on the bench table: traced memory per stage, bytes per structure.

    uv run python bench/memory_breakdown.py N [--xlsx]

Runs in the package's own environment, from the repository root. Two measurements in one run:

- ``tracemalloc`` (started before the input is read): Python's traced memory, current and peak,
  after each stage (reading the pickled table, loading, dropping the table, the full
  recalculation, building the dependency index, changing A1), with the process's resident
  memory beside it. Tracing slows the run; its timings are not the benchmark's.
- after loading and again after the recalculation, the bytes each structure holds: the
  container and every object inside it (ints, floats, strings, tuples), each object counted
  once, for the first structure that holds it in the order listed.

Prints one JSON line.
"""

from __future__ import annotations

import gc
import json
import pickle
import sys
import time
import tracemalloc
from array import array
from pathlib import Path
from typing import Any

from xlrecalc._engine import Engine
from xlrecalc._load import load

sys.path.insert(0, str(Path(__file__).resolve().parent))
from engine_bench import DATA, load_table  # noqa: E402
from rss import current_mb  # noqa: E402

_LEAVES = (int, float, str, bool, bytes)


def deep_size(obj: Any, seen: set[int]) -> int:
    """Bytes of a container and everything in it not counted before (iterative)."""
    total = 0
    stack = [obj]
    while stack:
        o = stack.pop()
        if id(o) in seen:
            continue
        seen.add(id(o))
        total += sys.getsizeof(o)
        if isinstance(o, _LEAVES) or isinstance(o, array):
            continue
        if isinstance(o, dict):
            stack.extend(o.keys())
            stack.extend(o.values())
        elif isinstance(o, (list, tuple, set, frozenset)):
            stack.extend(o)
    return total


def structures(engine: Engine) -> dict[str, float]:
    """MB per structure, in attribution order."""
    seen: set[int] = set()
    parts: dict[str, Any] = {}
    for s, sheet in enumerate(engine.sheets):
        parts[f"sheet{s}.values"] = sheet.values
        t = sheet.templates
        parts[f"sheet{s}.templates (key -> template id)"] = (
            (t._dense, t._base, t._sparse, t._counts) if hasattr(t, "_dense") else t
        )
        parts[f"sheet{s}.columns"] = sheet.columns
        parts[f"sheet{s}.formula_columns"] = sheet.formula_columns
    parts["stale"] = engine.stale
    parts["_unclosed"] = engine._unclosed
    parts["_volatile"] = engine._volatile
    parts["_raw_ids + _template_ids"] = (engine._raw_ids, engine._template_ids)
    parts["report"] = (
        engine.report.circular,
        engine.report.uncomputed,
        engine.report.refused,
        engine.report.exceptions,
        engine.report.unsupported,
    )
    if engine._index is not None:
        index = engine._index
        parts["dependency index"] = (
            index.offsets,
            index.edges,
            index.runs,
            index.sheets_of,
            index.climbs,
        )
    return {k: round(deep_size(v, seen) / 1e6, 1) for k, v in parts.items()}


def main() -> None:
    n = int(sys.argv[1])
    xlsx = "--xlsx" in sys.argv
    out: dict[str, Any] = {"N": n, "source": "xlsx" if xlsx else "pickle"}
    out["python"] = sys.version.split()[0]
    stages: list[dict[str, Any]] = []

    def stage(label: str, started: float) -> None:
        gc.collect()
        current, peak = tracemalloc.get_traced_memory()
        stages.append(
            {
                "stage": label,
                "traced_mb": round(current / 1e6, 1),
                "traced_peak_mb": round(peak / 1e6, 1),
                "rss_mb": round(current_mb(), 1),
                "s": round(time.perf_counter() - started, 2),
            }
        )

    out["rss_start_mb"] = current_mb()
    tracemalloc.start()
    table = None
    if not xlsx:
        t = time.perf_counter()
        with (DATA / f"bench_{n}.pkl").open("rb") as f:
            table = pickle.load(f)
        stage("read the pickled table", t)
    t = time.perf_counter()
    engine = load(DATA / f"bench_{n}.xlsx") if xlsx else load_table(table, False)  # type: ignore[arg-type]
    stage("load", t)
    if table is not None:
        t = time.perf_counter()
        del table
        stage("drop the table", t)
    out["after_load"] = structures(engine)
    tracemalloc.reset_peak()  # the measurement's own memory is not the engine's
    t = time.perf_counter()
    engine.recalculate(full=True)
    stage("full recalculation", t)
    t = time.perf_counter()
    engine.index()
    stage("dependency index", t)
    out["after_recalculation"] = structures(engine)
    tracemalloc.reset_peak()
    t = time.perf_counter()
    engine.set_value(0, 0, 0, 1000.0)
    engine.recalculate()
    stage("change A1, re-evaluated", t)
    out["stages"] = stages
    print(json.dumps(out), flush=True)


if __name__ == "__main__":
    main()
