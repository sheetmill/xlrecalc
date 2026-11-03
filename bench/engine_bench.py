"""The engine on the benchmark table: load, full recalculation, and changes (§6.4 acceptance).

    uv run python bench/engine_bench.py N [--xlsx] [--parse-every-formula]

Runs in the package's own environment, from the repository root. Loads ``bench/data/
bench_<N>.pkl`` (the table ``gen.py`` writes) into an engine, or with ``--xlsx`` reads
``bench/data/bench_<N>.xlsx`` (openpyxl's file) through the package's reader. Then: a full
recalculation; the dependency index build; changing A1 (the dirty set alone, then the change
re-evaluated); changing the last row's A. ``--parse-every-formula`` turns off the raw-text
template key (§5.3), to measure what it saves. Prints one JSON line, like the prototypes.
"""

from __future__ import annotations

import json
import pickle
import sys
import time
from pathlib import Path

from xlrecalc._engine import Engine, node
from xlrecalc._formula import parse, template_key
from xlrecalc._load import load

sys.path.insert(0, str(Path(__file__).resolve().parent))
from rss import current_mb, peak_mb  # noqa: E402

DATA = Path(__file__).resolve().parent / "data"


def parse_every_formula(engine: Engine) -> None:
    """Make the engine parse each formula, as without the raw-text key."""

    def template_for(text: str, row: int, col: int, raw: str | None = None) -> int:
        formula = parse(text)
        key = template_key(formula, row, col)
        tid = engine._template_ids.get(key)
        return engine._compile(formula, key, row, col) if tid is None else tid

    engine.template_for = template_for  # type: ignore[method-assign]


def load_table(table: list[list[object]], slow: bool) -> Engine:
    engine = Engine()
    if slow:
        parse_every_formula(engine)
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


def main() -> None:
    n = int(sys.argv[1])
    xlsx = "--xlsx" in sys.argv
    slow = "--parse-every-formula" in sys.argv
    out: dict[str, object] = {"pkg": "xlrecalc", "N": n, "source": "xlsx" if xlsx else "pickle"}
    out["python"] = sys.version.split()[0]
    out["rss_start_mb"] = current_mb()
    table = None
    if not xlsx:  # read the table first, outside the load, as the prototypes do
        with (DATA / f"bench_{n}.pkl").open("rb") as f:
            table = pickle.load(f)
        out["rss_table_mb"] = current_mb()
    if slow:
        out["raw_key"] = False
    t = time.perf_counter()
    if xlsx:
        if slow:
            import xlrecalc._engine as engine_module

            original = engine_module.Engine.__init__

            def init(self: Engine, *a: object, **k: object) -> None:
                original(self, *a, **k)  # type: ignore[arg-type]
                parse_every_formula(self)

            engine_module.Engine.__init__ = init  # type: ignore[method-assign]
        engine = load(DATA / f"bench_{n}.xlsx")
    else:
        engine = load_table(table, slow)
    out["load_s"] = time.perf_counter() - t
    out["n_formulas"] = engine.formula_count()
    out["n_templates"] = len(engine.templates)
    out["rss_after_load_mb"] = current_mb()
    out["peak_after_load_mb"] = peak_mb()
    if table is not None:
        del table
        out["rss_after_load_no_table_mb"] = current_mb()

    t = time.perf_counter()
    engine.recalculate(full=True)
    out["eval_s"] = time.perf_counter() - t
    out["values"] = [engine.get(0, r, 5) for r in range(3)]

    t = time.perf_counter()
    engine.index()
    out["index_build_s"] = time.perf_counter() - t

    a1 = node(0, 0, 0)
    t = time.perf_counter()
    runs = engine.dirty_runs([a1])
    out["dirty_only_A1_s"] = time.perf_counter() - t
    out["dirty_A1"] = sum(b - a + 1 for rs in runs.values() for a, b in rs)
    out["dirty_runs_A1"] = sum(len(rs) for rs in runs.values())

    t = time.perf_counter()
    engine.set_value(0, 0, 0, 1000.0)
    engine.recalculate()
    out["reeval_A1_s"] = time.perf_counter() - t
    out["values2"] = [engine.get(0, r, 5) for r in range(3)]

    last = node(0, n - 1, 0)
    t = time.perf_counter()
    engine.dirty_runs([last])
    out["dirty_only_Alast_ms"] = (time.perf_counter() - t) * 1000
    t = time.perf_counter()
    engine.set_value(0, n - 1, 0, 5.0)
    engine.recalculate()
    out["reeval_Alast_s"] = time.perf_counter() - t
    out["rss_end_mb"] = current_mb()
    out["peak_rss_mb"] = peak_mb()
    print(json.dumps(out), flush=True)


if __name__ == "__main__":
    main()
