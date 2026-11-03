"""``recalculate(path)`` on openpyxl's benchmark workbooks: load, evaluate, write (§9.2).

    uv run python bench/recalc_bench.py N [OUT]

Runs in the package's own environment, from the repository root. Copies ``bench/data/
bench_<N>.xlsx`` (openpyxl's file: no cached values, so every formula is evaluated) to ``OUT``
(default: a temporary directory), then does what ``recalculate`` does, timing each stage: the
load through the package's reader, the full evaluation, and the surgical write in place. Then
writes the file once more, now holding its values, to time a write that changes nothing. Prints
one JSON line.
"""

from __future__ import annotations

import json
import resource
import shutil
import sys
import tempfile
import time
from pathlib import Path

from xlrecalc._load import load
from xlrecalc._save import lookups
from xlrecalc._xlsx._write import write

DATA = Path(__file__).resolve().parent / "data"


def peak_mb() -> float:
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak / 1e6 if sys.platform == "darwin" else peak / 1e3


def main(n: int, out: Path | None) -> None:
    folder = out or Path(tempfile.mkdtemp())
    path = folder / f"bench_{n}.xlsx"
    shutil.copyfile(DATA / f"bench_{n}.xlsx", path)
    t0 = time.perf_counter()
    engine = load(path)
    t1 = time.perf_counter()
    engine.recalculate()
    t2 = time.perf_counter()
    written = write(path, path, lookups(engine))
    t3 = time.perf_counter()
    unchanged = write(path, folder / "again.xlsx", lookups(engine))
    t4 = time.perf_counter()
    print(
        json.dumps(
            {
                "rows": n,
                "formulas": engine.formula_count(),
                "load_s": round(t1 - t0, 2),
                "eval_s": round(t2 - t1, 2),
                "write_s": round(t3 - t2, 2),
                "total_s": round(t3 - t0, 2),
                "cells_written": written,
                "rewrite_unchanged_s": round(t4 - t3, 2),
                "cells_rewritten": unchanged,
                "size_before_mb": round((DATA / f"bench_{n}.xlsx").stat().st_size / 1e6, 1),
                "size_after_mb": round(path.stat().st_size / 1e6, 1),
                "peak_mb": round(peak_mb()),
                "file": str(path),
            }
        )
    )


if __name__ == "__main__":
    main(int(sys.argv[1]), Path(sys.argv[2]) if len(sys.argv) > 2 else None)
