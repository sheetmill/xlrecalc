"""Time and memory of reading every cell: xlrecalc's reader against openpyxl.

    uv run python bench/read_bench.py FILE_OR_DIRECTORY ... [--sample N] [--seed S]

Runs in the package's own environment. Each reader runs in a fresh process over the same files
and reads every cell of every worksheet, keeping nothing; the child reports its time and peak
resident memory (the import of the reader included). Modes:

- ``xlrecalc``: one pass, formulas and cached values together;
- ``openpyxl read-only``: formulas (``read_only=True``);
- ``openpyxl read-only data_only``: cached values (the engine needs both passes);
- ``openpyxl normal``: ``load_workbook`` in normal mode, formulas.
"""

from __future__ import annotations

import argparse
import json
import random
import resource
import subprocess
import sys
import time
from pathlib import Path

MODES = ["xlrecalc", "openpyxl read-only", "openpyxl read-only data_only", "openpyxl normal"]


def child(mode: str, files: list[str]) -> None:
    start = time.perf_counter()
    cells = 0
    if mode == "xlrecalc":
        from xlrecalc._xlsx import Book

        for f in files:
            with Book(f) as book:
                for sheet in book.sheets:
                    if sheet.kind == "worksheet":
                        for _ in book.cells(sheet):
                            cells += 1
    else:
        import warnings

        import openpyxl

        warnings.simplefilter("ignore")
        read_only = mode != "openpyxl normal"
        for f in files:
            wb = openpyxl.load_workbook(f, read_only=read_only, data_only=mode.endswith("data_only"))
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    cells += len(row)
            wb.close()
    elapsed = time.perf_counter() - start
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    peak_mb = peak / 1e6 if sys.platform == "darwin" else peak / 1e3
    print(json.dumps({"mode": mode, "seconds": elapsed, "peak_mb": peak_mb, "cells": cells}))


def main() -> int:
    if len(sys.argv) > 2 and sys.argv[1] == "--child":
        child(sys.argv[2], sys.argv[3:])
        return 0
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--sample", type=int, default=0)
    parser.add_argument("--seed", type=int, default=1)
    parser.add_argument("--modes", default=",".join(MODES))
    args = parser.parse_args()
    files: list[str] = []
    for p in args.paths:
        if p.exists():
            files.extend(str(f) for f in (sorted(p.rglob("*.xlsx")) if p.is_dir() else [p]))
    if not files:
        print("no workbooks found (local corpora are optional)")
        return 0
    if args.sample:
        files = random.Random(args.seed).sample(files, min(args.sample, len(files)))
    size = sum(Path(f).stat().st_size for f in files) / 1e6
    print(f"{len(files)} files, {size:.1f} MB")
    for mode in args.modes.split(","):
        out = subprocess.run(
            [sys.executable, __file__, "--child", mode, *files],
            capture_output=True,
            text=True,
            check=True,
        ).stdout
        r = json.loads(out.strip().splitlines()[-1])
        print(f"  {mode:30} {r['seconds']:7.2f} s  {r['peak_mb']:7.0f} MB peak")
    return 0


if __name__ == "__main__":
    sys.exit(main())
