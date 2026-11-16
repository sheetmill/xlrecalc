"""Load time on three kinds of workbook, best of N, for comparing two installs (xlrecalc, xlio).

    uv run python bench/load_parity.py --small LIST --mid PATH [--big PATH] [--repeat 3] [--json]

Runs in the package's own environment, from the repository root, so it measures whatever
xlrecalc and xlio that environment holds (both versions are printed). Each measurement is
``xlrecalc.load(path, recalc_on_load="never")``: reading the file into an engine, nothing
evaluated. ``--small`` is a text file of workbook paths, one a line (many small workbooks, timed
together: the fixed cost of a file shows there); ``--mid`` one mid-sized workbook; ``--big`` the
1,000,000-row benchmark file (``bench/data/bench_1000000.xlsx``, from ``bench/gen.py``; left out
when the file is absent). Each is timed ``--repeat`` times and the best kept. Third-party
workbooks are never committed: the lists point at local copies.
"""

from __future__ import annotations

import argparse
import json
import time
from importlib import metadata
from pathlib import Path

import xlrecalc

ROOT = Path(__file__).resolve().parents[1]
BIG = ROOT / "bench" / "data" / "bench_1000000.xlsx"


def best(paths: list[Path], repeat: int) -> float:
    """The shortest of ``repeat`` timings of loading every path in turn, in seconds."""
    times = []
    for _ in range(repeat):
        start = time.perf_counter()
        for path in paths:
            xlrecalc.load(path, recalc_on_load="never")
        times.append(time.perf_counter() - start)
    return min(times)


def main() -> None:
    """Time the three loads and print the figures."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--small", type=Path, required=True, help="file listing small workbooks")
    parser.add_argument("--mid", type=Path, required=True, help="one mid-sized workbook")
    parser.add_argument("--big", type=Path, default=BIG, help="the 1,000,000-row workbook")
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--json", action="store_true", help="one JSON line instead of text")
    args = parser.parse_args()
    small = [Path(line) for line in args.small.read_text().splitlines() if line.strip()]
    result: dict[str, object] = {
        "xlrecalc": metadata.version("xlrecalc"),
        "xlio": metadata.version("xlio") if _installed("xlio") else None,
        "repeat": args.repeat,
        "small_count": len(small),
        "small_s": best(small, args.repeat),
        "mid": args.mid.name,
        "mid_s": best([args.mid], args.repeat),
    }
    if args.big.exists():
        result["big_s"] = best([args.big], args.repeat)
    if args.json:
        print(json.dumps(result))
        return
    print(f"xlrecalc {result['xlrecalc']}, xlio {result['xlio']}, best of {args.repeat}")
    print(f"{len(small)} small workbooks: {result['small_s']:.3f} s")
    print(f"mid ({args.mid.name}): {result['mid_s']:.3f} s")
    if "big_s" in result:
        print(f"1M rows: {result['big_s']:.2f} s")


def _installed(name: str) -> bool:
    try:
        metadata.version(name)
    except metadata.PackageNotFoundError:
        return False
    return True


if __name__ == "__main__":
    main()
