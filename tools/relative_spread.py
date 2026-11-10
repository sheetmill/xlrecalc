"""How far the engine is from Excel on a capture, per function, in relative terms.

    uv run python tools/relative_spread.py SUITE [--import MODULE]

For each function a suite's rows call outermost: the rows compared, those equal bit for bit,
those where one side is exactly 0 and the other is not (a relative difference means nothing
there), and for the rest the count in each power of ten of |engine - Excel| / |Excel|, with the
largest. ``--import`` registers a module of functions first (one not yet imported by the
package). The measure for setting a function's declared relative tolerance (ARCHITECTURE.md §8.1).
"""

from __future__ import annotations

import importlib
import math
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

_OUTER = re.compile(r"^(?:_xlfn\.)?([A-Z][A-Z0-9.]*)\(")


def spread(suite_name: str) -> dict[str, dict[str, float]]:
    """Per function: ``rows``, ``exact``, ``zero`` (one side 0), ``1e-k`` bins and ``max``."""
    from tests.captures import load
    from tests.test_captures import run, same

    suite = load(suite_name)
    engine, sheet = run(suite)
    out: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for k, row in enumerate(suite.rows, 1):
        got, want = engine.get(sheet, k - 1, 0), row.expected
        found = _OUTER.match(row.formula)
        if not found or not isinstance(got, float) or not isinstance(want, float):
            continue
        counts = out[found[1]]
        counts["rows"] += 1
        if same(got, want):
            counts["exact"] += 1
        elif got == 0 or want == 0:
            counts["zero"] += 1
        else:
            rel = abs(got - want) / abs(want)
            power = min(0, math.ceil(math.log10(rel)))
            counts[f"1e{power}"] += 1
            counts["max"] = max(counts["max"], rel)
    return out


def main() -> None:
    args = sys.argv[1:]
    if "--import" in args:
        at = args.index("--import")
        importlib.import_module(args[at + 1])
        del args[at : at + 2]
    for name, counts in sorted(spread(args[0]).items()):
        bins = sorted((k for k in counts if k.startswith("1e")), key=lambda k: int(k[2:]))
        shown = " ".join(f"{k}:{int(counts[k])}" for k in bins)
        print(
            f"{name:14} rows {int(counts['rows']):4} exact {int(counts['exact']):4} "
            f"zero {int(counts['zero']):2} max {counts['max']:.1e}  {shown}"
        )


if __name__ == "__main__":
    main()
