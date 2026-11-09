"""Functions that make arrays from data (ARCHITECTURE.md §7.7).

Each function's rules come from the capture its docstring names, tried on two sets of values.
"""

from __future__ import annotations

import bisect
from typing import Any

from xlrecalc._arrays import to_array
from xlrecalc._formula import MISSING
from xlrecalc._functions._dynamic import VALUE
from xlrecalc._functions._registry import function
from xlrecalc._values import ErrorValue, Ref

S = "statistical"
T = "text"
M = "math"


def _elements(a: Any) -> list[Any]:
    """The values of a range, an array or one value."""
    if isinstance(a, Ref):
        return [v for row in to_array(a) for v in row]
    if isinstance(a, tuple):
        return [v for row in a for v in row]
    return [a]


def _numbers(a: Any) -> list[float] | ErrorValue:
    """The numbers among the values, in order; text, logicals and blanks are left out."""
    out = []
    for v in _elements(a):
        if isinstance(v, ErrorValue):
            return v
        if type(v) is float:
            out.append(v)
    return out


# FREQUENCY -------------------------------------------------------------------------------------


@function("FREQUENCY", "array", "array", category=S, makes_array=True)
def FREQUENCY(ctx: Any, data: Any, bins: Any) -> Any:
    """How many numbers fall in each bin, in the bins' order, and above the largest in a last row.

    A number counts in the smallest bin at or above it; equal bins take their count in the
    first; without a number among the bins there is one bin, 0 (capture ``dyn_frequency``).
    """
    if data is MISSING or bins is MISSING:  # an empty argument (capture empty_args)
        return VALUE
    values = _numbers(data)
    if isinstance(values, ErrorValue):
        return values
    edges = _numbers(bins)
    if isinstance(edges, ErrorValue):
        return edges
    if not edges:
        edges = [0.0]
    order = sorted(range(len(edges)), key=lambda i: edges[i])
    ordered = [edges[i] for i in order]
    counts = [0.0] * (len(edges) + 1)
    for v in values:
        k = bisect.bisect_left(ordered, v)
        counts[order[k] if k < len(ordered) else len(edges)] += 1.0
    return tuple((c,) for c in counts)


# MODE.MULT -------------------------------------------------------------------------------------


@function("_XLFN.MODE.MULT", "array", min_args=1, variadic=True, category=S, makes_array=True)
def MODE_MULT(ctx: Any, *args: Any) -> Any:
    """Every number that occurs most often, at least twice, in the order each first appears.

    Logicals and text are left out, given directly too (capture ``dyn_frequency``); none
    repeating is ``#N/A``.
    """
    counts: dict[float, int] = {}
    for a in args:
        if a is MISSING:
            return VALUE
        values = _numbers(a)
        if isinstance(values, ErrorValue):
            return values
        for v in values:
            counts[v] = counts.get(v, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        return ErrorValue.NA
    return tuple((v,) for v, n in counts.items() if n == best)
