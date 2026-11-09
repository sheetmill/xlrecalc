"""Functions that reshape arrays (ARCHITECTURE.md §7.7).

Each is captured from Excel as dynamic arrays (``dyn_*``), over two sets of values. Common to
them:

- The array argument is taken whole in every formula, a range included; an error given as the
  array (not inside one) is the result, and an empty one is ``#VALUE!``.
- A count or index is truncated; text is read as a number, a logical as 1 or 0, a blank cell as
  0.
- A blank in a result shows 0.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Sequence
from typing import Any

from xlrecalc._arrays import broadcast, to_array
from xlrecalc._coerce import scalar
from xlrecalc._formula import MISSING
from xlrecalc._functions._dynamic import (
    VALUE,
    Grid,
    _count,
    _grid,
    _limit,
    _shown,
)
from xlrecalc._functions._registry import function
from xlrecalc._values import ErrorValue, Ref

C = "lookup and reference"

Core = Callable[[Any, tuple[Any, ...], bool], Any]


def _spread(
    ctx: Any, core: Core, given: tuple[Any, ...], liftable: Sequence[int], ranges: bool
) -> Any:
    """``core`` element by element over the arrays at ``liftable``, or ``None`` if there are none.

    Each element gives the first value of its result, ``#VALUE!`` for a reference (an array of
    references). With ``ranges`` (an array context) a range there is lifted too.
    """

    def lifts(a: Any) -> bool:
        return type(a) is tuple or (ranges and isinstance(a, Ref) and not a.is_cell())

    if not any(lifts(given[i]) for i in liftable if i < len(given)):
        return None
    spread = [
        to_array(a) if i in liftable and isinstance(a, Ref) and lifts(a) else a
        for i, a in enumerate(given)
    ]
    lifted = [i in liftable and type(a) is tuple for i, a in enumerate(spread)]

    def leading(*xs: Any) -> Any:
        v = core(ctx, xs, ranges)
        if isinstance(v, Ref):
            return VALUE
        return v[0][0] if isinstance(v, tuple) else v

    return broadcast(leading, spread, lifted, getattr(ctx, "max_array_cells", None))


def _number(v: Any, ctx: Any, default: float) -> int | ErrorValue:
    """A count, truncated."""
    n = _count(v, ctx, default)
    return n if isinstance(n, ErrorValue) else math.trunc(n)


def _whole(a: Any) -> Grid | ErrorValue:
    """The array argument as rows of values; an error given as the array is the result.

    An empty array argument is ``#VALUE!`` (capture ``empty_args``).
    """
    if isinstance(a, ErrorValue):
        return a
    if a is MISSING:
        return VALUE
    return _grid(a)


def _shown_grid(rows: Sequence[Sequence[Any]]) -> Grid:
    return tuple(tuple(_shown(v) for v in row) for row in rows)


def _register(name: str, core: Core, *kinds: str, **options: Any) -> None:
    """Register ``core`` under ``name``: plain, and in an array context with ranges lifted."""

    def plain(ctx: Any, *given: Any) -> Any:
        return core(ctx, given, False)

    def in_array(ctx: Any, *given: Any) -> Any:
        return core(ctx, given, True)

    plain.__doc__ = core.__doc__
    plain.__name__ = name.removeprefix("_XLFN.")
    function(name, *kinds, category=C, in_array=in_array, makes_array=True, **options)(plain)


# CHOOSECOLS and CHOOSEROWS --------------------------------------------------------------------
# The indexes are lists: a row or column of them picks several, a two-dimensional one is #VALUE!.


def _indexes(given: Sequence[Any], ctx: Any, ranges: bool) -> list[int] | ErrorValue:
    """The indexes CHOOSECOLS and CHOOSEROWS are given: numbers, or rows or columns of them."""
    out: list[int] = []
    for a in given:
        if a is MISSING:
            return VALUE
        if isinstance(a, Ref) and not a.is_cell():
            a = to_array(a) if ranges else scalar(a, ctx)
        values = [a]
        if type(a) is tuple:
            if len(a) > 1 and len(a[0]) > 1:
                return VALUE
            values = [v for row in a for v in row]
        for v in values:
            n = _number(v, ctx, 0.0)
            if isinstance(n, ErrorValue):
                return n
            out.append(n)
    return out


def _choose(ctx: Any, given: tuple[Any, ...], ranges: bool, rows: bool) -> Any:
    grid = _whole(given[0])
    if isinstance(grid, ErrorValue):
        return grid
    indexes = _indexes(given[1:], ctx, ranges)
    if isinstance(indexes, ErrorValue):
        return indexes
    lines = list(grid) if rows else list(zip(*grid, strict=True))
    picked = []
    for i in indexes:
        if not 1 <= abs(i) <= len(lines):
            return VALUE
        picked.append(lines[i - 1 if i > 0 else i])
    _limit(ctx, len(picked), len(lines[0]))
    out = _shown_grid(picked)
    return out if rows else tuple(zip(*out, strict=True))


def _choosecols(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """The columns of ``array`` at the indexes given, negative ones from the end."""
    return _choose(ctx, given, ranges, rows=False)


def _chooserows(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """The rows of ``array`` at the indexes given, negative ones from the end."""
    return _choose(ctx, given, ranges, rows=True)


_register("_XLFN.CHOOSECOLS", _choosecols, "array", "raw", variadic=True, min_args=2)
_register("_XLFN.CHOOSEROWS", _chooserows, "array", "raw", variadic=True, min_args=2)
