"""Functions that reshape arrays (ARCHITECTURE.md §7.7).

Each is captured from Excel as dynamic arrays (``dyn_*``), over two sets of values. Common to
them:

- The array argument is taken whole in every formula, a range included; an error given as the
  array (not inside one) is the result, and an empty one is ``#VALUE!``.
- A count or index is truncated; text is read as a number, a logical as 1 or 0, a blank cell as
  0.
- An array given where a count belongs (and in a dynamic array, a range) is taken element by
  element: each element gives the first value of its result, ``#VALUE!`` for a reference. In a
  plain formula a range there is intersected.
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
    CALC,
    MAX_COLS,
    MAX_ROWS,
    VALUE,
    Grid,
    _count,
    _flag,
    _grid,
    _limit,
    _shown,
)
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Ref

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


# TAKE and DROP --------------------------------------------------------------------------------
# Of a range they give a range, which a plain formula's cell intersects as it does any reference
# (=TAKE(Z1:Z4,2) is Z2's value in row 2, #VALUE! in row 3).


def _span(n: int, size: int, take: bool) -> tuple[int, int]:
    """The first and past-the-last line kept of ``size`` when taking or dropping ``n``."""
    if take:
        n = max(-size, min(size, n))
        return (0, n) if n >= 0 else (size + n, size)
    if n >= 0:
        return min(n, size), size
    return 0, max(0, size + n)


def _take_drop(ctx: Any, given: tuple[Any, ...], ranges: bool, take: bool) -> Any:
    core = _take if take else _drop
    lifted = _spread(ctx, core, given, (1, 2), ranges)
    if lifted is not None:
        return lifted
    array = given[0]
    if isinstance(array, ErrorValue):
        return array
    if array is MISSING:
        return VALUE
    counts = []
    for a in given[1:]:
        n = _number(a, ctx, 0.0) if a is not MISSING else None
        if isinstance(n, ErrorValue):
            return n
        counts.append(n)
    counts += [None] * (2 - len(counts))
    if isinstance(array, Ref):
        size = (array.rows, array.columns)
    else:
        array = _grid(array)
        size = (len(array), len(array[0]))
    spans = []
    for n, s in zip(counts, size, strict=True):
        first, last = (0, s) if n is None else _span(n, s, take)
        if last <= first:
            return CALC
        spans.append((first, last))
    (r1, r2), (c1, c2) = spans
    if isinstance(array, Ref):  # a range gives a range, as Excel's TAKE and DROP do
        return Ref(
            array.source, array.sheet, array.r1 + r1, array.c1 + c1,
            array.r1 + r2 - 1, array.c1 + c2 - 1,
        )  # fmt: skip
    return _shown_grid([row[c1:c2] for row in array[r1:r2]])


def _take(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """The first ``rows`` and ``columns`` of ``array``, or the last for a negative count."""
    return _take_drop(ctx, given, ranges, take=True)


def _drop(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """``array`` without its first ``rows`` and ``columns``, or its last for a negative count."""
    return _take_drop(ctx, given, ranges, take=False)


_register("_XLFN.TAKE", _take, "array", "raw", "raw", min_args=2)
_register("_XLFN.DROP", _drop, "array", "raw", "raw", min_args=2)


# EXPAND ---------------------------------------------------------------------------------------
# An empty array argument is one blank here, and an empty pad_with is 0 (captures dyn_droptake,
# empty_args).


def _expand(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """``array`` padded to ``rows`` by ``columns`` with ``pad_with`` (``#N/A`` if left out)."""
    lifted = _spread(ctx, _expand, given, (1, 2), ranges)
    if lifted is not None:
        return lifted
    grid = _whole(BLANK if given[0] is MISSING else given[0])  # empty, it is one blank
    if isinstance(grid, ErrorValue):
        return grid
    size = [len(grid), len(grid[0])]
    for k, a in enumerate(given[1:3]):
        if a is MISSING:
            continue
        n = _number(a, ctx, 0.0)
        if isinstance(n, ErrorValue):
            return n
        if n > (MAX_ROWS, MAX_COLS)[k]:
            return ErrorValue.NUM
        if n < size[k]:
            return VALUE
        size[k] = n
    pad: Any = ErrorValue.NA
    if len(given) > 3:
        pad = given[3]
        if pad is MISSING:
            pad = 0.0
        elif isinstance(pad, Ref):
            if not pad.is_cell() and ranges:
                return VALUE
            pad = scalar(pad, ctx)
        elif type(pad) is tuple:
            return VALUE
    rows, cols = size
    _limit(ctx, rows, cols)
    width = len(grid[0])
    out = [tuple(_shown(v) for v in row) + (pad,) * (cols - width) for row in grid]
    out += [(pad,) * cols] * (rows - len(grid))
    return _shown_grid(out)


_register("_XLFN.EXPAND", _expand, "array", "raw", "raw", "raw", min_args=2)


# VSTACK and HSTACK ----------------------------------------------------------------------------


def _parts(given: Sequence[Any]) -> list[Grid] | ErrorValue:
    parts = []
    for a in given:
        grid = _whole(a)
        if isinstance(grid, ErrorValue):
            return grid
        parts.append(grid)
    return parts


@function("_XLFN.VSTACK", "array", variadic=True, min_args=1, category=C, makes_array=True)
def VSTACK(ctx: Any, *arrays: Any) -> Any:
    """The arrays one below another, padded with ``#N/A`` to the widest."""
    parts = _parts(arrays)
    if isinstance(parts, ErrorValue):
        return parts
    width = max(len(p[0]) for p in parts)
    _limit(ctx, sum(len(p) for p in parts), width)
    na = ErrorValue.NA
    return tuple(
        tuple(_shown(v) for v in row) + (na,) * (width - len(row)) for p in parts for row in p
    )


@function("_XLFN.HSTACK", "array", variadic=True, min_args=1, category=C, makes_array=True)
def HSTACK(ctx: Any, *arrays: Any) -> Any:
    """The arrays side by side, padded with ``#N/A`` to the tallest."""
    parts = _parts(arrays)
    if isinstance(parts, ErrorValue):
        return parts
    height = max(len(p) for p in parts)
    _limit(ctx, height, sum(len(p[0]) for p in parts))
    out = []
    for i in range(height):
        row: list[Any] = []
        for p in parts:
            if i < len(p):
                row.extend(_shown(v) for v in p[i])
            else:
                row.extend([ErrorValue.NA] * len(p[0]))
        out.append(tuple(row))
    return tuple(out)


# TOCOL and TOROW ------------------------------------------------------------------------------


def _flatten(ctx: Any, given: tuple[Any, ...], ranges: bool) -> list[Any] | ErrorValue:
    grid = _whole(given[0])
    if isinstance(grid, ErrorValue):
        return grid
    ignore = _number(given[1], ctx, 0.0) if len(given) > 1 else 0
    if isinstance(ignore, ErrorValue):
        return ignore
    if not 0 <= ignore <= 3:
        return VALUE
    by_col = _flag(given[2], ctx) if len(given) > 2 else False
    if isinstance(by_col, ErrorValue):
        return by_col
    lines = zip(*grid, strict=True) if by_col else grid
    values = [
        _shown(v)
        for line in lines
        for v in line
        if not (ignore & 1 and v is BLANK) and not (ignore & 2 and isinstance(v, ErrorValue))
    ]
    return values or CALC


def _tocol(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """The values of ``array`` in one column, row by row or by column, some kinds ignored."""
    lifted = _spread(ctx, _tocol, given, (1, 2), ranges)
    if lifted is not None:
        return lifted
    values = _flatten(ctx, given, ranges)
    if isinstance(values, ErrorValue):
        return values
    _limit(ctx, len(values), 1)
    return tuple((v,) for v in values)


def _torow(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """The values of ``array`` in one row, row by row or by column, some kinds ignored."""
    lifted = _spread(ctx, _torow, given, (1, 2), ranges)
    if lifted is not None:
        return lifted
    values = _flatten(ctx, given, ranges)
    if isinstance(values, ErrorValue):
        return values
    _limit(ctx, 1, len(values))
    return (tuple(values),)


_register("_XLFN.TOCOL", _tocol, "array", "raw", "raw", min_args=1)
_register("_XLFN.TOROW", _torow, "array", "raw", "raw", min_args=1)


# WRAPROWS and WRAPCOLS ------------------------------------------------------------------------


def _wrap(ctx: Any, given: tuple[Any, ...], ranges: bool, rows: bool) -> Any:
    core = _wraprows if rows else _wrapcols
    lifted = _spread(ctx, core, given, (1, 2), ranges)
    if lifted is not None:
        return lifted
    grid = _whole(given[0])
    if isinstance(grid, ErrorValue):
        return grid
    if len(grid) > 1 and len(grid[0]) > 1:
        return VALUE
    values = [_shown(v) for row in grid for v in row]
    count = _number(given[1], ctx, 0.0)
    if isinstance(count, ErrorValue):
        return count
    if count < 1:
        return ErrorValue.NUM
    pad: Any = given[2] if len(given) > 2 else MISSING
    if pad is MISSING:
        pad = ErrorValue.NA
    elif isinstance(pad, Ref):
        pad = _shown(scalar(pad, ctx))
    width = min(count, len(values))
    height = -(-len(values) // width)
    _limit(ctx, height, width)
    values += [pad] * (width * height - len(values))
    out = tuple(tuple(values[i * width : (i + 1) * width]) for i in range(height))
    return out if rows else tuple(zip(*out, strict=True))


def _wraprows(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """A row or column of values in rows of ``wrap_count``, the last padded."""
    return _wrap(ctx, given, ranges, rows=True)


def _wrapcols(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """A row or column of values in columns of ``wrap_count``, the last padded."""
    return _wrap(ctx, given, ranges, rows=False)


_register("_XLFN.WRAPROWS", _wraprows, "array", "raw", "raw", min_args=2)
_register("_XLFN.WRAPCOLS", _wrapcols, "array", "raw", "raw", min_args=2)
