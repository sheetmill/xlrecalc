"""Functions that make arrays (ARCHITECTURE.md §7.7).

Each returns rows of values that a dynamic-array formula spills. The rules come from the
``dyn_*`` captures, each tried on two sets of values.
"""

from __future__ import annotations

import math
from functools import cmp_to_key
from typing import Any

from xlrecalc._arrays import ArrayLimitError, broadcast, to_array
from xlrecalc._coerce import scalar, to_bool, to_number
from xlrecalc._compare import compare
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Ref

C = "lookup and reference"
L = "math"
VALUE, CALC = ErrorValue.VALUE, ErrorValue("#CALC!")
MAX_ROWS, MAX_COLS = 1_048_576, 16_384
_ERROR_ORDER = {
    "#NULL!": 1, "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7,
    "#GETTING_DATA": 8, "#SPILL!": 9, "#CALC!": 14,
}  # fmt: skip

Grid = tuple[tuple[Any, ...], ...]


def _grid(a: Any) -> Grid:
    """A range, an array or one value as rows of values."""
    if isinstance(a, Ref):
        return to_array(a)
    if isinstance(a, tuple):
        return a
    return ((a,),)


def _shown(v: Any) -> Any:
    """A value as a cell of a result holds it: a blank is 0."""
    return 0.0 if v is BLANK or v is MISSING else v


def _limit(ctx: Any, rows: int, cols: int) -> None:
    limit = getattr(ctx, "max_array_cells", None)
    if limit is not None and rows * cols > limit:
        raise ArrayLimitError(rows * cols, limit)


def _count(v: Any, ctx: Any, default: float) -> float | ErrorValue:
    """A count or a number argument: empty is the default; text is read as a number."""
    if v is MISSING:
        return default
    if isinstance(v, Ref):
        v = scalar(v, ctx)
    if v is BLANK:
        return 0.0
    return to_number(v, ctx)


# SEQUENCE --------------------------------------------------------------------------------------


@function("_XLFN.SEQUENCE", "raw", "raw", "raw", "raw", min_args=1, category=L)
def SEQUENCE(
    ctx: Any, rows: Any, cols: Any = MISSING, start: Any = MISSING, step: Any = MISSING
) -> Any:
    """Numbers from ``start`` by ``step``, row by row, in ``rows`` by ``cols`` cells.

    An array given for an argument is taken element by element, each giving the first number
    of its sequence (capture ``dyn_sequence``).
    """
    given = (rows, cols, start, step)
    if any(type(a) is tuple or (isinstance(a, Ref) and not a.is_cell()) for a in given):
        spread = [to_array(a) if isinstance(a, Ref) and not a.is_cell() else a for a in given]

        def leading(*xs: Any) -> Any:
            v = SEQUENCE(ctx, *xs)
            return v[0][0] if isinstance(v, tuple) else v

        limit = getattr(ctx, "max_array_cells", None)
        return broadcast(leading, spread, [True] * 4, limit)
    args = [_count(rows, ctx, 1.0), _count(cols, ctx, 1.0)]
    args += [_count(start, ctx, 1.0), _count(step, ctx, 1.0)]
    for a in args:
        if isinstance(a, ErrorValue):
            return a
    r, c = math.trunc(args[0]), math.trunc(args[1])  # type: ignore[arg-type]
    first, by = float(args[2]), float(args[3])  # type: ignore[arg-type]
    if r < 0 or c < 0 or r > MAX_ROWS or c > MAX_COLS:
        return VALUE
    if r == 0 or c == 0:
        return CALC
    _limit(ctx, r, c)

    values, v = [], first
    for _ in range(r * c):  # each the one before plus the step, as Excel adds them
        values.append(v if math.isfinite(v) else ErrorValue.NUM)
        v += by
    return tuple(tuple(values[i * c : (i + 1) * c]) for i in range(r))


# MMULT -----------------------------------------------------------------------------------------


@function("MMULT", "array", "array", category=L)
def MMULT(ctx: Any, a: Any, b: Any) -> Any:
    """The matrix product: every element a number, as many columns in ``a`` as rows in ``b``."""
    x, y = _grid(a), _grid(b)
    for m in (x, y):
        for row in m:
            for v in row:
                if isinstance(v, ErrorValue):
                    return v
    for m in (x, y):
        for row in m:
            for v in row:
                if type(v) is not float:
                    return VALUE
    n = len(x[0])
    if n != len(y):
        return VALUE
    _limit(ctx, len(x), len(y[0]))
    out = []
    for row in x:
        cells = []
        for j in range(len(y[0])):
            total = 0.0
            for k in range(n):
                total += row[k] * y[k][j]
            cells.append(total if math.isfinite(total) else ErrorValue.NUM)
        out.append(tuple(cells))
    return tuple(out)


# Order and equality, as SORT, SORTBY and UNIQUE see values ---------------------------------------


def _rank(v: Any) -> int:
    """Numbers, then text, then logicals, then errors; a blank after everything."""
    if v is BLANK or v is MISSING:
        return 4
    if isinstance(v, bool):
        return 2
    if isinstance(v, float):
        return 0
    if isinstance(v, ErrorValue):
        return 3
    return 1


def _order(a: Any, b: Any, ctx: Any) -> int:
    """-1, 0 or 1 for two non-blank values of the same rank, ascending."""
    if isinstance(a, ErrorValue):
        x, y = _ERROR_ORDER.get(str(a), 99), _ERROR_ORDER.get(str(b), 99)
        return (x > y) - (x < y)
    if isinstance(a, str):
        if compare(a, b, "<", ctx.quirks) is True:
            return -1
        return 1 if compare(a, b, ">", ctx.quirks) is True else 0
    return int((a > b) - (a < b))


def _sorted(items: list[Any], keys: list[tuple[list[Any], int]], ctx: Any) -> list[Any]:
    """``items`` in the order of ``keys``, each (one value per item, 1 or -1).

    Equal keys keep their order; a blank comes last in either order.
    """

    def cmp(i: int, j: int) -> int:
        for values, direction in keys:
            a, b = values[i], values[j]
            ra, rb = _rank(a), _rank(b)
            if ra == 4 or rb == 4:
                if ra != rb:
                    return 1 if ra == 4 else -1
                continue
            d = (ra > rb) - (ra < rb) if ra != rb else _order(a, b, ctx)
            if d:
                return d * direction
        return 0

    return [items[i] for i in sorted(range(len(items)), key=cmp_to_key(cmp))]


def _same_key(v: Any) -> tuple[Any, ...]:
    """A key equal for values UNIQUE takes as the same: text without regard to case."""
    rank = _rank(v)
    if rank == 1:
        return (1, v.casefold())
    if rank == 3:
        return (3, str(v))
    if rank == 4:
        return (4,)
    return (rank, v)


def _flag(v: Any, ctx: Any) -> bool | ErrorValue:
    """``by_col`` or ``exactly_once``: empty is FALSE, a number is TRUE unless 0, text #VALUE!."""
    if isinstance(v, Ref):
        v = scalar(v, ctx)
    if v is MISSING or v is BLANK:
        return False
    if isinstance(v, str):
        return VALUE
    return to_bool(v, ctx)


# UNIQUE ----------------------------------------------------------------------------------------


@function("_XLFN.UNIQUE", "array", "raw", "raw", min_args=1, category=C)
def UNIQUE(ctx: Any, array: Any, by_col: Any = MISSING, exactly_once: Any = MISSING) -> Any:
    """Distinct rows (or columns) in the order they first appear; text compares without case."""
    cols, once = _flag(by_col, ctx), _flag(exactly_once, ctx)
    if isinstance(cols, ErrorValue):
        return cols
    if isinstance(once, ErrorValue):
        return once
    grid = _grid(array)
    lines = list(zip(*grid, strict=True)) if cols else list(grid)
    counts: dict[tuple[Any, ...], int] = {}
    first: dict[tuple[Any, ...], tuple[Any, ...]] = {}
    for line in lines:
        key = tuple(_same_key(v) for v in line)
        counts[key] = counts.get(key, 0) + 1
        first.setdefault(key, line)
    kept = [first[k] for k in first if not once or counts[k] == 1]
    if not kept:
        return CALC
    out = tuple(tuple(_shown(v) for v in line) for line in kept)
    return tuple(zip(*out, strict=True)) if cols else out


# SORT and SORTBY -------------------------------------------------------------------------------


def _index_list(v: Any, ctx: Any, default: float) -> list[float] | ErrorValue:
    values = [x for row in _grid(v) for x in row] if isinstance(v, tuple | Ref) else [v]
    out = []
    for x in values:
        n = _count(x, ctx, default)
        if isinstance(n, ErrorValue):
            return n
        out.append(float(math.trunc(n)))
    return out


@function("_XLFN._XLWS.SORT", "array", "raw", "raw", "raw", min_args=1, category=C)
def SORT(
    ctx: Any,
    array: Any,
    sort_index: Any = MISSING,
    sort_order: Any = MISSING,
    by_col: Any = MISSING,
) -> Any:
    """Rows (or columns) in order of the values at ``sort_index``, ascending unless ``-1``."""
    cols = _flag(by_col, ctx)
    if isinstance(cols, ErrorValue):
        return cols
    grid = _grid(array)
    lines = list(zip(*grid, strict=True)) if cols else list(grid)
    indexes = _index_list(sort_index, ctx, 1.0)
    orders = _index_list(sort_order, ctx, 1.0)
    if isinstance(indexes, ErrorValue):
        return indexes
    if isinstance(orders, ErrorValue):
        return orders
    if len(orders) == 1:
        orders = orders * len(indexes)
    if len(orders) != len(indexes):
        return VALUE
    width = len(lines[0])
    keys = []
    for index, order in zip(indexes, orders, strict=True):
        if not 1 <= index <= width or order not in (1.0, -1.0):
            return VALUE
        keys.append(([line[int(index) - 1] for line in lines], int(order)))
    out = tuple(tuple(_shown(v) for v in line) for line in _sorted(lines, keys, ctx))
    return tuple(zip(*out, strict=True)) if cols else out


@function("_XLFN.SORTBY", "array", "array", "raw", variadic=True, min_args=2, cycle=2, category=C)
def SORTBY(ctx: Any, array: Any, *rest: Any) -> Any:
    """Rows (or columns) of ``array`` in order of other arrays, each with its own direction."""
    grid = _grid(array)
    keys: list[tuple[list[Any], int]] = []
    by_col: bool | None = None
    for k in range(0, len(rest), 2):
        by = _grid(rest[k])
        order = _count(rest[k + 1], ctx, 1.0) if k + 1 < len(rest) else 1.0
        if isinstance(order, ErrorValue):
            return order
        if math.trunc(order) not in (1, -1):
            return VALUE
        if len(by) == 1 and len(by[0]) == len(grid[0]) and len(grid[0]) > 1:
            cols, values = True, list(by[0])
        elif len(by[0]) == 1 and len(by) == len(grid):
            cols, values = False, [row[0] for row in by]
        else:
            return VALUE
        if by_col is not None and cols != by_col:
            return VALUE
        by_col = cols
        keys.append((values, math.trunc(order)))
    lines = list(zip(*grid, strict=True)) if by_col else list(grid)
    out = tuple(tuple(_shown(v) for v in line) for line in _sorted(lines, keys, ctx))
    return tuple(zip(*out, strict=True)) if by_col else out
