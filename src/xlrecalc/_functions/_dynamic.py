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
from xlrecalc._numtext import DBL_MIN
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


def _sequence(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """SEQUENCE over its four arguments.

    An array (and with ``ranges``, in an array context, a range) is taken element by element,
    each giving the first number of its sequence; in a plain formula a range is intersected
    instead (captures ``dyn_sequence``, ``empty_args``).
    """
    if any(type(a) is tuple or (ranges and isinstance(a, Ref) and not a.is_cell()) for a in given):
        spread = [
            to_array(a) if ranges and isinstance(a, Ref) and not a.is_cell() else a for a in given
        ]
        lifted = [type(a) is tuple for a in spread]

        def leading(*xs: Any) -> Any:
            v = _sequence(ctx, xs, ranges)
            return v[0][0] if isinstance(v, tuple) else v

        limit = getattr(ctx, "max_array_cells", None)
        return broadcast(leading, spread, lifted, limit)
    rows, cols, start, step = given
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


def _sequence_array(
    ctx: Any, rows: Any, cols: Any = MISSING, start: Any = MISSING, step: Any = MISSING
) -> Any:
    return _sequence(ctx, (rows, cols, start, step), True)


@function(
    "_XLFN.SEQUENCE",
    "raw",
    "raw",
    "raw",
    "raw",
    min_args=1,
    category=L,
    in_array=_sequence_array,
    makes_array=True,
)
def SEQUENCE(
    ctx: Any, rows: Any, cols: Any = MISSING, start: Any = MISSING, step: Any = MISSING
) -> Any:
    """Numbers from ``start`` by ``step``, row by row, in ``rows`` by ``cols`` cells."""
    return _sequence(ctx, (rows, cols, start, step), False)


# MMULT -----------------------------------------------------------------------------------------


def _numbers(a: Any) -> Grid | ErrorValue:
    """A matrix argument: every element a number, the first error otherwise, else ``#VALUE!``."""
    m = _grid(a)
    for row in m:
        for v in row:
            if isinstance(v, ErrorValue):
                return v
    for row in m:
        for v in row:
            if type(v) is not float:
                return VALUE
    return m


@function("MMULT", "array", "array", category=L, makes_array=True)
def MMULT(ctx: Any, a: Any, b: Any) -> Any:
    """The matrix product: every element a number, as many columns in ``a`` as rows in ``b``."""
    x, y = _grid(a), _grid(b)
    for m in (x, y):  # an error in either matrix before a value of another kind in either
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


# Square matrices ---------------------------------------------------------------------------------


def _square(a: Any) -> list[list[float]] | ErrorValue:
    m = _numbers(a)
    if isinstance(m, ErrorValue):
        return m
    if len(m) != len(m[0]):
        return VALUE
    return [list(row) for row in m]


def _pivot(m: list[list[float]], k: int) -> int:
    """The row from ``k`` down with the largest magnitude in column ``k``, the first if tied."""
    best = k
    for i in range(k + 1, len(m)):
        if abs(m[i][k]) > abs(m[best][k]):
            best = i
    return best


def _lu(m: list[list[float]]) -> tuple[list[int], int] | None:
    """Factor ``m`` in place by elimination with partial pivoting.

    As Excel's ``MDETERM`` and ``MINVERSE`` do (capture ``matrix_family``): ``U`` on and above the
    diagonal, the multipliers below it. Returns the row order and the number of swaps, or ``None``
    for a zero pivot.
    """
    n = len(m)
    order, swaps = list(range(n)), 0
    for k in range(n):
        p = _pivot(m, k)
        if m[p][k] == 0.0:
            return None
        if p != k:
            m[k], m[p] = m[p], m[k]
            order[k], order[p] = order[p], order[k]
            swaps += 1
        for i in range(k + 1, n):
            f = m[i][k] / m[k][k]
            m[i][k] = f
            for j in range(k + 1, n):
                m[i][j] -= f * m[k][j]
    return order, swaps


def _normal(v: float) -> float | ErrorValue:
    """A result as Excel holds it: no subnormal numbers, no infinity."""
    if not math.isfinite(v):
        return ErrorValue.NUM
    return 0.0 if abs(v) < DBL_MIN else v


@function("MDETERM", "array", category=L)
def MDETERM(ctx: Any, a: Any) -> Any:
    """The determinant: the pivots of the factorization multiplied in order, signed by swaps."""
    m = _square(a)
    if isinstance(m, ErrorValue):
        return m
    factored = _lu(m)
    if factored is None:
        return 0.0
    det = 1.0
    for k in range(len(m)):
        det *= m[k][k]
    return _normal(-det if factored[1] % 2 else det)


@function("MINVERSE", "array", category=L, makes_array=True)
def MINVERSE(ctx: Any, a: Any) -> Any:
    """The inverse: each column of the identity solved through the factorization.

    Forward and back substitution with sums in increasing order and a division by each pivot;
    singular is ``#NUM!``.
    """
    m = _square(a)
    if isinstance(m, ErrorValue):
        return m
    n = len(m)
    _limit(ctx, n, n)
    factored = _lu(m)
    if factored is None:
        return ErrorValue.NUM
    order = factored[0]
    columns = []
    for c in range(n):
        y = [1.0 if order[i] == c else 0.0 for i in range(n)]
        for i in range(n):
            t = y[i]
            for j in range(i):
                t -= m[i][j] * y[j]
            y[i] = t
        x = [0.0] * n
        for i in range(n - 1, -1, -1):
            t = y[i]
            for j in range(i + 1, n):
                t -= m[i][j] * x[j]
            x[i] = t / m[i][i]
        columns.append(x)
    out = tuple(tuple(_normal(columns[c][r]) for c in range(n)) for r in range(n))
    for row in out:
        for v in row:
            if isinstance(v, ErrorValue):
                return v
    return out


def _munit(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    (size,) = given
    if type(size) is tuple or (ranges and isinstance(size, Ref) and not size.is_cell()):
        spread = [to_array(size) if isinstance(size, Ref) else size]

        def leading(x: Any) -> Any:
            v = _munit(ctx, (x,), ranges)
            return v[0][0] if isinstance(v, tuple) else v

        return broadcast(leading, spread, [True], getattr(ctx, "max_array_cells", None))
    n = _count(size, ctx, 0.0)
    if isinstance(n, ErrorValue):
        return n
    k = math.trunc(n)
    if k < 1:
        return VALUE
    _limit(ctx, k, k)  # a size past the limit keeps the cell's value, reported
    return tuple(tuple(1.0 if j == i else 0.0 for j in range(k)) for i in range(k))


def _munit_array(ctx: Any, size: Any) -> Any:
    return _munit(ctx, (size,), True)


@function("_XLFN.MUNIT", "raw", category=L, in_array=_munit_array, makes_array=True)
def MUNIT(ctx: Any, size: Any) -> Any:
    """The identity matrix of ``size`` rows and columns."""
    return _munit(ctx, (size,), False)


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


@function("_XLFN.UNIQUE", "array", "raw", "raw", min_args=1, category=C, makes_array=True)
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


@function(
    "_XLFN._XLWS.SORT", "array", "raw", "raw", "raw", min_args=1, category=C, makes_array=True
)
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


@function(
    "_XLFN.SORTBY",
    "array",
    "array",
    "raw",
    variadic=True,
    min_args=2,
    cycle=2,
    category=C,
    makes_array=True,
)
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


# FILTER ----------------------------------------------------------------------------------------


@function("_XLFN._XLWS.FILTER", "array", "array", "array", min_args=2, category=C, makes_array=True)
def FILTER(ctx: Any, array: Any, include: Any, if_empty: Any = MISSING) -> Any:
    """The rows (or columns) whose ``include`` is TRUE or a number other than 0."""
    if include is MISSING:
        return VALUE
    grid, test = _grid(array), _grid(include)
    if len(test[0]) == 1 and len(test) == len(grid):
        flags, cols = [row[0] for row in test], False
    elif len(test) == 1 and len(test[0]) == len(grid[0]):
        flags, cols = list(test[0]), True
    else:
        return VALUE
    keep = []
    for v in flags:
        if isinstance(v, ErrorValue):
            return v
        if isinstance(v, str):
            return VALUE
        keep.append(v is not BLANK and v is not MISSING and bool(v))
    lines = list(zip(*grid, strict=True)) if cols else list(grid)
    kept = [line for line, k in zip(lines, keep, strict=True) if k]
    if not kept:
        if if_empty is MISSING:
            return CALC
        return if_empty if not isinstance(if_empty, Ref) else to_array(if_empty)
    out = tuple(tuple(_shown(v) for v in line) for line in kept)
    return tuple(zip(*out, strict=True)) if cols else out
