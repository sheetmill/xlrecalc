"""Lookup and reference functions: VLOOKUP, HLOOKUP, MATCH, INDEX, CHOOSE, ROW(S), COLUMN(S).

Exact lookups compare as ``lookup_equal`` does (kinds first, text without case, numbers
exactly) and accept wildcards in text (``*``, ``?``, ``~``). Approximate lookups search a sorted
range by bisection over the values of the lookup value's kind, as Excel's does.
"""

from __future__ import annotations

import math
import re
from typing import Any

from xlrecalc._arrays import check_area
from xlrecalc._coerce import Context, scalar
from xlrecalc._compare import compare, lookup_equal, wildcard
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._store import MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Ref

NA, VALUE, REF = ErrorValue.NA, ErrorValue.VALUE, ErrorValue.REF
L = "lookup and reference"
_WILD = re.compile(r"[*?~]")


def _exact(lookup: Any, values: list[Any]) -> int | None:
    if isinstance(lookup, str) and _WILD.search(lookup):
        pattern = wildcard(lookup)
        for i, v in enumerate(values):
            if isinstance(v, str) and pattern.fullmatch(v):
                return i
        return None
    for i, v in enumerate(values):
        if lookup_equal(lookup, v):
            return i
    return None


def _kind(v: Any) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    if isinstance(v, float):
        return 0
    return -1


def _approximate(lookup: Any, values: list[Any], descending: bool, ctx: Context) -> int | None:
    """Last position whose value is <= lookup (>= when descending), by bisection."""
    kind = _kind(lookup)
    candidates = [i for i, v in enumerate(values) if _kind(v) == kind]
    lo, hi, found = 0, len(candidates) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        v = values[candidates[mid]]
        order = compare(v, lookup, ">=" if descending else "<=", ctx.quirks)
        if order is True:
            found = candidates[mid]
            lo = mid + 1
        else:
            hi = mid - 1
    return found


class _Grid:
    """An array seen as a range: ``rows``, ``columns`` and ``value(row, col)`` (§7.7)."""

    def __init__(self, array: tuple[tuple[Any, ...], ...]) -> None:
        self.array = array
        self.rows = len(array)
        self.columns = len(array[0]) if array else 0

    def value(self, row: int, col: int) -> Any:
        return self.array[row][col]


def _line(ref: Ref | _Grid, horizontal: bool) -> list[Any]:
    if horizontal:
        return [ref.value(0, c) for c in range(ref.columns)]
    return [ref.value(r, 0) for r in range(ref.rows)]


def _lookup_value(lookup: Any, ctx: Context) -> Any:
    v = scalar(lookup, ctx)
    return 0.0 if v is BLANK else v


@function("MATCH", "value", "raw", "number", min_args=2, errors=True, category=L)
def MATCH(ctx: Context, lookup: Any, ref: Any, kind: Any = 1.0) -> float | ErrorValue:
    """Position of a value in a row or column: type 0 exact, 1 sorted ascending, -1 descending."""
    v = _lookup_value(lookup, ctx)
    if isinstance(v, ErrorValue):
        return v
    if isinstance(ref, ErrorValue):
        return ref
    if isinstance(kind, ErrorValue):
        return kind
    if isinstance(ref, tuple):
        values = [x for row in ref for x in row]
    elif isinstance(ref, Ref) and (ref.rows == 1 or ref.columns == 1):
        values = _line(ref, ref.rows == 1)
    else:
        return NA
    k = 0 if kind == 0 else (1 if kind > 0 else -1)
    i = _exact(v, values) if k == 0 else _approximate(v, values, k < 0, ctx)
    return NA if i is None else float(i + 1)


def _table_lookup(
    ctx: Context, lookup: Any, table: Any, index: Any, approximate: Any, horizontal: bool
) -> Any:
    v = _lookup_value(lookup, ctx)
    if isinstance(v, ErrorValue):
        return v
    if isinstance(table, ErrorValue):
        return table
    if isinstance(index, ErrorValue):
        return index
    if isinstance(table, tuple):
        table = _Grid(table)
    elif not isinstance(table, Ref):
        return NA  # a value, not a range (capture arrays)
    n = math.trunc(index)
    width = table.rows if horizontal else table.columns
    if n < 1:
        return VALUE
    if n > width:
        return REF
    approx = True if approximate is MISSING or approximate is None else approximate
    if approx is not True and approx is not False:
        approx = scalar(approx, ctx)
        if isinstance(approx, ErrorValue):
            return approx
        approx = bool(approx) if not isinstance(approx, str) else approx.upper() == "TRUE"
    keys = _line(table, horizontal)
    i = _approximate(v, keys, False, ctx) if approx else _exact(v, keys)
    if i is None:
        return NA
    result = table.value(n - 1, i) if horizontal else table.value(i, n - 1)
    return 0.0 if result is BLANK else result


@function("VLOOKUP", "key", "raw", "number", "raw", min_args=3, errors=True, category=L)
def VLOOKUP(ctx: Context, lookup: Any, table: Any, col: float, approximate: Any = None) -> Any:
    """Find in the first column, return from column ``col``; sorted unless the last is FALSE.

    An omitted fourth argument (``VLOOKUP(x,t,2,)``) is FALSE; an absent one is TRUE.
    """
    if approximate is MISSING:
        approximate = False
    return _table_lookup(ctx, lookup, table, col, approximate, horizontal=False)


@function("HLOOKUP", "key", "raw", "number", "raw", min_args=3, errors=True, category=L)
def HLOOKUP(ctx: Context, lookup: Any, table: Any, row: float, approximate: Any = None) -> Any:
    """Find in the first row, return from row ``row``."""
    if approximate is MISSING:
        approximate = False
    return _table_lookup(ctx, lookup, table, row, approximate, horizontal=True)


def _whole(ref: Any, r: int, c: int) -> Any:
    """``INDEX`` with a 0 (or omitted) index over a range of several rows and columns."""
    if isinstance(ref, Ref):
        rows, cols = ref.rows, ref.columns
    elif isinstance(ref, tuple):
        rows, cols = len(ref), len(ref[0])
    else:
        return None
    if r == 0 and c == 0:
        return ref
    if rows == 1 or cols == 1 or (r != 0 and c != 0) or r < 0 or c < 0:
        return None
    if (r == 0 and c > cols) or (c == 0 and r > rows):
        return REF
    if isinstance(ref, Ref):
        if r == 0:
            return Ref(ref.source, ref.sheet, ref.r1, ref.c1 + c - 1, ref.r2, ref.c1 + c - 1)
        return Ref(ref.source, ref.sheet, ref.r1 + r - 1, ref.c1, ref.r1 + r - 1, ref.c2)
    if r == 0:
        return tuple((line[c - 1],) for line in ref)
    return (ref[r - 1],)


@function("INDEX", "array", "number", "number", min_args=2, category=L)
def INDEX(ctx: Context, ref: Any, row: float, col: float = 0.0) -> Any:
    """The value at ``row``, ``col`` (1-based) of a range; one index for a single row or column.

    A 0 index gives a whole column (``INDEX(A1:B5,0,1)``) or row, as a reference into a
    range (which a legacy formula's cell then intersects) or an array from an array; as does a
    row index alone into a range of several columns. ``INDEX`` as a reference
    (``A1:INDEX(...)``) comes with dynamic references.
    """
    if isinstance(ref, ErrorValue):  # INDEX(TRANSPOSE(A1:C2),...) in a legacy formula
        return ref
    r, c = math.trunc(row), math.trunc(col)
    whole = _whole(ref, r, c)
    if whole is not None:
        return whole
    if isinstance(ref, tuple):
        rows, cols = len(ref), len(ref[0])

        def get(i: int, j: int) -> Any:
            return ref[i][j]

    elif isinstance(ref, Ref):
        rows, cols = ref.rows, ref.columns
        get = ref.value  # type: ignore[assignment]
    else:
        v = scalar(ref, ctx)
        return v if (r, c) in ((1, 0), (1, 1), (0, 1)) else REF
    if rows == 1 and c == 0:
        r, c = 1, r
    elif c == 0:
        c = 1
    if r < 1 or c < 1:
        return VALUE
    if r > rows or c > cols:
        return REF
    if isinstance(ref, Ref):  # a reference, so A1:INDEX(...) and CELL(...,INDEX(...)) work
        return Ref(
            ref.source, ref.sheet, ref.r1 + r - 1, ref.c1 + c - 1, ref.r1 + r - 1, ref.c1 + c - 1
        )
    v = get(r - 1, c - 1)
    return 0.0 if v is BLANK else v


@function("CHOOSE", "number", "branch", min_args=2, variadic=True, category=L)
def CHOOSE(ctx: Context, index: float, *values: Any) -> Any:
    """The ``index``-th value."""
    i = math.trunc(index)
    if not 1 <= i <= len(values):
        return VALUE
    v = values[i - 1]
    return 0.0 if v is MISSING else v


@function("ROWS", "array", category=L)
def ROWS(ctx: Context, ref: Any) -> float | ErrorValue:
    """Rows of a reference or array."""
    if isinstance(ref, Ref):
        return float(ref.rows)
    if isinstance(ref, tuple):
        return float(len(ref))
    return 1.0 if not isinstance(scalar(ref, ctx), ErrorValue) else scalar(ref, ctx)  # type: ignore[return-value]


@function("COLUMNS", "array", category=L)
def COLUMNS(ctx: Context, ref: Any) -> float | ErrorValue:
    """Columns of a reference or array."""
    if isinstance(ref, Ref):
        return float(ref.columns)
    if isinstance(ref, tuple):
        return float(len(ref[0]))
    return 1.0 if not isinstance(scalar(ref, ctx), ErrorValue) else scalar(ref, ctx)  # type: ignore[return-value]


def _not_one_area(ref: Any) -> ErrorValue:
    """``ROW`` of what is not one area: a union is ``#REF!`` (capture ``array_refs``)."""
    return ref if isinstance(ref, ErrorValue) and ref is not VALUE else REF


def _rows_array(ctx: Context, ref: Any = MISSING) -> Any:
    """``ROW`` in an array context: one column, a row number for each row of a range (§7.7)."""
    if not isinstance(ref, Ref) or ref.rows == 1:
        return ROW(ctx, ref)
    check_area(ref.rows, 1, getattr(ctx, "max_array_cells", None))
    return tuple((float(r + 1),) for r in range(ref.r1, ref.r2 + 1))


def _columns_array(ctx: Context, ref: Any = MISSING) -> Any:
    """``COLUMN`` in an array context: one row, a column number for each column of a range."""
    if not isinstance(ref, Ref) or ref.columns == 1:
        return COLUMN(ctx, ref)
    check_area(1, ref.columns, getattr(ctx, "max_array_cells", None))
    return (tuple(float(c + 1) for c in range(ref.c1, ref.c2 + 1)),)


@function("ROW", "raw", min_args=0, category=L, in_array=_rows_array)
def ROW(ctx: Context, ref: Any = MISSING) -> float | ErrorValue:
    """Row number of a reference's first cell, or of the formula's cell."""
    if ref is MISSING:
        return float(ctx.row + 1)
    return float(ref.r1 + 1) if isinstance(ref, Ref) else _not_one_area(ref)


@function("COLUMN", "raw", min_args=0, category=L, in_array=_columns_array)
def COLUMN(ctx: Context, ref: Any = MISSING) -> float | ErrorValue:
    """Column number of a reference's first cell, or of the formula's cell."""
    if ref is MISSING:
        return float(ctx.col + 1)
    return float(ref.c1 + 1) if isinstance(ref, Ref) else _not_one_area(ref)


@function("_XLFN.ANCHORARRAY", "raw", category=L)
def ANCHORARRAY(ctx: Any, ref: Any) -> Any:
    """``A1#``: the range the dynamic array at A1 spills into (§7.7)."""
    return ctx.spill_range(ref)


@function("_XLFN.SINGLE", "raw", category=L)
def SINGLE(ctx: Any, value: Any) -> Any:
    """``@x``: the implicit intersection of a reference, the first element of an array."""
    return ctx.ops.single(value)


def _n(v: Any) -> Any:
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, float | ErrorValue):
        return v
    return 0.0  # text, blank


@function("N", "raw", category=L)
def N(ctx: Context, value: Any) -> Any:
    """A value as a number: logicals 1 or 0, text and blanks 0, errors kept.

    A range gives its top-left cell, not an implicit intersection; an array is taken element by
    element (capture ``n_lookup``).
    """
    if isinstance(value, Ref):
        return _n(value.value(0, 0))
    if isinstance(value, tuple):
        return tuple(tuple(_n(x) for x in row) for row in value)
    return _n(value)


def _grid(a: Any) -> Ref | _Grid:
    if isinstance(a, Ref):
        return a
    if isinstance(a, tuple):
        return _Grid(a)
    return _Grid(((a,),))


@function("LOOKUP", "value", "array", "array", min_args=2, category=L)
def LOOKUP(ctx: Context, value: Any, lookup: Any, result: Any = MISSING) -> Any:
    """The approximate match of ``value`` in a sorted row or column, and the matching result.

    With ``result``, the position found is taken along it, a range extended past its end in
    its own direction (a single cell to the right). Without it, an array wider than tall is
    searched in its first row and answered from its last; any other in its first column and
    answered from its last (capture ``n_lookup``).
    """
    v = _lookup_value(value, ctx)
    if isinstance(v, ErrorValue):
        return v
    if isinstance(lookup, ErrorValue):
        return lookup
    if result is not MISSING and isinstance(result, ErrorValue):
        return result
    grid = _grid(lookup)
    if result is MISSING:
        wide = grid.columns > grid.rows
        keys = _line(grid, wide)
        i = _approximate(v, keys, False, ctx)
        if i is None:
            return NA
        found = grid.value(grid.rows - 1, i) if wide else grid.value(i, grid.columns - 1)
        return 0.0 if found is BLANK else found
    keys = _line(grid, grid.rows == 1)
    i = _approximate(v, keys, False, ctx)
    if i is None:
        return NA
    if isinstance(result, Ref):
        r, c = (result.r1, result.c1 + i) if result.rows == 1 else (result.r1 + i, result.c1)
        if r > MAX_ROW or c > MAX_COL:
            return REF
        found = Ref(result.source, result.sheet, r, c, r, c).value(0, 0)
    else:
        line = [x for row in _grid(result).array for x in row]  # type: ignore[union-attr]
        if i >= len(line):
            return NA
        found = line[i]
    return 0.0 if found is BLANK else found
