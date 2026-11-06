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
from xlrecalc._coerce import Context, scalar, to_number
from xlrecalc._compare import compare, lookup_equal, wildcard
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._store import MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Ref

NA, VALUE, REF = ErrorValue.NA, ErrorValue.VALUE, ErrorValue.REF
L = "lookup and reference"
_WILD = re.compile(r"[*?~]")


def _exact(lookup: Any, cells: list[tuple[int, Any]]) -> int | None:
    """The first position holding the value. A blank never matches, so only cells are visited."""
    if isinstance(lookup, str) and _WILD.search(lookup):
        pattern = wildcard(lookup)
        for i, v in cells:
            if isinstance(v, str) and pattern.fullmatch(v):
                return i
        return None
    for i, v in cells:
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


def _approximate(
    lookup: Any, cells: list[tuple[int, Any]], descending: bool, ctx: Context
) -> int | None:
    """Last position whose value is <= lookup (>= when descending), by bisection.

    It bisects the cells of the lookup value's type: a blank is never one.
    """
    kind = _kind(lookup)
    candidates = [cell for cell in cells if _kind(cell[1]) == kind]
    lo, hi, found = 0, len(candidates) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        i, v = candidates[mid]
        order = compare(v, lookup, ">=" if descending else "<=", ctx.quirks)
        if order is True:
            found = i
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


def _line(ref: Ref | _Grid, horizontal: bool) -> list[tuple[int, Any]]:
    """A range's first row or column as (position, value), blanks left out (§7.7)."""
    if isinstance(ref, Ref):
        last_row, last_col = (ref.r1, ref.c2) if horizontal else (ref.r2, ref.c1)
        line = Ref(ref.source, ref.sheet, ref.r1, ref.c1, last_row, last_col)
        return [(c if horizontal else r, v) for r, c, v in line.cells()]
    if horizontal:
        return [(c, ref.value(0, c)) for c in range(ref.columns)]
    return [(r, ref.value(r, 0)) for r in range(ref.rows)]


def _walked(ref: Ref | _Grid, horizontal: bool) -> Any:
    """``_line`` read as far as a search goes.

    An exact search stops at its match and depends only on the cells before it, so a formula
    below the match is not circular (capture ``index_range``).
    """
    if not isinstance(ref, Ref):
        return _line(ref, horizontal)
    last_row, last_col = (ref.r1, ref.c2) if horizontal else (ref.r2, ref.c1)
    line = Ref(ref.source, ref.sheet, ref.r1, ref.c1, last_row, last_col)
    return ((c if horizontal else r, v) for r, c, v in line.walk())


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
        cells = list(enumerate(x for row in ref for x in row))
    elif isinstance(ref, Ref) and (ref.rows == 1 or ref.columns == 1):
        cells = _walked(ref, ref.rows == 1) if kind == 0 else _line(ref, ref.rows == 1)
    else:
        return NA
    k = 0 if kind == 0 else (1 if kind > 0 else -1)
    i = _exact(v, cells) if k == 0 else _approximate(v, cells, k < 0, ctx)
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
    i = (
        _approximate(v, _line(table, horizontal), False, ctx)
        if approx
        else _exact(v, _walked(table, horizontal))
    )
    if i is None:
        return NA
    result = table.value(n - 1, i) if horizontal else table.value(i, n - 1)
    return result  # blank when it finds a blank cell: ISBLANK sees it (capture criteria)


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


@function("INDEX", "array", "number", "number", min_args=2, reference=True, category=L)
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
    """``ROW`` of what is not a reference.

    An error stays, a value (a name for one, capture ``names``) is ``#VALUE!``, and anything else
    (a union) is ``#REF!``.
    """
    if isinstance(ref, float | str | bool):
        return VALUE
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
        return found
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
    return found


def _mode(arg: Any, default: int, allowed: tuple[int, ...], ctx: Context) -> int | ErrorValue:
    """A match or search mode: truncated, an omitted one its default."""
    if arg is MISSING:
        return default
    n = to_number(arg, ctx)
    if isinstance(n, ErrorValue):
        return n
    k = math.trunc(n)
    return k if k in allowed else VALUE


def _x_cells(array: Any, lazily: bool = False) -> Any:
    """A one-line range or array as (position, value), blanks left out; ``lazily``, as searched."""
    if isinstance(array, ErrorValue):
        return array
    if isinstance(array, tuple):
        if len(array) > 1 and len(array[0]) > 1:
            return VALUE
        return list(enumerate(x for row in array for x in row))
    if isinstance(array, Ref):
        if array.rows > 1 and array.columns > 1:
            return VALUE
        return _walked(array, array.rows == 1) if lazily else _line(array, array.rows == 1)
    return [(0, array)]


def _x_size(array: Any) -> int:
    """The number of positions in a one-line range or array."""
    if isinstance(array, Ref):
        return array.rows * array.columns
    if isinstance(array, tuple):
        return len(array) * len(array[0])
    return 1


def _x_rank(v: Any) -> int:
    """Excel's order across types: numbers, text, logicals, and a blank key after them."""
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0 if isinstance(v, float) else 3


def _x_compare(key: Any, v: Any, ctx: Context) -> int | None:
    """-1, 0 or 1 for a cell against the key; ``None`` for an error, which no search takes."""
    if isinstance(v, ErrorValue) or v is BLANK:
        return None
    if key is BLANK:
        return -1
    a, b = _x_rank(v), _x_rank(key)
    if a != b:
        return -1 if a < b else 1
    if lookup_equal(key, v):
        return 0
    return -1 if compare(v, key, "<", ctx.quirks) is True else 1


def _x_gap(cells: list[tuple[int, Any]], size: int, search: int) -> int | None:
    """The first blank position of a line of ``size``, or the last when searching back."""
    held = {i for i, _ in cells}
    if len(held) >= size:
        return None
    positions = range(size) if search > 0 else range(size - 1, -1, -1)
    return next(i for i in positions if i not in held)


def _x_search(
    key: Any, cells: list[tuple[int, Any]], size: int, match: int, search: int, ctx: Context
) -> int | None:
    """The position ``XLOOKUP`` and ``XMATCH`` find, or ``None``.

    A blank key finds a blank cell; otherwise blanks are never found, so only cells are read.
    """
    if key is BLANK:
        gap = _x_gap(cells, size, search)
        if gap is not None:
            return gap
    if search in (2, -2):
        return _x_binary(key, cells, match, search, ctx)
    order = cells if search == 1 else reversed(cells)
    if match == 2 and isinstance(key, str):
        pattern = wildcard(key)
        return next((i for i, v in order if isinstance(v, str) and pattern.fullmatch(v)), None)
    best = None
    best_value: Any = None
    for i, v in order:
        c = _x_compare(key, v, ctx)
        if c is None:
            continue
        if c == 0:
            return i
        if c == match and (best is None or _x_compare(best_value, v, ctx) == -match):
            best, best_value = i, v
    return best


def _x_binary(
    key: Any, cells: list[tuple[int, Any]], match: int, search: int, ctx: Context
) -> int | None:
    """Bisection over the cells, sorted ascending (``search`` 2) or descending (-2)."""
    lo, hi = 0, len(cells) - 1
    below = above = None
    while lo <= hi:
        mid = (lo + hi) // 2
        i, v = cells[mid]
        c = _x_compare(key, v, ctx)
        if c == 0:
            return i
        step_up = (c is None or c < 0) if search == 2 else (c is not None and c > 0)
        if c is not None:
            if c < 0:
                below = i if below is None or search == 2 else below
                below = i
            else:
                above = i
        if step_up:
            lo = mid + 1
        else:
            hi = mid - 1
    return below if match == -1 else above if match == 1 else None


@function(
    "_xlfn.XMATCH",
    "value",
    "array",
    "raw",
    "raw",
    min_args=2,
    errors=True,
    status="partial",
    note="the regular-expression match mode is not implemented",
    category=L,
)
def XMATCH(ctx: Context, key: Any, array: Any, match: Any = MISSING, search: Any = MISSING) -> Any:
    """The position of a value in a row or column, by match and search mode."""
    v = scalar(key, ctx)
    if isinstance(v, ErrorValue):
        return v
    m = _mode(match, 0, (0, -1, 1, 2, 3), ctx)
    if isinstance(m, ErrorValue):
        return m
    if m == 3:
        return ctx.unsupported("the regular-expression match mode")  # type: ignore[attr-defined]
    s = _mode(search, 1, (1, -1, 2, -2), ctx)
    if isinstance(s, ErrorValue):
        return s
    if m == 2 and s in (2, -2):
        return VALUE  # wildcards cannot be bisected (capture modern)
    cells = _x_cells(array, lazily=s == 1 and v is not BLANK and m in (0, 2))
    if isinstance(cells, ErrorValue):
        return cells
    found = _x_search(v, cells, _x_size(array), m, s, ctx)
    return NA if found is None else float(found + 1)


@function(
    "_xlfn.XLOOKUP",
    "key",
    "array",
    "array",
    "raw",
    "raw",
    "raw",
    min_args=3,
    errors=True,
    reference=True,
    status="partial",
    note="the regular-expression match mode is not implemented",
    category=L,
)
def XLOOKUP(
    ctx: Context,
    key: Any,
    array: Any,
    results: Any,
    not_found: Any = MISSING,
    match: Any = MISSING,
    search: Any = MISSING,
) -> Any:
    """The result at the position found: a cell, or the row or column of a wider range."""
    v = scalar(key, ctx)
    if isinstance(v, ErrorValue):
        return v
    m = _mode(match, 0, (0, -1, 1, 2, 3), ctx)
    if isinstance(m, ErrorValue):
        return m
    if m == 3:
        return ctx.unsupported("the regular-expression match mode")  # type: ignore[attr-defined]
    s = _mode(search, 1, (1, -1, 2, -2), ctx)
    if isinstance(s, ErrorValue):
        return s
    if m == 2 and s in (2, -2):
        return VALUE  # wildcards cannot be bisected (capture modern)
    cells = _x_cells(array, lazily=s == 1 and v is not BLANK and m in (0, 2))
    if isinstance(cells, ErrorValue):
        return cells
    if isinstance(results, ErrorValue):
        return results
    across = isinstance(array, Ref) and array.rows == 1 and array.columns > 1
    if isinstance(array, tuple):
        across = len(array) == 1 and len(array[0]) > 1
    size = _x_size(array)
    if isinstance(results, Ref):
        if (results.columns if across else results.rows) != size:
            return VALUE
    elif isinstance(results, tuple):
        if (len(results[0]) if across else len(results)) != size:
            return VALUE
    elif size != 1:
        return VALUE
    found = _x_search(v, cells, size, m, s, ctx)
    if found is None:
        return NA if not_found is MISSING else not_found
    if isinstance(results, Ref):
        if across:
            return Ref(results.source, results.sheet, results.r1, results.c1 + found,
                       results.r2, results.c1 + found)  # fmt: skip
        return Ref(results.source, results.sheet, results.r1 + found, results.c1,
                   results.r1 + found, results.c2)  # fmt: skip
    if isinstance(results, tuple):
        if across:
            return tuple((row[found],) for row in results)
        return (results[found],)
    return results
