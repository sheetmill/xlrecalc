"""The functions that call a ``LAMBDA`` (ARCHITECTURE.md §5.4): ``MAP`` and its siblings.

Each calls its function value through the engine's call stack (``Ops.call``) and makes an
array that a dynamic-array formula spills. The rules come from capture ``lambda_helpers``: a
range's elements are passed as references to their cells, an array's as values; ``REDUCE`` and
``SCAN`` go row by row; a call whose result is not one value makes the whole result ``#CALC!``
(``REDUCE``'s last excepted: its result may be a range); anything but a function value where one
belongs is ``#VALUE!``.
"""

from __future__ import annotations

import math
from collections.abc import Iterator
from typing import Any

from xlrecalc._arrays import check_area, element
from xlrecalc._coerce import scalar, to_number
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Lambda, Ref

L = "logical"
VALUE, CALC = ErrorValue.VALUE, ErrorValue("#CALC!")
_NOT_ONE = object()  # a call's result that is not one value


def _limit(ctx: Any, rows: int, cols: int) -> None:
    check_area(rows, cols, getattr(ctx, "max_array_cells", None))


def _size(a: Any) -> tuple[int, int]:
    if isinstance(a, Ref):
        return a.rows, a.columns
    if type(a) is tuple:
        return len(a), len(a[0])
    return 1, 1


def _at(a: Any, rows: int, cols: int, i: int, j: int) -> Any:
    """Element (i, j): a range's as a reference to its cell, an array's as its value."""
    if isinstance(a, Ref):
        if (rows > 1 and i >= rows) or (cols > 1 and j >= cols):
            return ErrorValue.NA
        r, c = a.r1 + (0 if rows == 1 else i), a.c1 + (0 if cols == 1 else j)
        return Ref(a.source, a.sheet, r, c, r, c)
    if type(a) is tuple:
        return element(a, rows, cols, i, j)
    return BLANK if a is MISSING else a


def _one(v: Any) -> Any:
    """A call's result as one value, or ``_NOT_ONE`` for a range or an array of several."""
    if isinstance(v, Ref):
        return v.value(0, 0) if v.is_cell() else _NOT_ONE
    if type(v) is tuple:
        return v[0][0] if len(v) == 1 and len(v[0]) == 1 else _NOT_ONE
    return BLANK if v is MISSING else v


def _shown(v: Any) -> Any:
    """A value as a cell of a result holds it: a blank is 0."""
    return 0.0 if v is BLANK else v


def _elements(a: Any) -> Iterator[Any]:
    """Every element of a range, an array or one value, row by row."""
    rows, cols = _size(a)
    for i in range(rows):
        for j in range(cols):
            yield _at(a, rows, cols, i, j)


@function("_XLFN.MAP", "raw", "function", min_args=2, variadic=True, category=L, makes_array=True)
def MAP(ctx: Any, *args: Any) -> Any:
    """``fn`` of the elements at each position of the arrays, a shorter one padded with ``#N/A``."""
    *arrays, fn = args
    if type(fn) is not Lambda:
        return VALUE
    sizes = [_size(a) for a in arrays]
    rows, cols = max(s[0] for s in sizes), max(s[1] for s in sizes)
    _limit(ctx, rows, cols)
    call = ctx.ops.call
    out = []
    for i in range(rows):
        row = []
        for j in range(cols):
            given = tuple(_at(a, r, c, i, j) for a, (r, c) in zip(arrays, sizes, strict=True))
            v = _one(call(fn, given))
            if v is _NOT_ONE:
                return CALC
            row.append(_shown(v))
        out.append(tuple(row))
    return tuple(out)


@function("_XLFN.REDUCE", "raw", "raw", "function", category=L, makes_array=True)
def REDUCE(ctx: Any, initial: Any, array: Any, fn: Any) -> Any:
    """``fn(total, element)`` over the elements in turn from ``initial``; its last result."""
    if type(fn) is not Lambda:
        return VALUE
    total = BLANK if initial is MISSING else initial
    call = ctx.ops.call
    for v in _elements(array):
        total = call(fn, (total, v))
    return total


@function("_XLFN.SCAN", "raw", "raw", "function", category=L, makes_array=True)
def SCAN(ctx: Any, initial: Any, array: Any, fn: Any) -> Any:
    """``REDUCE``'s running results, in the shape of the array."""
    if type(fn) is not Lambda:
        return VALUE
    rows, cols = _size(array)
    _limit(ctx, rows, cols)
    total = BLANK if initial is MISSING else initial
    call = ctx.ops.call
    values = []
    for v in _elements(array):
        total = _one(call(fn, (total, v)))
        if total is _NOT_ONE:
            return CALC
        values.append(_shown(total))
    return tuple(tuple(values[i * cols : (i + 1) * cols]) for i in range(rows))


def _lines(ctx: Any, array: Any, fn: Any, by_row: bool) -> Any:
    if type(fn) is not Lambda:
        return VALUE
    rows, cols = _size(array)
    n = rows if by_row else cols
    _limit(ctx, n, 1)
    call = ctx.ops.call
    out = []
    for k in range(n):
        if isinstance(array, Ref):
            r1, c1 = (array.r1 + k, array.c1) if by_row else (array.r1, array.c1 + k)
            r2, c2 = (r1, array.c2) if by_row else (array.r2, c1)
            line: Any = Ref(array.source, array.sheet, r1, c1, r2, c2)
        elif type(array) is tuple:
            line = (array[k],) if by_row else tuple((row[k],) for row in array)
        else:
            line = array
        v = _one(call(fn, (line,)))
        if v is _NOT_ONE:
            return CALC
        out.append(_shown(v))
    return tuple((v,) for v in out) if by_row else (tuple(out),)


@function("_XLFN.BYROW", "raw", "function", category=L, makes_array=True)
def BYROW(ctx: Any, array: Any, fn: Any) -> Any:
    """``fn`` of each row (a range's as a reference), one value a row."""
    return _lines(ctx, array, fn, True)


@function("_XLFN.BYCOL", "raw", "function", category=L, makes_array=True)
def BYCOL(ctx: Any, array: Any, fn: Any) -> Any:
    """``fn`` of each column (a range's as a reference), one value a column."""
    return _lines(ctx, array, fn, False)


def _dimension(v: Any, ctx: Any) -> int | ErrorValue:
    """A size: a positive whole number, else ``#VALUE!`` (2.7 included: capture lambda_helpers)."""
    n = to_number(scalar(v, ctx), ctx)
    if isinstance(n, ErrorValue):
        return n
    if n < 1 or n != math.floor(n):
        return VALUE
    return int(n)


@function("_XLFN.MAKEARRAY", "raw", "raw", "function", category=L, makes_array=True)
def MAKEARRAY(ctx: Any, rows: Any, cols: Any, fn: Any) -> Any:
    """A ``rows`` by ``cols`` array of ``fn(row, column)``, counted from 1."""
    r, c = _dimension(rows, ctx), _dimension(cols, ctx)
    if isinstance(r, ErrorValue):
        return r
    if isinstance(c, ErrorValue):
        return c
    if type(fn) is not Lambda:
        return VALUE
    _limit(ctx, r, c)
    call = ctx.ops.call
    out = []
    for i in range(r):
        row = []
        for j in range(c):
            v = _one(call(fn, (float(i + 1), float(j + 1))))
            if v is _NOT_ONE:
                return CALC
            row.append(_shown(v))
        out.append(tuple(row))
    return tuple(out)
