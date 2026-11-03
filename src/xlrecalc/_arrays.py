"""Arrays: broadcasting, filling a declared range, spilling (ARCHITECTURE.md §7.7).

An array is a tuple of rows, each a tuple of values (``((1.0, 2.0), (3.0, 4.0))``), as array
constants already are. Excel's rules, from capture ``arrays``:

- an operation between two arrays has the larger of their sizes in each dimension; a
  dimension of 1 repeats, and a position past the end of the other operand is ``#N/A``
  (``{1,2,3}+{1;2}`` is 2 by 3; ``{1,2,3}+{1,2}`` is ``{2,4,#N/A}``);
- a function taking one value, given an array, is applied to each element;
- a legacy array formula fills its declared range with the same rule: an array of 1 row or 1
  column repeats, a position past the end is ``#N/A``.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import Any

from xlrecalc._values import BLANK, ErrorValue, Ref

__all__ = ["broadcast", "fill", "is_array", "shape", "to_array"]

Array = tuple[tuple[Any, ...], ...]
NA = ErrorValue.NA


def is_array(v: object) -> bool:
    return type(v) is tuple


def to_array(ref: Ref) -> Array:
    """A reference's values as an array (blanks stay blank)."""
    return tuple(tuple(ref.value(r, c) for c in range(ref.columns)) for r in range(ref.rows))


def shape(a: Array) -> tuple[int, int]:
    return len(a), len(a[0]) if a else 0


def _at(a: Array, rows: int, cols: int, i: int, j: int) -> Any:
    ri = 0 if rows == 1 else i
    cj = 0 if cols == 1 else j
    if ri >= rows or cj >= cols:
        return NA
    return a[ri][cj]


def broadcast(fn: Callable[..., Any], args: Sequence[Any], lifted: Sequence[bool]) -> Array:
    """``fn`` applied element by element over the arguments marked in ``lifted``.

    Each lifted argument is an array (or a value, taken as a 1 by 1 array); the others are
    passed to every call as they are.
    """
    arrays = [a if lifted[k] and is_array(a) else None for k, a in enumerate(args)]
    shapes = [shape(a) for a in arrays if a is not None]
    rows = max((s[0] for s in shapes), default=1)
    cols = max((s[1] for s in shapes), default=1)
    out = []
    for i in range(rows):
        row = []
        for j in range(cols):
            call = [
                _at(a, *shape(a), i, j) if a is not None else args[k] for k, a in enumerate(arrays)
            ]
            value = fn(*call)
            row.append(value[0][0] if is_array(value) else value)
        out.append(tuple(row))
    return tuple(out)


def fill(value: Any, rows: int, cols: int) -> Array:
    """A result spread over a declared range of ``rows`` by ``cols``, as a legacy array fills."""
    a: Array = value if is_array(value) else ((value,),)
    r, c = shape(a)
    return tuple(tuple(_at(a, r, c, i, j) for j in range(cols)) for i in range(rows))


def blank_to_zero(a: Array) -> Array:
    """An array as cells hold it: a blank element shows 0."""
    return tuple(tuple(0.0 if v is BLANK else v for v in row) for row in a)
