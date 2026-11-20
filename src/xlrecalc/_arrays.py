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

__all__ = [
    "ArrayLimitError",
    "NotImplementedFormError",
    "Padded",
    "broadcast",
    "check_area",
    "element",
    "fill",
    "is_array",
    "shape",
    "to_array",
]

Array = tuple[tuple[Any, ...], ...]
NA = ErrorValue.NA


class ArrayLimitError(Exception):
    """An array a formula would make has more cells than ``ReadLimits.max_array_cells``.

    Not an Excel error value: the engine catches it per cell, keeps the cell's cached value and
    reports it.
    """

    def __init__(self, cells: int, limit: int) -> None:
        super().__init__(f"an array of {cells:,} cells, more than the limit of {limit:,}")
        self.limit = limit


class NotImplementedFormError(Exception):
    """A formula uses a form that is not implemented, which only evaluating it can show.

    Not an Excel error value: the engine catches it per cell, keeps the cell's cached value and
    reports it, so that a form with no answer here never shows a wrong one.
    """

    def __init__(self, label: str) -> None:
        super().__init__(label)
        self.label = label


def check_area(rows: int, cols: int, limit: int | None) -> None:
    """Refuse an array of ``rows`` by ``cols`` past ``limit``, before anything is allocated."""
    if limit is not None and rows * cols > limit:
        raise ArrayLimitError(rows * cols, limit)


def is_array(v: object) -> bool:
    return type(v) is tuple


def to_array(ref: Ref) -> Array:
    """A reference's values as an array (blanks stay blank), within the engine's array limit."""
    check_area(ref.rows, ref.columns, getattr(ref.source, "limit", None))
    return tuple(tuple(ref.value(r, c) for c in range(ref.columns)) for r in range(ref.rows))


class Padded:
    """An array of ``rows`` by ``cols`` stored as its top rows and one value for the rest.

    A range of whole columns made into an array is mostly blank: ``head`` holds its rows down to
    the last one holding anything, every row below holds ``fill`` in each column. It equals its
    dense expansion (:meth:`dense`) in every respect. Only the elementwise operators and the
    functions taught to read it ever receive one (ARCHITECTURE.md §7.7); the compiler hands it to
    no other code.
    """

    __slots__ = ("cols", "fill", "head", "rows")

    def __init__(self, head: Array, rows: int, cols: int, fill: Any) -> None:
        self.head, self.rows, self.cols, self.fill = head, rows, cols, fill

    @property
    def padding(self) -> int:
        """How many cells hold ``fill``: those below the head."""
        return (self.rows - len(self.head)) * self.cols

    def dense(self, limit: int | None = None) -> Array:
        """The array in full, as ``to_array`` would make it, within ``limit``."""
        check_area(self.rows, self.cols, limit)
        row = (self.fill,) * self.cols
        return self.head + (row,) * (self.rows - len(self.head))

    def map(self, fn: Callable[[Any], Any]) -> Padded:
        """``fn`` applied to every element: to each of the head's and once to the fill."""
        head = tuple(tuple(fn(v) for v in row) for row in self.head)
        return Padded(head, self.rows, self.cols, fn(self.fill))

    def __repr__(self) -> str:
        return f"Padded({len(self.head)} of {self.rows} by {self.cols} rows, fill {self.fill!r})"


def pair(fn: Callable[[Any, Any], Any], a: Any, b: Any) -> Padded | None:
    """``fn`` element by element over two operands, at least one padded, kept padded.

    ``None`` when the result cannot keep one fill: two padded arrays of different sizes, or a
    padded array with an array that is not padded (the caller then works on dense arrays).
    """
    if type(a) is Padded and type(b) is Padded:
        if (a.rows, a.cols) != (b.rows, b.cols):
            return None
        height = max(len(a.head), len(b.head))
        x, y = _rows(a, height), _rows(b, height)
        head = tuple(
            tuple(fn(u, v) for u, v in zip(r, s, strict=True)) for r, s in zip(x, y, strict=True)
        )
        return Padded(head, a.rows, a.cols, fn(a.fill, b.fill))
    if type(a) is Padded and type(b) is not tuple:
        return a.map(lambda v: fn(v, b))
    if type(b) is Padded and type(a) is not tuple:
        return b.map(lambda v: fn(a, v))
    return None


def _rows(a: Padded, height: int) -> Array:
    """``a``'s first ``height`` rows, the head followed by rows of its fill."""
    return a.head + ((a.fill,) * a.cols,) * (height - len(a.head))


def shape(a: Array) -> tuple[int, int]:
    return len(a), len(a[0]) if a else 0


def element(a: Array, rows: int, cols: int, i: int, j: int) -> Any:
    """Element (i, j) of a ``rows`` by ``cols`` array spread over a larger range.

    A dimension of 1 repeats; a position past the end is ``#N/A``.
    """
    return _at(a, rows, cols, i, j)


def _at(a: Array, rows: int, cols: int, i: int, j: int) -> Any:
    ri = 0 if rows == 1 else i
    cj = 0 if cols == 1 else j
    if ri >= rows or cj >= cols:
        return NA
    return a[ri][cj]


def broadcast(
    fn: Callable[..., Any],
    args: Sequence[Any],
    lifted: Sequence[bool],
    limit: int | None = None,
) -> Array:
    """``fn`` applied element by element over the arguments marked in ``lifted``.

    Each lifted argument is an array (or a value, taken as a 1 by 1 array); the others are
    passed to every call as they are. A row and a column make an array of every pair, so the
    result's size is checked against ``limit`` first.
    """
    arrays = [a if lifted[k] and is_array(a) else None for k, a in enumerate(args)]
    shapes = [shape(a) for a in arrays if a is not None]
    rows = max((s[0] for s in shapes), default=1)
    cols = max((s[1] for s in shapes), default=1)
    check_area(rows, cols, limit)
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
