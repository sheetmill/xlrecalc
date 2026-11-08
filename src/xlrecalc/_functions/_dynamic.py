"""Functions that make arrays (ARCHITECTURE.md §7.7).

Each returns rows of values that a dynamic-array formula spills. The rules come from the
``dyn_*`` captures, each tried on two sets of values.
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._arrays import ArrayLimitError, broadcast, to_array
from xlrecalc._coerce import scalar, to_number
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
