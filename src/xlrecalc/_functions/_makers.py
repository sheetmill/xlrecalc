"""Functions that make arrays from data (ARCHITECTURE.md §7.7).

Each function's rules come from the capture its docstring names, tried on two sets of values.
"""

from __future__ import annotations

import bisect
import math
import random
import re
from typing import Any

from xlrecalc._arrays import to_array
from xlrecalc._coerce import scalar, to_text
from xlrecalc._formula import MISSING
from xlrecalc._functions._dynamic import CALC, MAX_COLS, MAX_ROWS, VALUE, _count, _flag, _limit
from xlrecalc._functions._registry import function
from xlrecalc._functions._reshape import _register, _spread
from xlrecalc._values import ErrorValue, Ref

S = "statistical"
T = "text"
M = "math"


def _elements(a: Any) -> list[Any]:
    """The values of a range, an array or one value."""
    if isinstance(a, Ref):
        return [v for row in to_array(a) for v in row]
    if isinstance(a, tuple):
        return [v for row in a for v in row]
    return [a]


def _numbers(a: Any) -> list[float] | ErrorValue:
    """The numbers among the values, in order; text, logicals and blanks are left out."""
    out = []
    for v in _elements(a):
        if isinstance(v, ErrorValue):
            return v
        if type(v) is float:
            out.append(v)
    return out


# FREQUENCY -------------------------------------------------------------------------------------


@function("FREQUENCY", "array", "array", category=S, makes_array=True)
def FREQUENCY(ctx: Any, data: Any, bins: Any) -> Any:
    """How many numbers fall in each bin, in the bins' order, and above the largest in a last row.

    A number counts in the smallest bin at or above it; equal bins take their count in the
    first; without a number among the bins there is one bin, 0 (capture ``dyn_frequency``).
    """
    if data is MISSING or bins is MISSING:  # an empty argument (capture empty_args)
        return VALUE
    values = _numbers(data)
    if isinstance(values, ErrorValue):
        return values
    edges = _numbers(bins)
    if isinstance(edges, ErrorValue):
        return edges
    if not edges:
        edges = [0.0]
    order = sorted(range(len(edges)), key=lambda i: edges[i])
    ordered = [edges[i] for i in order]
    counts = [0.0] * (len(edges) + 1)
    for v in values:
        k = bisect.bisect_left(ordered, v)
        counts[order[k] if k < len(ordered) else len(edges)] += 1.0
    return tuple((c,) for c in counts)


# MODE.MULT -------------------------------------------------------------------------------------


@function("_XLFN.MODE.MULT", "array", min_args=1, variadic=True, category=S, makes_array=True)
def MODE_MULT(ctx: Any, *args: Any) -> Any:
    """Every number that occurs most often, at least twice, in the order each first appears.

    Logicals and text are left out, given directly too (capture ``dyn_frequency``); none
    repeating is ``#N/A``.
    """
    counts: dict[float, int] = {}
    for a in args:
        if a is MISSING:
            return VALUE
        values = _numbers(a)
        if isinstance(values, ErrorValue):
            return values
        for v in values:
            counts[v] = counts.get(v, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        return ErrorValue.NA
    return tuple((v,) for v, n in counts.items() if n == best)


# TEXTSPLIT -------------------------------------------------------------------------------------


def _delimiters(a: Any, ctx: Any) -> list[str] | ErrorValue | None:
    """The delimiters given, ``None`` for none; an empty one is ``#VALUE!``."""
    if a is MISSING:
        return None
    out = []
    for v in _elements(a):
        if isinstance(v, ErrorValue):
            return v
        text = to_text(v, ctx)
        if isinstance(text, ErrorValue):
            return text
        if text == "":
            return VALUE
        out.append(text)
    return out


def _pieces(text: str, delimiters: list[str] | None, fold: bool, ignore: bool) -> list[str]:
    if delimiters is None:
        return [text]
    pattern = "|".join(re.escape(d) for d in sorted(delimiters, key=len, reverse=True))
    parts = re.split(pattern, text, flags=re.IGNORECASE if fold else 0)
    return [p for p in parts if p] if ignore else parts


def _textsplit(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """A text split into columns at ``col_delimiter`` and rows at ``row_delimiter``.

    Rows shorter than the longest are padded with ``pad_with`` (``#N/A`` by default); empty
    pieces are dropped with ``ignore_empty``; ``match_mode`` 1 matches without regard to case;
    empty text is ``#VALUE!``. A range or array of texts is taken element by element (capture
    ``dyn_textsplit``).
    """
    lifted = _spread(ctx, _textsplit, given, (0,), ranges)
    if lifted is not None:
        return lifted
    padded: list[Any] = [*given, *[MISSING] * (6 - len(given))]
    text, cols, rows, ignore, mode, pad = padded
    if isinstance(text, Ref):
        text = scalar(text, ctx)
    if isinstance(text, ErrorValue):
        return text
    whole = to_text(text, ctx)
    if isinstance(whole, ErrorValue):
        return whole
    if whole == "":
        return VALUE
    col_delimiters, row_delimiters = _delimiters(cols, ctx), _delimiters(rows, ctx)
    for d in (col_delimiters, row_delimiters):
        if isinstance(d, ErrorValue):
            return d
    skip = _flag(ignore, ctx)
    if isinstance(skip, ErrorValue):
        return skip
    fold = _count(mode, ctx, 0.0)
    if isinstance(fold, ErrorValue):
        return fold
    if fold not in (0.0, 1.0):
        return VALUE
    filler = ErrorValue.NA if pad is MISSING else scalar(pad, ctx) if isinstance(pad, Ref) else pad
    lines = [
        _pieces(line, col_delimiters, fold == 1.0, skip)  # type: ignore[arg-type]
        for line in _pieces(whole, row_delimiters, fold == 1.0, skip)  # type: ignore[arg-type]
    ]
    lines = [line for line in lines if line] or [[""]]
    width = max(len(line) for line in lines)
    _limit(ctx, len(lines), width)
    return tuple(tuple(line) + (filler,) * (width - len(line)) for line in lines)


_register(
    "_XLFN.TEXTSPLIT", _textsplit, "raw", "raw", "raw", "raw", "raw", "raw", min_args=2, category=T
)


# RANDARRAY -------------------------------------------------------------------------------------


def _randarray(ctx: Any, given: tuple[Any, ...], ranges: bool) -> Any:
    """Random numbers, ``rows`` by ``columns``, from ``min`` up to ``max`` or whole between them.

    Counts are truncated, 0 is ``#CALC!`` and a negative one ``#VALUE!``; ``min`` above ``max``,
    or bounds that are not whole with ``whole_number``, are ``#VALUE!`` (capture
    ``dyn_randarray``). New at every calculation.
    """
    padded: list[Any] = [*given, *[MISSING] * (5 - len(given))]
    rows, cols, low, high, whole = padded
    numbers = [
        _count(rows, ctx, 1.0), _count(cols, ctx, 1.0), _count(low, ctx, 0.0),
        _count(high, ctx, 1.0),
    ]  # fmt: skip
    for n in numbers:
        if isinstance(n, ErrorValue):
            return n
    integer = _flag(whole, ctx)
    if isinstance(integer, ErrorValue):
        return integer
    r, c = math.trunc(numbers[0]), math.trunc(numbers[1])  # type: ignore[arg-type]
    lo, hi = float(numbers[2]), float(numbers[3])  # type: ignore[arg-type]
    if r < 0 or c < 0 or r > MAX_ROWS or c > MAX_COLS or lo > hi:
        return VALUE
    if r == 0 or c == 0:
        return CALC
    if integer and (lo != math.floor(lo) or hi != math.floor(hi)):
        return VALUE
    _limit(ctx, r, c)
    if integer:
        a, b = int(lo), int(hi)
        return tuple(tuple(float(random.randint(a, b)) for _ in range(c)) for _ in range(r))
    return tuple(tuple(lo + random.random() * (hi - lo) for _ in range(c)) for _ in range(r))


_register(
    "_XLFN.RANDARRAY",
    _randarray,
    "raw",
    "raw",
    "raw",
    "raw",
    "raw",
    min_args=0,
    volatile=True,
    category=M,
)
