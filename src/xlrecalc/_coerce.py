"""Coercions shared by the operators and the functions (ARCHITECTURE.md §4.3).

``Context`` is what evaluation provides: the quirks, the date system, the date for texts
without a year, and the cell being evaluated (for implicit intersection).
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Iterator
from typing import Protocol

from xlrecalc._formula import MISSING
from xlrecalc._numtext import number_to_text, text_to_number
from xlrecalc._quirks import Quirks
from xlrecalc._values import BLANK, ErrorValue, Ref, Value

__all__ = [
    "Context",
    "intersect",
    "numbers",
    "scalar",
    "to_bool",
    "to_number",
    "to_text",
]

VALUE = ErrorValue.VALUE
Raw = Value | Ref | object  # object: MISSING


class Context(Protocol):
    """What the operators and functions see of the evaluation."""

    quirks: Quirks
    date1904: bool
    today: dt.date
    sheet: int
    row: int
    col: int


def intersect(ref: Ref, ctx: Context) -> Value:
    """Implicit intersection: a reference used where one value is wanted."""
    if ref.is_cell():
        return ref.value(0, 0)
    if ref.c1 == ref.c2 and ref.r1 <= ctx.row <= ref.r2:
        return ref.value(ctx.row - ref.r1, 0)
    if ref.r1 == ref.r2 and ref.c1 <= ctx.col <= ref.c2:
        return ref.value(0, ctx.col - ref.c1)
    return VALUE


def scalar(v: Raw, ctx: Context) -> Value:
    """A raw argument as one value: references intersected, an omitted argument blank."""
    if isinstance(v, Ref):
        return intersect(v, ctx)
    if v is MISSING:
        return BLANK
    return v  # type: ignore[return-value]


def to_number(v: Raw, ctx: Context) -> float | ErrorValue:
    """A value as a number, as arithmetic and numeric arguments convert it."""
    v = scalar(v, ctx)
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, float):
        return v
    if isinstance(v, ErrorValue):
        return v
    if v is BLANK:
        return 0.0
    assert isinstance(v, str)
    n = text_to_number(v, ctx.quirks, date1904=ctx.date1904, today=ctx.today)
    return VALUE if n is None else n


def to_text(v: Raw, ctx: Context) -> str | ErrorValue:
    """A value as text, as ``&`` and text arguments convert it."""
    v = scalar(v, ctx)
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return number_to_text(v, ctx.quirks)
    if isinstance(v, ErrorValue):
        return v
    return ""


def to_bool(v: Raw, ctx: Context) -> bool | ErrorValue:
    """A value as a logical, as ``IF``'s condition and logical arguments convert it."""
    v = scalar(v, ctx)
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        return v != 0
    if isinstance(v, ErrorValue):
        return v
    if v is BLANK:
        return False
    upper = str(v).upper()
    if upper in ("TRUE", "FALSE"):
        return upper == "TRUE"
    return VALUE


def numbers(
    args: tuple[Raw, ...], ctx: Context, *, direct_text: str = "convert"
) -> Iterator[float | ErrorValue]:
    """The numbers an aggregate (``SUM``, ``AVERAGE``, ``MIN``...) takes from its arguments.

    Inside references only numbers count; text, logicals and blanks are skipped, errors are
    yielded. A direct argument is converted: a logical is 1 or 0, an omitted argument is
    skipped, and text is converted (``direct_text="convert"``, an error when it is not a
    number) or skipped when it is not a number (``"skip"``, as ``COUNT``).
    """
    for arg in args:
        if isinstance(arg, Ref):
            for v in arg.values():
                if isinstance(v, float | ErrorValue):
                    yield v
        elif arg is MISSING:
            continue
        elif isinstance(arg, bool):
            yield float(arg)
        elif isinstance(arg, float | ErrorValue):
            yield arg
        elif arg is BLANK:
            yield 0.0
        elif isinstance(arg, str):
            n = text_to_number(arg, ctx.quirks, date1904=ctx.date1904, today=ctx.today)
            if n is not None:
                yield n
            elif direct_text == "convert":
                yield VALUE
        elif isinstance(arg, tuple):  # an array constant
            for row in arg:
                for v in row:
                    if isinstance(v, float | ErrorValue) and not isinstance(v, bool):
                        yield v
