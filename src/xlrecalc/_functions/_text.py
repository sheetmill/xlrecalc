"""Text functions: CONCATENATE, LEFT, RIGHT, MID, LEN, FIND, TRIM, UPPER, LOWER, VALUE, TEXT."""

from __future__ import annotations

import math
import re
from decimal import ROUND_HALF_UP, Decimal, localcontext
from typing import Any

import xlformat

from xlrecalc._coerce import Context, scalar, to_number, to_text
from xlrecalc._formula import MISSING
from xlrecalc._functions._math import _decimal
from xlrecalc._functions._registry import function
from xlrecalc._numtext import text_to_number
from xlrecalc._values import BLANK, ErrorValue

VALUE = ErrorValue.VALUE
T = "text"
_SPACES = re.compile(" +")


@function("CONCATENATE", "text", min_args=1, variadic=True, category=T)
def CONCATENATE(ctx: Context, *parts: str) -> str | ErrorValue:
    """The texts joined."""
    text = "".join(parts)
    return text if len(text) <= 32767 else VALUE


def _count(n: float, default: int) -> int | ErrorValue:
    k = math.trunc(n)
    return VALUE if k < 0 else k


@function("LEFT", "text", "number", min_args=1, category=T)
def LEFT(ctx: Context, text: str, n: float = 1.0) -> str | ErrorValue:
    """The first ``n`` characters."""
    k = _count(n, 1)
    return k if isinstance(k, ErrorValue) else text[:k]


@function("RIGHT", "text", "number", min_args=1, category=T)
def RIGHT(ctx: Context, text: str, n: float = 1.0) -> str | ErrorValue:
    """The last ``n`` characters."""
    k = _count(n, 1)
    if isinstance(k, ErrorValue):
        return k
    return text[len(text) - k :] if k else ""


@function("MID", "text", "number", "number", category=T)
def MID(ctx: Context, text: str, start: float, n: float) -> str | ErrorValue:
    """``n`` characters from position ``start`` (1-based)."""
    s, k = math.trunc(start), math.trunc(n)
    if s < 1 or k < 0:
        return VALUE
    return text[s - 1 : s - 1 + k]


@function("LEN", "text", category=T)
def LEN(ctx: Context, text: str) -> float:
    """Number of characters."""
    return float(len(text))


@function("FIND", "text", "text", "number", min_args=2, category=T)
def FIND(ctx: Context, needle: str, haystack: str, start: float = 1.0) -> float | ErrorValue:
    """Position of ``needle`` in ``haystack`` from ``start``, with case; ``#VALUE!`` if absent."""
    s = math.trunc(start)
    if s < 1 or s > len(haystack) + 1:
        return VALUE
    i = haystack.find(needle, s - 1)
    return VALUE if i < 0 else float(i + 1)


@function("TRIM", "text", category=T)
def TRIM(ctx: Context, text: str) -> str:
    """Spaces removed at both ends and runs of spaces inside made one."""
    return _SPACES.sub(" ", text.strip(" "))


@function("UPPER", "text", category=T)
def UPPER(ctx: Context, text: str) -> str:
    """Upper case."""
    return text.upper()


@function("LOWER", "text", category=T)
def LOWER(ctx: Context, text: str) -> str:
    """Lower case."""
    return text.lower()


@function("VALUE", "value", category=T)
def VALUE_(ctx: Context, v: Any) -> float | ErrorValue:
    """Text as a number; a logical is ``#VALUE!``, unlike in arithmetic (§4.3)."""
    if isinstance(v, bool):
        return VALUE
    return to_number(v, ctx)


@function("TEXT", "value", "value", category=T)
def TEXT(ctx: Context, v: Any, fmt: Any) -> str | ErrorValue:
    """A value formatted with a number format, by ``xlformat.text_function`` (§8.3).

    Text that reads as a number or a date is formatted as that number, and a logical format is
    ``#VALUE!`` (capture ``textfmt``); xlformat does the rest as ``TEXT`` does it: the empty
    format, a format Excel refuses, a value the format cannot show.
    """
    if v is BLANK or v is MISSING:
        v = 0.0
    elif isinstance(v, str):
        n = text_to_number(v, ctx.quirks, date1904=ctx.date1904, today=ctx.today)
        if n is not None:
            v = n
    if isinstance(fmt, bool):
        return VALUE
    if fmt is BLANK or fmt is MISSING:
        fmt = ""
    fmt = to_text(fmt, ctx)
    if isinstance(fmt, ErrorValue):
        return fmt
    text = xlformat.text_function(v, fmt, date1904=ctx.date1904, quirks=ctx.quirks.to_xlformat())
    if isinstance(text, xlformat.ExcelError):
        return ErrorValue(str(text))
    return text


_MAX_TEXT = 32767


@function("REPT", "text", "number", category=T)
def REPT(ctx: Context, text: str, times: float) -> str | ErrorValue:
    """A text repeated; ``#VALUE!`` past 32,767 characters."""
    n = math.trunc(times)
    if n < 0 or len(text) * n > _MAX_TEXT:
        return VALUE
    return text * n


@function("REPLACE", "text", "number", "number", "text", category=T)
def REPLACE(ctx: Context, old: str, start: float, count: float, new: str) -> str | ErrorValue:
    """A text with ``count`` characters from ``start`` replaced by ``new``."""
    s, n = math.trunc(start), math.trunc(count)
    if s < 1 or n < 0 or s > _MAX_TEXT:
        return VALUE
    result = old[: s - 1] + new + old[s - 1 + n :]
    return VALUE if len(result) > _MAX_TEXT else result


# Excel for Mac numbers characters in the Mac OS Roman set (capture misc); Excel for Windows
# uses its ANSI code page instead.
@function("CHAR", "number", category=T)
def CHAR(ctx: Context, n: float) -> str | ErrorValue:
    """The character with a code from 1 to 255 (Mac OS Roman on the Mac)."""
    k = math.trunc(n)
    if not 1 <= k <= 255:
        return VALUE
    return bytes([k]).decode("mac_roman")


@function("CODE", "text", category=T)
def CODE(ctx: Context, text: str) -> float | ErrorValue:
    """The code of a text's first character (Mac OS Roman; 95 outside it)."""
    if not text:
        return VALUE
    try:
        return float(text[0].encode("mac_roman")[0])
    except UnicodeEncodeError:
        return 95.0


def _fixed_text(x: float, decimals: float, grouping: bool) -> str | ErrorValue:
    """``x`` rounded as ROUND rounds it, written with ``decimals`` places (none below 0).

    The rounding works on the 15-significant-digit decimal a number shows, so digits past the
    fifteenth are zeros (``FIXED(1.5E20)`` is ``150,000,000,000,000,000,000.00``).
    """
    d = math.trunc(decimals)
    if d > 127:
        return VALUE
    with localcontext() as context:
        context.prec = 500
        try:
            rounded = _decimal(x).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
        except ArithmeticError:
            return VALUE
        if rounded == 0:
            rounded = abs(rounded)
        return format(rounded, f"{',' if grouping else ''}.{max(d, 0)}f")


@function("DOLLAR", "number", "number", min_args=1, category=T)
def DOLLAR(ctx: Context, x: float, decimals: float = 2.0) -> str | ErrorValue:
    """A number as currency text: ``$1,234.57``, negatives in parentheses.

    With decimals of 0 or more a negative that rounds to zero keeps its parentheses
    (``DOLLAR(-0.004)`` is ``($0.00)``); with fewer it is rounded first (captured).
    """
    text = _fixed_text(x, decimals, True)
    if isinstance(text, ErrorValue):
        return text
    if text.startswith("-"):
        text = text[1:]
    elif not (x < 0 and math.trunc(decimals) >= 0):
        return f"${text}"
    return f"(${text})"


@function("FIXED", "number", "raw", "logical", min_args=1, category=T)
def FIXED(ctx: Context, x: float, decimals: Any = 2.0, no_commas: bool = False) -> Any:
    """A number as text with a fixed number of decimals, grouped unless ``no_commas``.

    An empty decimals argument means 2 here (``FIXED(x,,TRUE)``), where DOLLAR takes 0.
    """
    if decimals is MISSING:
        decimals = 2.0
    d = to_number(scalar(decimals, ctx), ctx)
    if isinstance(d, ErrorValue):
        return d
    return _fixed_text(x, d, not no_commas)
