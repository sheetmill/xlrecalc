"""Engineering functions: number bases, bits, DELTA, GESTEP (capture ``engineering_family``).

A number in another base has at most 10 digits, and with all 10 the top bit makes it negative in
two's complement. A negative result always has all 10 digits, whatever ``places`` says.
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._coerce import Context, to_number, to_text
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue

NUM, VALUE = ErrorValue.NUM, ErrorValue.VALUE
E = "engineering"
_BITS = {"BIN": (2, 1), "OCT": (8, 3), "HEX": (16, 4)}
_DIGITS = "0123456789ABCDEF"
_MAX_DIGITS = 10
_MAX_BIT_NUMBER = 2**48
_MAX_SHIFT = 53


def _strict(v: Any, ctx: Context) -> float | ErrorValue:
    """A number as these functions take one: a logical is ``#VALUE!``."""
    return VALUE if isinstance(v, bool) else to_number(v, ctx)


def _decode(v: Any, source: str, ctx: Context) -> int | ErrorValue:
    """A number written in a base, as an integer."""
    if isinstance(v, bool):
        return VALUE
    text = "" if v is BLANK else to_text(v, ctx)
    if isinstance(text, ErrorValue):
        return text
    base, bits = _BITS[source]
    digits = text.upper()
    if len(digits) > _MAX_DIGITS or any(d not in _DIGITS[:base] for d in digits):
        return NUM
    value = int(digits, base) if digits else 0
    width = _MAX_DIGITS * bits
    return value - 2**width if len(digits) == _MAX_DIGITS and value >= 2 ** (width - 1) else value


def _encode(n: int, target: str, places: Any, ctx: Context) -> str | ErrorValue:
    """An integer written in a base, in ``places`` digits if given."""
    base, bits = _BITS[target]
    width = _MAX_DIGITS * bits
    # BIN holds 10 bits; OCT and HEX keep to their own widths (30 and 40).
    if not -(2 ** (width - 1)) <= n < 2 ** (width - 1):
        return NUM
    count = None
    if places is not MISSING:
        p = _strict(places, ctx)
        if isinstance(p, ErrorValue):
            return p
        count = math.trunc(p)
        if not 1 <= count <= _MAX_DIGITS:
            return NUM
    value = n + 2**width if n < 0 else n
    digits = ""
    while value:
        value, d = divmod(value, base)
        digits = _DIGITS[d] + digits
    digits = digits or "0"
    if n < 0 or count is None:
        return digits
    return NUM if len(digits) > count else digits.rjust(count, "0")


def _from_decimal(v: Any, ctx: Context) -> int | ErrorValue:
    x = _strict(v, ctx)
    if isinstance(x, ErrorValue):
        return x
    return math.trunc(x) if abs(x) < 2.0**63 else NUM


def _converter(source: str, target: str) -> None:
    name = f"{source}2{target}"
    if target == "DEC":

        @function(name, "value", category=E)
        def to_decimal(ctx: Context, v: Any) -> float | ErrorValue:
            n = _decode(v, source, ctx)
            return n if isinstance(n, ErrorValue) else float(n)

        to_decimal.__doc__ = f"A number in {source} as a decimal number."
        return

    @function(name, "value", "value", min_args=1, errors=True, category=E)
    def convert(ctx: Context, v: Any, places: Any = MISSING) -> str | ErrorValue:
        if isinstance(places, bool) or isinstance(v, bool):  # first (capture engineering_family)
            return VALUE
        if isinstance(v, ErrorValue):
            return v
        if isinstance(places, ErrorValue):
            return places
        n = _from_decimal(v, ctx) if source == "DEC" else _decode(v, source, ctx)
        return n if isinstance(n, ErrorValue) else _encode(n, target, places, ctx)

    convert.__doc__ = f"A number in {source} as a number in {target}."


for _source in ("BIN", "OCT", "DEC", "HEX"):
    for _target in ("BIN", "OCT", "DEC", "HEX"):
        if _source != _target:
            _converter(_source, _target)


def _bit_number(x: float) -> int | ErrorValue:
    """A number the bit functions take: whole, not negative, below 2^48."""
    return int(x) if 0 <= x < _MAX_BIT_NUMBER and x == math.floor(x) else NUM


def _bitwise(name: str, combine: Any, doc: str) -> None:
    @function(f"_xlfn.{name}", "number", "number", category=E)
    def bits(ctx: Context, a: float, b: float) -> float | ErrorValue:
        x, y = _bit_number(a), _bit_number(b)
        if isinstance(x, ErrorValue):
            return x
        return y if isinstance(y, ErrorValue) else float(combine(x, y))

    bits.__doc__ = doc


_bitwise("BITAND", int.__and__, "Bitwise AND of two numbers.")
_bitwise("BITOR", int.__or__, "Bitwise OR of two numbers.")
_bitwise("BITXOR", int.__xor__, "Bitwise exclusive OR of two numbers.")


def _shift(a: float, count: float) -> float | ErrorValue:
    """``a`` shifted left by ``count``.

    0 at a count of 53 and ``#NUM!`` past it, and ``#NUM!`` for a result of 2^48 or more
    (``BITLSHIFT(5,47)``), but 0 stays 0 at any count.
    """
    x = _bit_number(a)
    if isinstance(x, ErrorValue):
        return x
    k = math.trunc(count)
    if x == 0:
        return 0.0
    if abs(k) > _MAX_SHIFT:
        return NUM
    if abs(k) == _MAX_SHIFT:
        return 0.0
    result = x << k if k >= 0 else x >> -k
    return NUM if result >= _MAX_BIT_NUMBER else float(result)


@function("_xlfn.BITLSHIFT", "number", "number", category=E)
def BITLSHIFT(ctx: Context, a: float, count: float) -> float | ErrorValue:
    """A number's bits moved left, or right for a negative count."""
    return _shift(a, count)


@function("_xlfn.BITRSHIFT", "number", "number", category=E)
def BITRSHIFT(ctx: Context, a: float, count: float) -> float | ErrorValue:
    """A number's bits moved right, or left for a negative count."""
    return _shift(a, -count)


def _pair(a: Any, b: Any, ctx: Context) -> tuple[Any, Any]:
    """Two numbers, or the problem as the first: the second argument's comes first."""
    y = _strict(b, ctx)
    if isinstance(y, ErrorValue):
        return y, None
    x = _strict(a, ctx)
    return (x, None) if isinstance(x, ErrorValue) else (x, y)


@function("DELTA", "value", "value", min_args=1, errors=True, category=E)
def DELTA(ctx: Context, a: Any, b: Any = 0.0) -> float | ErrorValue:
    """1 when two numbers are equal, else 0."""
    x, y = _pair(a, b, ctx)
    return x if isinstance(x, ErrorValue) else 1.0 if x == y else 0.0


@function("GESTEP", "value", "value", min_args=1, errors=True, category=E)
def GESTEP(ctx: Context, a: Any, step: Any = 0.0) -> float | ErrorValue:
    """1 when a number is at least the step, else 0."""
    x, y = _pair(a, step, ctx)
    return x if isinstance(x, ErrorValue) else 1.0 if x >= y else 0.0
