"""Logical and information functions: IF, AND, OR, NOT, IFERROR, IFNA, the IS functions, NA."""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._coerce import Context, scalar, to_bool, to_number
from xlrecalc._compare import switch_equal
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Ref

NA, VALUE = ErrorValue.NA, ErrorValue.VALUE
L, INFO = "logical", "information"


@function("IF", "logical", "branch", "branch", min_args=1, category=L)
def IF(ctx: Context, test: bool, yes: Any = False, no: Any = False) -> Any:
    """``yes`` when the test is true, else ``no``; an omitted branch is 0, an absent one FALSE."""
    result = yes if test else no
    return 0.0 if result is MISSING else result


def _logicals(args: tuple[Any, ...], ctx: Context) -> list[bool] | ErrorValue:
    """The logicals AND and OR combine: in references logicals and numbers, text skipped."""
    out: list[bool] = []
    for arg in args:
        if arg is MISSING:  # an empty argument is FALSE (capture statistics_family)
            out.append(False)
        elif isinstance(arg, Ref):
            for _, _, v in arg.cells():
                if isinstance(v, ErrorValue):
                    return v
                if isinstance(v, bool | float):
                    out.append(bool(v))
            continue
        v = scalar(arg, ctx)
        if isinstance(v, ErrorValue):
            return v
        if isinstance(v, bool | float):
            out.append(bool(v))
        elif v is BLANK:
            continue
        elif isinstance(v, str):
            upper = v.upper()
            if upper not in ("TRUE", "FALSE"):
                return VALUE
            out.append(upper == "TRUE")
        elif isinstance(v, tuple):
            out.extend(bool(x) for row in v for x in row if isinstance(x, bool | float))
    return out


@function("AND", "raw", min_args=1, variadic=True, category=L)
def AND(ctx: Context, *args: Any) -> bool | ErrorValue:
    """TRUE when every logical is; ``#VALUE!`` when there are none."""
    values = _logicals(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    return all(values) if values else VALUE


@function("OR", "raw", min_args=1, variadic=True, category=L)
def OR(ctx: Context, *args: Any) -> bool | ErrorValue:
    """TRUE when any logical is; ``#VALUE!`` when there are none."""
    values = _logicals(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    return any(values) if values else VALUE


@function("NOT", "logical", category=L)
def NOT(ctx: Context, value: bool) -> bool:
    """The opposite logical."""
    return not value


@function("TRUE", category=L)
def TRUE(ctx: Context) -> bool:
    """TRUE."""
    return True


@function("FALSE", category=L)
def FALSE(ctx: Context) -> bool:
    """FALSE."""
    return False


@function("IFERROR", "branch", "branch", errors=True, category=L)
def IFERROR(ctx: Context, value: Any, alternative: Any) -> Any:
    """``alternative`` when ``value`` is an error."""
    v = scalar(value, ctx)
    if isinstance(v, ErrorValue):
        return 0.0 if alternative is MISSING else alternative
    return value


@function("_XLFN.LET", "raw", "raw", "raw", min_args=2, variadic=True, cycle=2, category=L)
def LET(ctx: Context, *args: Any) -> Any:
    """``LET``: names bound to values for a calculation, compiled by the code generator.

    Its names are resolved when the formula is compiled and its bindings evaluated only when
    read (``_codegen.let``; capture ``dyn_let``); this core is never called.
    """
    raise AssertionError("LET is compiled by the code generator")


@function("_XLFN.IFNA", "branch", "branch", errors=True, category=L)
def IFNA(ctx: Context, value: Any, alternative: Any) -> Any:
    """``alternative`` when ``value`` is ``#N/A``."""
    v = scalar(value, ctx)
    if isinstance(v, ErrorValue) and v.code == "#N/A":
        return 0.0 if alternative is MISSING else alternative
    return value


@function("ISNUMBER", "value", errors=True, category=INFO)
def ISNUMBER(ctx: Context, v: Any) -> bool:
    """Whether the value is a number."""
    return isinstance(v, float) and not isinstance(v, bool)


@function("ISTEXT", "value", errors=True, category=INFO)
def ISTEXT(ctx: Context, v: Any) -> bool:
    """Whether the value is text."""
    return isinstance(v, str)


@function("ISLOGICAL", "value", errors=True, category=INFO)
def ISLOGICAL(ctx: Context, v: Any) -> bool:
    """Whether the value is a logical."""
    return isinstance(v, bool)


@function("ISBLANK", "value", errors=True, category=INFO)
def ISBLANK(ctx: Context, v: Any) -> bool:
    """Whether the cell is empty (an empty text is not)."""
    return v is BLANK


@function("ISERROR", "value", errors=True, category=INFO)
def ISERROR(ctx: Context, v: Any) -> bool:
    """Whether the value is any error."""
    return isinstance(v, ErrorValue)


@function("ISERR", "value", errors=True, category=INFO)
def ISERR(ctx: Context, v: Any) -> bool:
    """Whether the value is an error other than ``#N/A``."""
    return isinstance(v, ErrorValue) and v.code != "#N/A"


@function("ISNA", "value", errors=True, category=INFO)
def ISNA(ctx: Context, v: Any) -> bool:
    """Whether the value is ``#N/A``."""
    return isinstance(v, ErrorValue) and v.code == "#N/A"


@function("NA", category=INFO)
def NA_(ctx: Context) -> ErrorValue:
    """``#N/A``."""
    return NA


@function(
    "_xlfn.IFS", "value", "branch", min_args=2, variadic=True, errors=True, cycle=2, category=L
)
def IFS(ctx: Context, *args: Any) -> Any:
    """The value after the first condition that holds; ``#N/A`` if none (element by element)."""
    if len(args) % 2:
        return VALUE
    for k in range(0, len(args), 2):
        t = to_bool(args[k], ctx)
        if isinstance(t, ErrorValue):
            return t
        if t:
            return args[k + 1]
    return NA


@function(
    "_xlfn.SWITCH",
    "value",
    "value",
    "branch",
    min_args=3,
    variadic=True,
    errors=True,
    cycle=2,
    category=L,
)
def SWITCH(ctx: Context, expression: Any, *pairs: Any) -> Any:
    """The result after the first value equal to ``expression``, or the default, or ``#N/A``."""
    v = scalar(expression, ctx)
    if isinstance(v, ErrorValue):
        return v
    n = len(pairs) - len(pairs) % 2
    for k in range(0, n, 2):
        candidate = scalar(pairs[k], ctx)
        if isinstance(candidate, ErrorValue):
            return candidate
        if switch_equal(v, candidate):
            return pairs[k + 1]
    return pairs[-1] if len(pairs) % 2 else NA


def _ignored_text(v: Any) -> bool:
    return isinstance(v, str) and v.upper() not in ("TRUE", "FALSE")


@function("_xlfn.XOR", "raw", min_args=1, variadic=True, category=L)
def XOR(ctx: Context, *args: Any) -> bool | ErrorValue:
    """TRUE when an odd number of logicals are; ``#VALUE!`` when there are none."""
    values = _logicals(tuple(a for a in args if not _ignored_text(a)), ctx)
    if isinstance(values, ErrorValue):
        return values
    return sum(values) % 2 == 1 if values else VALUE


def _parity(v: Any, ctx: Context) -> int | ErrorValue:
    """A number's whole part modulo 2, as the old Analysis ToolPak functions take a number."""
    if isinstance(v, bool):
        return VALUE
    x = to_number(v, ctx)
    return x if isinstance(x, ErrorValue) else int(math.fmod(math.trunc(x), 2))


@function("ISEVEN", "value", category=INFO)
def ISEVEN(ctx: Context, v: Any) -> bool | ErrorValue:
    """Whether a number's whole part is even."""
    odd = _parity(v, ctx)
    return odd if isinstance(odd, ErrorValue) else odd == 0


@function("ISODD", "value", category=INFO)
def ISODD(ctx: Context, v: Any) -> bool | ErrorValue:
    """Whether a number's whole part is odd."""
    odd = _parity(v, ctx)
    return odd if isinstance(odd, ErrorValue) else odd != 0


@function("ISNONTEXT", "value", errors=True, category=INFO)
def ISNONTEXT(ctx: Context, v: Any) -> bool:
    """Whether the value is anything but text."""
    return not isinstance(v, str)


@function("ISREF", "raw", errors=True, category=INFO)
def ISREF(ctx: Context, v: Any) -> bool:
    """Whether the argument is a reference."""
    return isinstance(v, Ref)


@function("TYPE", "raw", errors=True, category=INFO)
def TYPE(ctx: Context, v: Any) -> float:
    """The kind of a value: 1 number or blank, 2 text, 4 logical, 16 error, 64 array."""
    if isinstance(v, Ref):
        if not v.is_cell():
            return 16.0
        v = v.value(0, 0)
    if isinstance(v, tuple):
        return 64.0
    if isinstance(v, str):
        return 2.0
    if isinstance(v, bool):
        return 4.0
    return 16.0 if isinstance(v, ErrorValue) else 1.0


_ERROR_TYPES = {
    "#NULL!": 1.0, "#DIV/0!": 2.0, "#VALUE!": 3.0, "#REF!": 4.0, "#NAME?": 5.0, "#NUM!": 6.0,
    "#N/A": 7.0, "#GETTING_DATA": 8.0, "#SPILL!": 9.0, "#CALC!": 14.0,
}  # fmt: skip


@function("ERROR.TYPE", "value", errors=True, category=INFO)
def ERROR_TYPE(ctx: Context, v: Any) -> float | ErrorValue:
    """The number of an error; ``#N/A`` for anything else."""
    return _ERROR_TYPES.get(v.code, NA) if isinstance(v, ErrorValue) else NA
