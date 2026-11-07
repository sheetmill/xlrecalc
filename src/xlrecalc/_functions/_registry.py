"""The function registry (ARCHITECTURE.md §8.1).

``@function("ROUND", "number", "number")`` registers a function with the kinds of its
arguments; the wrapper checks the argument count, coerces each argument, and returns the first
error among scalar arguments unless the function inspects errors itself (``errors=True``). The
core function receives the evaluation context first, then the coerced arguments:

- ``number``, ``text``, ``logical``: one value converted (``_coerce.py``), an omitted argument
  as 0, ``""`` or FALSE;
- ``value``: one value as it is (references intersected, an omitted argument blank);
- ``raw``: the argument untouched, a ``Ref`` for references, ``MISSING`` when omitted;
- ``array``: untouched too, and evaluated as an array even in a legacy formula, where an
  operator over a range would otherwise take its implicit intersection (``SUMPRODUCT``'s
  arguments, ``INDEX``'s array: capture ``arrays``);
- ``branch``: untouched, like ``raw``, but taken element by element when the call is (``IF``'s
  results);
- ``matrix``: a whole array, never taken element by element; a reference in a legacy formula
  is intersected, as for ``value``, and in an array context it arrives as an array
  (``TRANSPOSE``: capture ``misc``).
- ``key``: one value in a legacy formula, inside ``SUMPRODUCT`` too (an array's first element, a
  reference intersected), and taken element by element in an array formula (``VLOOKUP``'s key:
  capture ``array_refs``).

When an argument of a one-value kind (or a ``branch``) is an array, the function is applied to
each element and returns an array (§7.7).

The last kind repeats for variadic functions, or the last ``cycle`` kinds (``SUMIFS``'s range and
criterion). A function returning a reference (``reference``), given an array, makes an array of
references; a formula passing one to another function is reported, not computed (§7.7). A
function may have another core for an array context (``in_array``): ``ROW`` of a
range is its first row in a legacy formula, and all of them in an array context. Names are the
file form, upper case, with ``_xlfn.`` for functions newer than Excel 2007.

Results are compared with Excel's bit for bit. A function whose results could not be made
bit-exact within a time-boxed attempt declares a small tolerance in units in the last place
(``ulps``) with its reason (``ulps_reason``); the capture tests count rows within it separately
from exact matches, and ``QUIRKS.md`` lists it as a known deviation. An iterative function
(IRR, XIRR, RATE, YIELD), whose result Excel itself guarantees only to its documented stopping
tolerance, may instead declare a relative tolerance (``rel``, 1E-9, with ``rel_reason``),
counted apart from both.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from xlrecalc._arrays import broadcast, to_array
from xlrecalc._coerce import Context, scalar, to_bool, to_number, to_text
from xlrecalc._formula import MISSING
from xlrecalc._formula._codegen import kind_at
from xlrecalc._values import ErrorValue, Ref

__all__ = ["REGISTRY", "SPECS", "Spec", "function"]

VALUE = ErrorValue.VALUE

_LIFTED = frozenset({"number", "text", "logical", "value", "branch", "key"})
_CONVERT: dict[str, Callable[[Any, Context], Any]] = {
    "number": to_number,
    "text": to_text,
    "logical": to_bool,
    "value": scalar,
    "key": scalar,
    "matrix": lambda a, ctx: scalar(a, ctx) if isinstance(a, Ref) else a,
}


@dataclass(frozen=True, slots=True)
class Spec:
    """A registered function: its kinds, argument counts and flags."""

    name: str
    kinds: tuple[str, ...]
    min_args: int
    max_args: int | None
    errors: bool
    volatile: bool
    category: str
    ulps: int = 0
    ulps_reason: str = ""
    status: str = "implemented"
    note: str = ""
    rel: float = 0.0
    rel_reason: str = ""


REGISTRY: dict[str, Callable[..., Any]] = {}
SPECS: dict[str, Spec] = {}


# What an empty argument does differs by function (capture empty_args). These give #N/A for an
# empty required argument, whatever the others hold: mostly the old Analysis ToolPak functions.
_EMPTY_REQUIRED_IS_NA = frozenset(
    [
        "NETWORKDAYS.INTL",
        "_XLFN.NETWORKDAYS.INTL",
        "BIN2HEX",
        "BIN2OCT",
        "COUPDAYBS",
        "COUPDAYS",
        "COUPDAYSNC",
        "COUPNCD",
        "COUPNUM",
        "COUPPCD",
        "CUMIPMT",
        "CUMPRINC",
        "DEC2BIN",
        "DEC2HEX",
        "DEC2OCT",
        "DELTA",
        "DISC",
        "DOLLARDE",
        "DOLLARFR",
        "DURATION",
        "EDATE",
        "EFFECT",
        "EOMONTH",
        "FVSCHEDULE",
        "GESTEP",
        "HEX2BIN",
        "HEX2OCT",
        "INTRATE",
        "MDURATION",
        "MROUND",
        "NETWORKDAYS",
        "NOMINAL",
        "OCT2BIN",
        "OCT2HEX",
        "PRICE",
        "PRICEDISC",
        "QUOTIENT",
        "RECEIVED",
        "SERIESSUM",
        "TBILLEQ",
        "TBILLPRICE",
        "TBILLYIELD",
        "WEEKNUM",
        "WORKDAY",
        "WORKDAY.INTL",
        "XIRR",
        "XNPV",
        "YEARFRAC",
        "YIELD",
        "YIELDDISC",
        "_XLFN.WORKDAY.INTL",
    ]
)
# These give #VALUE! for an empty argument where a range belongs, at these positions.
_EMPTY_IS_VALUE: dict[str, tuple[int, ...]] = {
    "INDEX": (0,), "IRR": (0,), "LINEST": (0,), "LOOKUP": (1, 2), "MIRR": (0,), "PROB": (0, 1),
    "_XLFN.XMATCH": (1,),
    "SUMX2MY2": (0, 1), "SUMX2PY2": (0, 1), "SUMXMY2": (0, 1),
}  # fmt: skip
# These take an empty optional argument as if it were left out. Others take it as 0 or FALSE.
_EMPTY_OPTIONAL_IS_DEFAULT = frozenset(
    [
        "PROB",
        "ADDRESS",
        "BIN2HEX",
        "BIN2OCT",
        "DEC2BIN",
        "DEC2HEX",
        "DEC2OCT",
        "HEX2BIN",
        "HEX2OCT",
        "ISO.CEILING",
        "LINEST",
        "OCT2BIN",
        "OCT2HEX",
        "PERCENTRANK",
        "WEEKNUM",
        "WORKDAY.INTL",
        "_XLFN.CEILING.MATH",
        "_XLFN.CEILING.PRECISE",
        "_XLFN.FLOOR.MATH",
        "_XLFN.FLOOR.PRECISE",
        "_XLFN.PERCENTRANK.EXC",
        "_XLFN.PERCENTRANK.INC",
        "_XLFN.TEXTAFTER",
        "_XLFN.TEXTBEFORE",
        "_XLFN.WORKDAY.INTL",
    ]
)


def function(
    name: str,
    *kinds: str,
    min_args: int | None = None,
    variadic: bool = False,
    errors: bool = False,
    volatile: bool = False,
    category: str = "",
    ulps: int = 0,
    ulps_reason: str = "",
    status: str = "implemented",
    note: str = "",
    rel: float = 0.0,
    rel_reason: str = "",
    cycle: int = 1,
    in_array: Callable[..., Any] | None = None,
    reference: bool = False,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register a function under its file-form name."""
    if (ulps and not ulps_reason) or (rel and not rel_reason):
        raise ValueError(f"{name}: a tolerance needs its reason")
    if status not in ("implemented", "partial") or (status == "partial") != bool(note):
        raise ValueError(f"{name}: status 'implemented', or 'partial' with a note")
    lo = len(kinds) if min_args is None else min_args
    hi = None if variadic else len(kinds)

    def wrap(core: Callable[..., Any]) -> Callable[..., Any]:
        lifted_kinds = [k in _LIFTED for k in kinds]
        plain = not any(lifted_kinds) and all(_CONVERT.get(k) is None for k in kinds)

        required_na = name.upper() in _EMPTY_REQUIRED_IS_NA
        value_at = _EMPTY_IS_VALUE.get(name.upper(), ())
        defaults = (
            [p.default for p in list(inspect.signature(core).parameters.values())[1:]]
            if name.upper() in _EMPTY_OPTIONAL_IS_DEFAULT
            else None
        )

        def call(ctx: Context, *args: Any) -> Any:
            n = len(args)
            if n < lo or (hi is not None and n > hi):
                return VALUE
            if required_na and any(a is MISSING for a in args[:lo]):
                return ErrorValue.NA
            if value_at and any(i < n and args[i] is MISSING for i in value_at):
                return VALUE
            if defaults is not None:
                args = tuple(
                    defaults[i]
                    if a is MISSING and i >= lo and defaults[i] is not inspect.Parameter.empty
                    else a
                    for i, a in enumerate(args)
                )
            if plain:  # every argument passed as it is (SUM, COUNT, ...): nothing to coerce
                return core(ctx, *args)
            lifted = [kind_at(kinds, i, cycle) in _LIFTED for i in range(n)]
            if any(lifted[i] and type(a) is tuple for i, a in enumerate(args)):
                spread = [
                    to_array(a) if lifted[i] and isinstance(a, Ref) and not a.is_cell() else a
                    for i, a in enumerate(args)
                ]
                limit = getattr(ctx, "max_array_cells", None)
                result = broadcast(lambda *xs: call(ctx, *xs), spread, lifted, limit)
                if reference:  # an array of references: only the cell itself may take one
                    ctx.reference_array = result  # type: ignore[attr-defined]
                return result
            coerced = []
            for i, arg in enumerate(args):
                convert = _CONVERT.get(kind_at(kinds, i, cycle))
                if convert is None or (
                    defaults is not None and i >= lo and (arg is MISSING or arg is None)
                ):
                    coerced.append(arg)
                    continue
                value = convert(arg, ctx)
                if isinstance(value, ErrorValue) and not errors:
                    return value
                coerced.append(value)
            return core(ctx, *coerced)

        call.kinds = kinds  # type: ignore[attr-defined]
        call.cycle = cycle  # type: ignore[attr-defined]
        call.__name__ = core.__name__
        call.__doc__ = core.__doc__
        return call

    def register(core: Callable[..., Any]) -> Callable[..., Any]:
        call = wrap(core)
        if in_array is not None:
            call.in_array = wrap(in_array)  # type: ignore[attr-defined]
        upper = name.upper()
        REGISTRY[upper] = call
        SPECS[upper] = Spec(
            upper, kinds, lo, hi, errors, volatile, category, ulps, ulps_reason, status, note,
            rel, rel_reason,
        )  # fmt: skip
        return core

    return register
