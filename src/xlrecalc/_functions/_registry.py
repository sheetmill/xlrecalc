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
- ``function``: a function value (a ``Lambda``: ``MAP``'s last argument), untouched.

A function value given at any other kind's position never reaches the core: the call is
``#VALUE!``, except that the ``IS`` functions give FALSE (``ISNONTEXT`` TRUE) and ``TYPE`` 128
(captures ``lambda_calls``, ``lambda_more``). ``branch`` arguments pass one through, so ``IF``
and ``CHOOSE`` can choose one to call.

When an argument of a one-value kind (or a ``branch``) is an array, the function is applied to
each element and returns an array (§7.7).

The last kind repeats for variadic functions, or the last ``cycle`` kinds (``SUMIFS``'s range and
criterion). A function returning a reference (``reference``), given an array, makes an array of
references; a formula passing one to another function is reported, not computed (§7.7). A
function may have another core for an array context (``in_array``): ``ROW`` of a
range is its first row in a legacy formula, and all of them in an array context. Names are the
file form, upper case, with ``_xlfn.`` for functions newer than Excel 2007.

Results are compared with Excel's bit for bit. A function whose results still differ after a fixed
time spent fitting Excel's last bits declares a small tolerance in units in the last place
(``ulps``) with its reason (``ulps_reason``); the capture tests count rows within it separately
from exact matches, and ``QUIRKS.md`` lists it as a known deviation. A function may declare a
relative tolerance instead, or as well (``rel``, with ``rel_reason``), counted apart from both: the
iterative functions (IRR, XIRR, RATE, YIELD), whose result Excel itself guarantees only to its
documented stopping tolerance, at 1E-9; the beta, gamma, chi-squared, F and t distributions at
1E-15 to 1E-12; and the regression over several x columns (LINEST, LOGEST, TREND, GROWTH) at
5E-14.
"""

from __future__ import annotations

import inspect
import math
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from xlrecalc._arrays import NotImplementedFormError, Padded, broadcast, is_array, to_array
from xlrecalc._coerce import Context, scalar, to_bool, to_number, to_text
from xlrecalc._formula import MISSING
from xlrecalc._formula._codegen import kind_at
from xlrecalc._values import Areas, ErrorValue, Lambda, Ref

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
    makes_array: bool = False


REGISTRY: dict[str, Callable[..., Any]] = {}
SPECS: dict[str, Spec] = {}


# What an empty argument does differs by function (capture empty_args). These give #N/A for an
# empty required argument, whatever the others hold: mostly the old Analysis ToolPak functions.
_EMPTY_REQUIRED_IS_NA = frozenset(
    [
        "NETWORKDAYS.INTL",
        "_XLFN.NETWORKDAYS.INTL",
        "BESSELI",
        "BESSELJ",
        "BESSELK",
        "BESSELY",
        "BIN2HEX",
        "CONVERT",
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
        "COMPLEX",
        "IMDIV",
        "IMPOWER",
        "IMPRODUCT",
        "IMSUB",
        "IMSUM",
    ]
)
# These give #VALUE! for an empty argument where a range belongs, at these positions.
# These give #N/A for any empty argument, an optional one too (capture empty_args: ERF(2,)).
_EMPTY_ANY_IS_NA = frozenset(["ERF"])
_EMPTY_IS_VALUE: dict[str, tuple[int, ...]] = {
    "INDEX": (0,), "IRR": (0,), "LINEST": (0,), "LOOKUP": (1, 2), "MIRR": (0,), "PROB": (0, 1),
    "_XLFN.XMATCH": (1,),
    "SUMX2MY2": (0, 1), "SUMX2PY2": (0, 1), "SUMXMY2": (0, 1),
    **dict.fromkeys(("CHITEST", "_XLFN.CHISQ.TEST", "FTEST", "_XLFN.F.TEST"), (0, 1)),
    **dict.fromkeys(("TTEST", "_XLFN.T.TEST"), (0, 1)),
}  # fmt: skip
# These take an empty optional argument as if it were left out. Others take it as 0 or FALSE.
_EMPTY_OPTIONAL_IS_DEFAULT = frozenset(
    [
        "_XLFN.BINOM.DIST.RANGE",
        "BETADIST",
        "BETAINV",
        "_XLFN.BETA.DIST",
        "_XLFN.BETA.INV",
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
        "LOGEST",
        "TREND",
        "GROWTH",
        "OCT2BIN",
        "OCT2HEX",
        "PERCENTRANK",
        "WEEKNUM",
        "WORKDAY.INTL",
        "XIRR",
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

# What these give for a function value, where any other function gives #VALUE! (captures
# lambda_calls, lambda_more).
_GIVEN_LAMBDA: dict[str, Any] = {
    "ISERROR": False, "ISERR": False, "ISNA": False, "ISREF": False, "ISBLANK": False,
    "ISNUMBER": False, "ISTEXT": False, "ISLOGICAL": False, "ISNONTEXT": True, "TYPE": 128.0,
}  # fmt: skip
_TAKES_LAMBDA = frozenset({"branch", "function"})


def _given_lambda(name: str, kinds: tuple[str, ...], cycle: int, args: tuple[Any, ...]) -> Any:
    """The result of a call given a function value where it takes none, else ``None``."""
    for i, a in enumerate(args):
        if type(a) is Lambda and (not kinds or kind_at(kinds, i, cycle) not in _TAKES_LAMBDA):
            return _GIVEN_LAMBDA.get(name, VALUE)
    return None


# What each function does with a reference of several areas (captures three_d and unions, both
# platforms), one table for a 3-D reference and one for a union: "expand", read each area as an
# argument of its own (the lists of values); "stack", read the cells of every area as one array
# (an argument that takes one range); "pass", take the areas as they are (IF and CHOOSE return
# their branch; RANK reads every area); "first", read the first area's first cell (N, T,
# FORMULATEXT of a union); "index", pick the area INDEX's area number names; or the result
# whatever the areas hold. A function not in the table of its form was not captured with it:
# its cell is reported as not computed.
_AREAS: dict[str, Any] = {
    **dict.fromkeys(
        (
            "SUM", "AVERAGE", "AVERAGEA", "COUNT", "COUNTA", "MAX", "MAXA", "MIN", "MINA",
            "PRODUCT", "STDEV", "_XLFN.STDEV.S", "_XLFN.STDEV.P", "STDEVA", "STDEVP", "STDEVPA",
            "VAR", "_XLFN.VAR.S", "_XLFN.VAR.P", "VARA", "VARP", "VARPA", "MEDIAN", "AVEDEV",
            "DEVSQ", "GEOMEAN", "HARMEAN", "KURT", "SKEW", "SUMSQ", "_XLFN.CONCAT",
            "_XLFN.TEXTJOIN",
        ),
        "expand",
    ),
    **dict.fromkeys(("LARGE", "SMALL", "PERCENTILE", "QUARTILE"), "stack"),
    "IF": "pass",
    "RANK": "pass",
    **dict.fromkeys(
        (
            "COUNTBLANK", "SUBTOTAL", "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTIFS",
            "_XLFN.MAXIFS", "VLOOKUP", "HLOOKUP", "MATCH", "INDEX", "LOOKUP", "_XLFN.XLOOKUP",
            "ROWS", "COLUMNS", "ROW", "COLUMN", "AREAS", "N", "T", "CELL", "_XLFN.FORMULATEXT",
            "OFFSET",
        ),
        ErrorValue.VALUE,
    ),
    "MODE": ErrorValue.REF,
    "SUMPRODUCT": ErrorValue.REF,
    "ISREF": False,
    "ISBLANK": False,
    "_XLFN.SHEETS": lambda areas: float(len(areas.refs)),
    "_XLFN.SHEET": lambda areas: float(areas.refs[0].sheet + 1),
}  # fmt: skip


_UNION_AREAS: dict[str, Any] = {
    **dict.fromkeys(
        (
            "SUM", "AVERAGE", "AVERAGEA", "COUNT", "COUNTA", "MAX", "MAXA", "MIN", "MINA",
            "PRODUCT", "STDEV", "_XLFN.STDEV.S", "_XLFN.STDEV.P", "STDEVA", "STDEVP", "STDEVPA",
            "VAR", "_XLFN.VAR.S", "_XLFN.VAR.P", "VARA", "VARP", "VARPA", "MEDIAN", "AVEDEV",
            "DEVSQ", "GEOMEAN", "HARMEAN", "KURT", "SKEW", "SUMSQ", "SUBTOTAL",
        ),
        "expand",
    ),
    **dict.fromkeys(("LARGE", "SMALL", "PERCENTILE", "QUARTILE"), "stack"),
    **dict.fromkeys(("IF", "CHOOSE", "RANK"), "pass"),
    **dict.fromkeys(("N", "T", "_XLFN.FORMULATEXT", "_XLFN.ISFORMULA"), "first"),
    "INDEX": "index",
    **dict.fromkeys(
        (
            "COUNTBLANK", "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTIFS", "SUMIFS", "_XLFN.MAXIFS",
            "MODE", "SUMPRODUCT", "_XLFN.CONCAT", "_XLFN.TEXTJOIN", "OFFSET",
        ),
        ErrorValue.VALUE,
    ),
    **dict.fromkeys(
        ("VLOOKUP", "HLOOKUP", "MATCH", "LOOKUP", "_XLFN.XLOOKUP"), ErrorValue.NA
    ),
    **dict.fromkeys(("ROWS", "COLUMNS", "ROW", "COLUMN"), ErrorValue.REF),
    "ISREF": True,
    "ISBLANK": False,
    # as the value #VALUE! (capture info_family)
    "ISERROR": True, "ISNONTEXT": True, "TYPE": 16.0, "ERROR.TYPE": 3.0,
    "ISEVEN": ErrorValue.VALUE, "ISODD": ErrorValue.VALUE,
    "AREAS": lambda areas: float(len(areas.refs)),
    "_XLFN.SHEETS": 1.0,
    "_XLFN.SHEET": lambda areas: float(areas.refs[0].sheet + 1),
}  # fmt: skip


def _with_areas(name: str, args: tuple[Any, ...], ctx: Context) -> tuple[Any, ...] | Any:
    """The arguments a function reads given a reference of several areas, or its result.

    Returns a tuple of arguments, or the result itself (any other value).

    Raises:
        NotImplementedFormError: The function was not captured with such a reference.
    """
    areas = next(a for a in args if type(a) is Areas)
    rule = (_UNION_AREAS if areas.union else _AREAS).get(name)
    if rule is None:
        form = "a union of ranges" if areas.union else "a 3-D reference"
        raise NotImplementedFormError(f"{form} in {name}")
    if rule == "pass":
        return args
    if rule == "expand":
        return tuple(r for a in args for r in (a.refs if type(a) is Areas else (a,)))
    if rule == "stack":
        return tuple(
            tuple(row for ref in a.refs for row in to_array(ref)) if type(a) is Areas else a
            for a in args
        )
    if rule == "first":
        return tuple(_first_cell(a) if type(a) is Areas else a for a in args)
    if rule == "index":
        return _index_area(args, ctx)
    if callable(rule):
        return rule(areas)
    return rule


def _first_cell(areas: Areas) -> Ref:
    first = areas.refs[0]
    return Ref(first.source, first.sheet, first.r1, first.c1, first.r1, first.c1)


def _index_area(args: tuple[Any, ...], ctx: Context) -> tuple[Any, ...] | ErrorValue:
    """``INDEX``'s arguments with the area its area number picks (capture unions).

    The number is truncated and is 1 when left out; below 1 it is ``#VALUE!``, past the last
    area ``#REF!``.
    """
    if type(args[0]) is not Areas:
        return VALUE
    given = args[3] if len(args) > 3 and args[3] is not MISSING else 1.0
    n = to_number(given, ctx)
    if isinstance(n, ErrorValue):
        return n
    k = math.trunc(n)
    if k < 1:
        return VALUE
    if k > len(args[0].refs):
        return ErrorValue.REF
    return (args[0].refs[k - 1], *args[1:3], 1.0)


def _map_padded(call: Callable[..., Any], args: tuple[Any, ...], pads: list[int]) -> Padded:
    """``call`` element by element over padded arguments of one size.

    The other arguments are passed as they are. Each position of the rows with data (to the
    tallest of them) is one call, and the fills one more.
    """
    first = args[pads[0]]
    height = max(len(args[i].head) for i in pads)

    def at(r: int, c: int) -> list[Any]:
        out = list(args)
        for i in pads:
            a = args[i]
            out[i] = a.head[r][c] if r < len(a.head) else a.fill
        return out

    def one(xs: list[Any]) -> Any:
        v = call(*xs)
        return v[0][0] if is_array(v) else v

    head = tuple(tuple(one(at(r, c)) for c in range(first.cols)) for r in range(height))
    fills = list(args)
    for i in pads:
        fills[i] = args[i].fill
    return Padded(head, first.rows, first.cols, one(fills))


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
    makes_array: bool = False,
    padded: bool = False,
    elementwise: bool = False,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register a function under its file-form name.

    ``makes_array``: the function makes an array (or, as ``INDIRECT``, may make a range), so a
    formula whose text carries a call of it to the result is written with the dynamic-array
    marker whatever the size of its result, as Excel writes it typed (captures ``marker_typed``
    and ``dyn_*``; the rule is ``_workbook._may_be_array``).

    ``padded``: the function reads a :class:`~xlrecalc._arrays.Padded` array in its ``array``
    arguments, so the compiler may give it one there (ARCHITECTURE.md §7.7).

    ``elementwise``: the function, given an array in its lifted argument, is applied to each
    element and nothing else; given a padded array there, it maps the rows and the fill, and the
    compiler lets one through only where its result reaches a reader of padded arrays.
    """
    if (ulps and not ulps_reason) or (rel and not rel_reason):
        raise ValueError(f"{name}: a tolerance needs its reason")
    if status not in ("implemented", "partial") or (status == "partial") != bool(note):
        raise ValueError(f"{name}: status 'implemented', or 'partial' with a note")
    lo = len(kinds) if min_args is None else min_args
    hi = None if variadic else len(kinds)
    upper = name.upper()

    def wrap(core: Callable[..., Any]) -> Callable[..., Any]:
        lifted_kinds = [k in _LIFTED for k in kinds]
        plain = not any(lifted_kinds) and all(_CONVERT.get(k) is None for k in kinds)

        required_na = name.upper() in _EMPTY_REQUIRED_IS_NA
        any_na = name.upper() in _EMPTY_ANY_IS_NA
        value_at = _EMPTY_IS_VALUE.get(name.upper(), ())
        defaults = (
            [p.default for p in list(inspect.signature(core).parameters.values())[1:]]
            if name.upper() in _EMPTY_OPTIONAL_IS_DEFAULT
            else None
        )

        def call(ctx: Context, *args: Any) -> Any:
            if Areas in map(type, args):
                given = _with_areas(upper, args, ctx)
                if type(given) is not tuple:
                    return given
                args = given
            if Lambda in map(type, args):
                given = _given_lambda(upper, kinds, cycle, args)
                if given is not None:
                    return given
            n = len(args)
            if n < lo or (hi is not None and n > hi):
                return VALUE
            if (required_na or any_na) and any(a is MISSING for a in args[: n if any_na else lo]):
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
            if any(lifted[i] and type(a) is Padded for i, a in enumerate(args)):
                pads = [i for i, a in enumerate(args) if lifted[i] and type(a) is Padded]
                arrays = [
                    i
                    for i, a in enumerate(args)
                    if lifted[i] and (type(a) is tuple or (isinstance(a, Ref) and not a.is_cell()))
                ]
                sizes = {(args[i].rows, args[i].cols) for i in pads}
                if elementwise and len(sizes) == 1 and not arrays:  # the rows, then the fill once
                    return _map_padded(lambda *xs: call(ctx, *xs), args, pads)
                limit = getattr(ctx, "max_array_cells", None)
                args = tuple(a.dense(limit) if type(a) is Padded else a for a in args)
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
        call.padded = padded  # type: ignore[attr-defined]
        call.elementwise = elementwise  # type: ignore[attr-defined]
        call.__name__ = core.__name__
        call.__doc__ = core.__doc__
        return call

    def register(core: Callable[..., Any]) -> Callable[..., Any]:
        call = wrap(core)
        if in_array is not None:
            call.in_array = wrap(in_array)  # type: ignore[attr-defined]
        REGISTRY[upper] = call
        SPECS[upper] = Spec(
            upper, kinds, lo, hi, errors, volatile, category, ulps, ulps_reason, status, note,
            rel, rel_reason, makes_array,
        )  # fmt: skip
        return core

    return register
