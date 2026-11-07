"""Math and aggregation: SUM, ROUND, INT, MOD, SUMIF, SUMPRODUCT, SUBTOTAL and friends."""

from __future__ import annotations

import math
import random
import statistics
from collections.abc import Callable
from decimal import ROUND_DOWN, ROUND_HALF_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal
from fractions import Fraction
from typing import Any

from xlrecalc._arith import flush, snap
from xlrecalc._arrays import check_area, element, is_array, shape, to_array
from xlrecalc._coerce import Context, numbers, scalar, to_number
from xlrecalc._compare import criterion
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._numtext import DBL_MIN
from xlrecalc._store import MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Ref

NUM, DIV0, VALUE = ErrorValue.NUM, ErrorValue.DIV0, ErrorValue.VALUE
C = "math"


def _total(values: Any, ctx: Context) -> float | ErrorValue:
    """A sum as Excel's SUM takes it: in order, each addition snapped (§4.4)."""
    s = 0.0
    for v in values:
        if isinstance(v, ErrorValue):
            return v
        r = s + v
        s = snap(s, r, ctx.quirks) if s else flush(r)
    return s if math.isfinite(s) else NUM


@function("SUM", "raw", min_args=1, variadic=True, category=C)
def SUM(ctx: Context, *args: Any) -> float | ErrorValue:
    """Sum of numbers."""
    if len(args) == 1 and type(args[0]) is Ref:  # the common case, as _total in one loop
        snapping = ctx.quirks.cancellation_snapping
        s = 0.0
        for v in args[0].values():
            if type(v) is float:
                r = s + v
                if s and snapping and r and abs(r) < 8 * math.ulp(s):
                    s = 0.0
                else:
                    s = 0.0 if -DBL_MIN < r < DBL_MIN else r
            elif type(v) is ErrorValue:
                return v
        return s if math.isfinite(s) else NUM
    return _total(numbers(args, ctx), ctx)


@function("PRODUCT", "raw", min_args=1, variadic=True, category=C)
def PRODUCT(ctx: Context, *args: Any) -> float | ErrorValue:
    """Product of numbers; 0 when there are none."""
    p, seen = 1.0, False
    for v in numbers(args, ctx, skip_empty=True):
        if isinstance(v, ErrorValue):
            return v
        p *= v
        seen = True
    return (flush(p) if math.isfinite(p) else NUM) if seen else 0.0


def _decimal(x: float) -> Decimal:
    """The 15-significant-digit decimal a number shows, which Excel's rounding works on."""
    if x == 0:
        return Decimal(0)
    d = Decimal(x)
    return d.quantize(Decimal(1).scaleb(d.adjusted() - 14), rounding=ROUND_HALF_DOWN)


def _round(x: float, digits: float, mode: str) -> float | ErrorValue:
    n = math.trunc(digits)
    try:
        q = Decimal(1).scaleb(-n)
        result = float(_decimal(x).quantize(q, rounding=mode))
    except ArithmeticError:
        return x if n > 0 else 0.0
    return result if math.isfinite(result) else NUM


@function("ROUND", "number", "number", category=C)
def ROUND(ctx: Context, x: float, digits: float) -> float | ErrorValue:
    """Round half away from zero, on the decimal value (``ROUND(2.675,2)`` is 2.68)."""
    return _round(x, digits, ROUND_HALF_UP)


@function("ROUNDUP", "number", "number", category=C)
def ROUNDUP(ctx: Context, x: float, digits: float) -> float | ErrorValue:
    """Round away from zero.

    Excel computes it as ROUNDDOWN plus one unit at the last kept place, added in doubles,
    when digits are dropped: ``ROUNDUP(2.675,2)`` is ``2.67 + 0.01``, 2.6799999999999997
    (capture ``rounding``, every row).
    """
    n = math.trunc(digits)
    down = _round(x, digits, ROUND_DOWN)
    if isinstance(down, ErrorValue):
        return down
    try:
        exact = _decimal(x) == _decimal(x).quantize(Decimal(1).scaleb(-n), rounding=ROUND_DOWN)
        unit = 10.0**-n
    except (ArithmeticError, OverflowError):
        return _round(x, digits, ROUND_UP)
    if exact:
        return down
    result = down + unit if x > 0 else down - unit
    return result if math.isfinite(result) else NUM


@function("ROUNDDOWN", "number", "number", category=C)
def ROUNDDOWN(ctx: Context, x: float, digits: float) -> float | ErrorValue:
    """Round toward zero."""
    return _round(x, digits, ROUND_DOWN)


@function("INT", "number", category=C)
def INT(ctx: Context, x: float) -> float:
    """Round down to an integer."""
    return float(math.floor(x))


@function("ABS", "number", category=C)
def ABS(ctx: Context, x: float) -> float:
    """Absolute value."""
    return abs(x)


@function("SQRT", "number", category=C)
def SQRT(ctx: Context, x: float) -> float | ErrorValue:
    """Square root; ``#NUM!`` for a negative number."""
    return NUM if x < 0 else math.sqrt(x)


@function("LN", "number", category=C)
def LN(ctx: Context, x: float) -> float | ErrorValue:
    """Natural logarithm; ``#NUM!`` unless positive."""
    return NUM if x <= 0 else math.log(x)


@function("EXP", "number", category=C)
def EXP(ctx: Context, x: float) -> float | ErrorValue:
    """E to the power x."""
    try:
        return flush(math.exp(x))
    except OverflowError:
        return NUM


@function("MOD", "number", "number", category=C)
def MOD(ctx: Context, n: float, d: float) -> float | ErrorValue:
    """Remainder with the sign of the divisor, computed exactly and rounded once."""
    if d == 0:
        return DIV0
    # Exactly, then rounded once (MOD(-8,1/3) is 0.33333333333333287, not 0).
    fn, fd = Fraction(n), Fraction(d)
    return flush(float(fn - fd * math.floor(fn / fd)))


@function("POWER", "number", "number", category=C)
def POWER(ctx: Context, x: float, y: float) -> float | ErrorValue:
    """X to the power y."""
    return power(x, y)


def power(x: float, y: float) -> float | ErrorValue:
    """``x^y`` as Excel computes it.

    ``0^0`` is ``#NUM!`` and ``0`` to a negative power ``#DIV/0!``; a negative number to a
    fractional power is its odd root when ``1/y`` is an odd integer (``(-8)^(1/3)`` is
    -1.9999999999999998, computed as ``-exp(y*ln|x|)``), else ``#NUM!``.
    """
    if x == 0 and y <= 0:
        return NUM if y == 0 else DIV0
    if x < 0 and y != math.floor(y):
        inverse = 1 / y
        odd = round(inverse)
        if abs(inverse - odd) > 1e-9 * abs(inverse) or odd % 2 == 0:
            return NUM
        try:
            r = -math.exp(y * math.log(-x))
        except OverflowError:
            return NUM
        return flush(r) if math.isfinite(r) else NUM
    try:
        r = math.pow(x, y)
    except OverflowError:
        return NUM
    return flush(r) if math.isfinite(r) else NUM


@function("PI", category=C)
def PI(ctx: Context) -> float:
    """Pi."""
    return math.pi


def _numeric(v: Any) -> bool:
    """What the ``SUMIF`` family adds: numbers and errors, not text or logicals."""
    return isinstance(v, float | ErrorValue) and not isinstance(v, bool)


@function("SUMIF", "raw", "value", "raw", min_args=2, errors=True, category=C)
def SUMIF(ctx: Context, ref: Any, crit: Any, sums: Any = MISSING) -> float | ErrorValue:
    """Sum of the cells of ``sums`` (or ``ref``) whose cell in ``ref`` meets the criterion.

    ``sums`` takes the shape of ``ref`` from its top-left cell, as Excel does, cut off at the
    sheet's edge. Only non-blank cells are visited, so the cost follows the cells a sheet holds,
    not a range's size (§7.7).
    """
    if not isinstance(ref, Ref):
        return VALUE
    target = ref
    if isinstance(sums, Ref):
        target = _sized_like(ref, sums)
    elif sums is not MISSING:
        return VALUE
    test = criterion(scalar(crit, ctx), ctx.quirks)
    if test(BLANK):  # a blank cell of ref matches, and a blank cell of target adds nothing
        ref.check()
        picked = (v for r, c, v in target.cells() if _numeric(v) and test(ref.value(r, c)))
    else:  # only non-blank cells of ref can match
        rows, cols = target.rows, target.columns
        cells = ref.cells()
        picked = (target.value(r, c) for r, c, v in cells if r < rows and c < cols and test(v))
    return _total((v for v in picked if _numeric(v)), ctx)


def _conditions(
    ctx: Context, pairs: tuple[Any, ...]
) -> list[tuple[Ref, Callable[[Any], bool]]] | ErrorValue:
    """The (range, test) of every (range, criterion) pair, all of one shape, or ``#VALUE!``.

    The ranges are checked as a read of all their cells would check them (§6.3), though the walks
    below read only some of those cells.
    """
    if not pairs or len(pairs) % 2 or not isinstance(pairs[0], Ref):
        return VALUE
    rows, cols = pairs[0].rows, pairs[0].columns
    conditions = []
    for k in range(0, len(pairs), 2):
        ref = pairs[k]
        if not isinstance(ref, Ref) or ref.rows != rows or ref.columns != cols:
            return VALUE
        ref.check()
        conditions.append((ref, criterion(scalar(pairs[k + 1], ctx), ctx.quirks)))
    return conditions


@function(
    "SUMIFS", "raw", "raw", "value", min_args=3, variadic=True, errors=True, cycle=2, category=C
)
def SUMIFS(ctx: Context, sums: Any, *pairs: Any) -> float | ErrorValue:
    """Sum of ``sums`` where every range meets its criterion.

    Only the non-blank cells of ``sums`` are visited: a blank one adds nothing (§7.7).
    """
    if not isinstance(sums, Ref):
        return VALUE
    conditions = _conditions(ctx, pairs)
    if isinstance(conditions, ErrorValue):
        return conditions
    if sums.rows != pairs[0].rows or sums.columns != pairs[0].columns:
        return VALUE
    picked = (
        v
        for r, c, v in sums.cells()
        if _numeric(v) and all(test(ref.value(r, c)) for ref, test in conditions)
    )
    return _total(picked, ctx)


def _picked(sums: Ref, conditions: list[tuple[Ref, Callable[[Any], bool]]]) -> Any:
    """The numbers and errors of ``sums`` where every condition holds, in order, cells only."""
    return (
        v
        for r, c, v in sums.cells()
        if _numeric(v) and all(test(ref.value(r, c)) for ref, test in conditions)
    )


def _sized_like(ref: Ref, sums: Ref) -> Ref:
    """``sums`` given ``ref``'s shape from its top-left cell, cut off at the sheet's edge."""
    return Ref(sums.source, sums.sheet, sums.r1, sums.c1,
               min(sums.r1 + ref.rows - 1, MAX_ROW),
               min(sums.c1 + ref.columns - 1, MAX_COL))  # fmt: skip


def _mean(values: list[Any], ctx: Context) -> float | ErrorValue:
    if not values:
        return DIV0
    total = _total(values, ctx)
    return total if isinstance(total, ErrorValue) else total / len(values)


def _ifs_values(ctx: Context, sums: Any, pairs: tuple[Any, ...]) -> list[Any] | ErrorValue:
    """What ``SUMIFS``'s relatives read: ``sums`` where every range meets its criterion."""
    if not isinstance(sums, Ref):
        return VALUE
    conditions = _conditions(ctx, pairs)
    if isinstance(conditions, ErrorValue):
        return conditions
    if sums.rows != pairs[0].rows or sums.columns != pairs[0].columns:
        return VALUE
    return list(_picked(sums, conditions))


@function("AVERAGEIF", "raw", "value", "raw", min_args=2, errors=True, category=C)
def AVERAGEIF(ctx: Context, ref: Any, crit: Any, cells: Any = MISSING) -> float | ErrorValue:
    """Mean of the cells of ``cells`` (or ``ref``) whose cell in ``ref`` meets the criterion."""
    if not isinstance(ref, Ref):
        return VALUE
    target = ref
    if isinstance(cells, Ref):
        target = _sized_like(ref, cells)
    elif cells is not MISSING:
        return VALUE
    test = criterion(scalar(crit, ctx), ctx.quirks)
    ref.check()
    return _mean(list(_picked(target, [(ref, test)])), ctx)


@function(
    "AVERAGEIFS", "raw", "raw", "value", min_args=3, variadic=True, errors=True, cycle=2, category=C
)
def AVERAGEIFS(ctx: Context, cells: Any, *pairs: Any) -> float | ErrorValue:
    """Mean of ``cells`` where every range meets its criterion."""
    values = _ifs_values(ctx, cells, pairs)
    return values if isinstance(values, ErrorValue) else _mean(values, ctx)


def _extreme(ctx: Context, cells: Any, pairs: tuple[Any, ...], pick: Any) -> float | ErrorValue:
    values = _ifs_values(ctx, cells, pairs)
    if isinstance(values, ErrorValue):
        return values
    for v in values:
        if isinstance(v, ErrorValue):
            return v
    return pick(values) if values else 0.0


@function(
    "_xlfn.MAXIFS",
    "raw",
    "raw",
    "value",
    min_args=3,
    variadic=True,
    errors=True,
    cycle=2,
    category=C,
)
def MAXIFS(ctx: Context, cells: Any, *pairs: Any) -> float | ErrorValue:
    """Largest of ``cells`` where every range meets its criterion; 0 when there is none."""
    return _extreme(ctx, cells, pairs, max)


@function(
    "_xlfn.MINIFS",
    "raw",
    "raw",
    "value",
    min_args=3,
    variadic=True,
    errors=True,
    cycle=2,
    category=C,
)
def MINIFS(ctx: Context, cells: Any, *pairs: Any) -> float | ErrorValue:
    """Smallest of ``cells`` where every range meets its criterion; 0 when there is none."""
    return _extreme(ctx, cells, pairs, min)


@function("COUNTIF", "raw", "value", errors=True, category=C)
def COUNTIF(ctx: Context, ref: Any, crit: Any) -> float | ErrorValue:
    """Cells meeting the criterion."""
    if not isinstance(ref, Ref):
        return VALUE
    test = criterion(scalar(crit, ctx), ctx.quirks)
    filled = 0
    count = 0
    for _, _, v in ref.cells():
        filled += 1
        count += test(v)
    if test(BLANK):
        count += ref.rows * ref.columns - filled
    return float(count)


@function("COUNTIFS", "raw", "value", min_args=2, variadic=True, errors=True, cycle=2, category=C)
def COUNTIFS(ctx: Context, *pairs: Any) -> float | ErrorValue:
    """Positions where every range meets its criterion.

    Only non-blank cells are visited (§7.7). When a criterion rejects a blank, only that range's
    non-blank cells can count; otherwise every position counts but those where a non-blank cell
    fails its criterion.
    """
    conditions = _conditions(ctx, pairs)
    if isinstance(conditions, ErrorValue):
        return conditions
    for lead, lead_test in conditions:
        if not lead_test(BLANK):
            count = 0
            for r, c, v in lead.cells():
                count += lead_test(v) and all(test(ref.value(r, c)) for ref, test in conditions)
            return float(count)
    rows, cols = pairs[0].rows, pairs[0].columns
    failed = {r * cols + c for ref, test in conditions for r, c, v in ref.cells() if not test(v)}
    return float(rows * cols - len(failed))


@function("SUMPRODUCT", "array", min_args=1, variadic=True, category=C)
def SUMPRODUCT(ctx: Context, *arrays: Any) -> float | ErrorValue:
    """Sum of products of same-shaped ranges; non-numbers count as 0."""
    grids = []
    for a in arrays:
        if isinstance(a, Ref):
            check_area(a.rows, a.columns, getattr(ctx, "max_array_cells", None))
            grids.append([[a.value(r, c) for c in range(a.columns)] for r in range(a.rows)])
        elif isinstance(a, tuple):
            grids.append([list(row) for row in a])
        else:
            v = scalar(a, ctx)
            if isinstance(v, ErrorValue):
                return v
            grids.append([[v]])
    shape = (len(grids[0]), len(grids[0][0]))
    if any((len(g), len(g[0])) != shape for g in grids):
        return VALUE
    total = 0.0
    for r in range(shape[0]):
        for c in range(shape[1]):
            p = 1.0
            for g in grids:
                v = g[r][c]
                if isinstance(v, ErrorValue):
                    return v
                p *= v if isinstance(v, float) and not isinstance(v, bool) else 0.0
            total += p
    return flush(total) if math.isfinite(total) else NUM


_SUBTOTAL = {1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT",
             7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP"}  # fmt: skip


@function(
    "SUBTOTAL",
    "number",
    "raw",
    min_args=2,
    variadic=True,
    category=C,
)
def SUBTOTAL(ctx: Context, kind: float, *refs: Any) -> float | ErrorValue:
    """An aggregate over references, leaving out cells that hold SUBTOTAL themselves.

    Codes 101 to 111 leave out every hidden row. Codes 1 to 11 do so too when a filter is hiding
    rows on any sheet the call names, each sheet by its own hidden rows, even one with no filter
    (capture ``hidden_rows``). Hidden columns stay in.
    """
    code = int(kind) % 100
    name = _SUBTOTAL.get(code)
    if name is None or int(kind) not in (code, code + 100):
        return VALUE
    if not all(isinstance(r, Ref) for r in refs):
        return VALUE
    skip = getattr(ctx, "is_subtotal", lambda sheet, row, col: False)
    values: list[Any] = []
    filtering = getattr(ctx, "filtering", lambda sheet: False)
    row_hidden = getattr(ctx, "row_hidden", lambda sheet, row: False)
    leave_hidden = int(kind) > 100 or any(filtering(ref.sheet) for ref in refs)
    for ref in refs:
        for r, c, v in ref.cells():
            row = ref.r1 + r
            if skip(ref.sheet, row, ref.c1 + c) or (leave_hidden and row_hidden(ref.sheet, row)):
                continue
            values.append(v)
    nums = [v for v in values if isinstance(v, float) and not isinstance(v, bool)]
    errors = [v for v in values if isinstance(v, ErrorValue)]
    if errors:
        return errors[0]
    if name == "SUM":
        return _total(nums, ctx)
    if name == "COUNT":
        return float(len(nums))
    if name == "COUNTA":
        return float(sum(1 for v in values if v is not BLANK))
    if name == "AVERAGE":
        return _total(nums, ctx) / len(nums) if nums else DIV0  # type: ignore[operator]
    if name in ("MAX", "MIN"):
        return (max if name == "MAX" else min)(nums) if nums else 0.0
    if name == "PRODUCT":
        return math.prod(nums) if nums else 0.0
    need = 2 if name in ("STDEV", "VAR") else 1
    if len(nums) < need:
        return DIV0
    fn = {"STDEV": statistics.stdev, "STDEVP": statistics.pstdev,
          "VAR": statistics.variance, "VARP": statistics.pvariance}[name]  # fmt: skip
    return float(fn(nums))


@function("RAND", volatile=True, category=C)
def RAND(ctx: Context) -> float:
    """A random number from 0 up to 1, new at every calculation."""
    return random.random()


@function("RANDBETWEEN", "number", "number", volatile=True, category=C)
def RANDBETWEEN(ctx: Context, low: float, high: float) -> float | ErrorValue:
    """A random whole number between two bounds, both rounded up first (captured)."""
    lo, hi = math.ceil(low), math.ceil(high)
    if lo > hi:
        return NUM
    return float(random.randint(lo, hi))


# Trigonometry, logarithms and sign ---------------------------------------------------------------


_LIBM = "Excel's own sine, cosine, logarithm and exponential differ in the last bits"


def _finite(x: float) -> float | ErrorValue:
    return x if math.isfinite(x) else NUM


def _circular_argument(x: float) -> ErrorValue | None:
    """Excel refuses an angle too large to reduce (capture ``math_trig``)."""
    return NUM if abs(x) >= _LARGEST_ANGLE else None


_LARGEST_ANGLE = 2.0**27


@function("SIN", "number", category=C, ulps=4, ulps_reason=_LIBM)
def SIN(ctx: Context, x: float) -> float | ErrorValue:
    """Sine of an angle in radians."""
    return _circular_argument(x) or math.sin(x)


@function("COS", "number", category=C, ulps=4, ulps_reason=_LIBM)
def COS(ctx: Context, x: float) -> float | ErrorValue:
    """Cosine of an angle in radians."""
    return _circular_argument(x) or math.cos(x)


@function("TAN", "number", category=C, ulps=4, ulps_reason=_LIBM)
def TAN(ctx: Context, x: float) -> float | ErrorValue:
    """Tangent of an angle in radians."""
    return _circular_argument(x) or math.tan(x)


@function("_XLFN.COT", "number", category=C, ulps=4, ulps_reason=_LIBM)
def COT(ctx: Context, x: float) -> float | ErrorValue:
    """Cotangent; ``#DIV/0!`` at 0."""
    if x == 0:
        return DIV0
    return _circular_argument(x) or 1.0 / math.tan(x)


@function("_XLFN.SEC", "number", category=C, ulps=4, ulps_reason=_LIBM)
def SEC(ctx: Context, x: float) -> float | ErrorValue:
    """Secant."""
    return _circular_argument(x) or 1.0 / math.cos(x)


@function("_XLFN.CSC", "number", category=C, ulps=4, ulps_reason=_LIBM)
def CSC(ctx: Context, x: float) -> float | ErrorValue:
    """Cosecant; ``#DIV/0!`` at 0."""
    if x == 0:
        return DIV0
    return _circular_argument(x) or 1.0 / math.sin(x)


@function("ASIN", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ASIN(ctx: Context, x: float) -> float | ErrorValue:
    """Arcsine; ``#NUM!`` outside -1 to 1."""
    return NUM if abs(x) > 1 else math.asin(x)


@function("ACOS", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ACOS(ctx: Context, x: float) -> float | ErrorValue:
    """Arccosine; ``#NUM!`` outside -1 to 1."""
    return NUM if abs(x) > 1 else math.acos(x)


@function("ATAN", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ATAN(ctx: Context, x: float) -> float:
    """Arctangent."""
    return math.atan(x)


@function("ATAN2", "number", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ATAN2(ctx: Context, x: float, y: float) -> float | ErrorValue:
    """The angle of the point (x, y): Excel's order, ``#DIV/0!`` at the origin."""
    return DIV0 if x == 0 and y == 0 else math.atan2(y, x)


@function("_XLFN.ACOT", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ACOT(ctx: Context, x: float) -> float:
    """Arccotangent, between 0 and pi."""
    if x == 0:
        return math.pi / 2
    return math.atan(1.0 / x) + (math.pi if x < 0 else 0.0)


def _hyper(x: float, sign: float, over: float) -> float | ErrorValue:
    """(e^x + sign * e^-x) / over, as Excel takes them from 1 up (capture ``math_trig``)."""
    if sign < 0 and abs(x) < 1:
        return math.sinh(x)
    try:
        return (math.exp(x) + sign * math.exp(-x)) / over
    except OverflowError:
        return NUM


@function("SINH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def SINH(ctx: Context, x: float) -> float | ErrorValue:
    """Hyperbolic sine."""
    return _hyper(x, -1.0, 2.0)


@function("COSH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def COSH(ctx: Context, x: float) -> float | ErrorValue:
    """Hyperbolic cosine."""
    return _hyper(x, 1.0, 2.0)


@function("TANH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def TANH(ctx: Context, x: float) -> float:
    """Hyperbolic tangent."""
    if abs(x) < 1:
        return math.tanh(x)
    if abs(x) > 300:
        return math.copysign(1.0, x)
    a, b = math.exp(x), math.exp(-x)
    return (a - b) / (a + b)


@function("_XLFN.COTH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def COTH(ctx: Context, x: float) -> float | ErrorValue:
    """Hyperbolic cotangent; ``#DIV/0!`` at 0."""
    if x == 0:
        return DIV0
    if abs(x) < 1:
        return 1.0 / math.tanh(x)
    if abs(x) > 300:
        return math.copysign(1.0, x)
    a, b = math.exp(x), math.exp(-x)
    return (a + b) / (a - b)


@function("_XLFN.SECH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def SECH(ctx: Context, x: float) -> float:
    """Hyperbolic secant."""
    value = _hyper(x, 1.0, 2.0)
    return 0.0 if isinstance(value, ErrorValue) else 1.0 / value


@function("_XLFN.CSCH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def CSCH(ctx: Context, x: float) -> float | ErrorValue:
    """Hyperbolic cosecant; ``#DIV/0!`` at 0."""
    if x == 0:
        return DIV0
    value = _hyper(x, -1.0, 2.0)
    return 0.0 if isinstance(value, ErrorValue) else 1.0 / value


@function("ASINH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ASINH(ctx: Context, x: float) -> float | ErrorValue:
    """Inverse hyperbolic sine."""
    square = x * x
    if not math.isfinite(square):
        return NUM
    return math.copysign(math.log(abs(x) + math.sqrt(square + 1)), x)


@function("ACOSH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ACOSH(ctx: Context, x: float) -> float | ErrorValue:
    """Inverse hyperbolic cosine; ``#NUM!`` below 1."""
    if x < 1 or not math.isfinite(x * x):
        return NUM
    return math.log(x + math.sqrt(x * x - 1))


@function("ATANH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ATANH(ctx: Context, x: float) -> float | ErrorValue:
    """Inverse hyperbolic tangent; ``#NUM!`` from -1 and 1 outward."""
    if abs(x) >= 1:
        return NUM
    return math.atanh(x) if abs(x) < 1e-5 else 0.5 * math.log((1 + x) / (1 - x))


@function("_XLFN.ACOTH", "number", category=C, ulps=4, ulps_reason=_LIBM)
def ACOTH(ctx: Context, x: float) -> float | ErrorValue:
    """Inverse hyperbolic cotangent; ``#NUM!`` between -1 and 1."""
    return NUM if abs(x) <= 1 else math.atanh(1.0 / x)


@function("DEGREES", "number", category=C)
def DEGREES(ctx: Context, x: float) -> float | ErrorValue:
    """Radians as degrees."""
    return _finite(x * (180.0 / math.pi))


@function("RADIANS", "number", category=C)
def RADIANS(ctx: Context, x: float) -> float:
    """Degrees as radians."""
    return x * (math.pi / 180.0)


@function("SIGN", "number", category=C)
def SIGN(ctx: Context, x: float) -> float:
    """1, -1 or 0."""
    return 1.0 if x > 0 else -1.0 if x < 0 else 0.0


MISSING_BASE = float("nan")  # compared by identity, never a cell's value


@function("LOG", "number", "number", min_args=1, category=C, ulps=4, ulps_reason=_LIBM)
def LOG(ctx: Context, x: float, base: float = MISSING_BASE) -> float | ErrorValue:
    """Logarithm to a base, 10 when left out; a ratio of logarithms when it is given."""
    if base is MISSING_BASE:
        return NUM if x <= 0 else math.log10(x)
    if x <= 0 or base <= 0:
        return NUM
    if base == 1:
        return DIV0
    return math.log(x) / math.log(base)


@function("LOG10", "number", category=C, ulps=4, ulps_reason=_LIBM)
def LOG10(ctx: Context, x: float) -> float | ErrorValue:
    """Base-10 logarithm."""
    return NUM if x <= 0 else math.log10(x)


@function("SQRTPI", "value", category=C)
def SQRTPI(ctx: Context, a: Any) -> float | ErrorValue:
    """Square root of a number times pi."""
    x = _strict(a, ctx)
    if isinstance(x, ErrorValue):
        return x
    return NUM if x < 0 else _finite(math.sqrt(x * math.pi))


@function("TRUNC", "number", "number", min_args=1, category=C)
def TRUNC(ctx: Context, x: float, digits: float = 0.0) -> float | ErrorValue:
    """Cut to a number of digits; at 15 or more, the value at 15 significant digits."""
    if digits >= 15:
        return float(f"{x:.15g}")
    if digits == 0 and abs(x) >= 2.0**53:
        return x
    return ROUNDDOWN(ctx, x, digits)  # type: ignore[no-any-return]


def _strict(v: Any, ctx: Context) -> float | ErrorValue:
    """A number as the old Analysis ToolPak functions take one: a logical is ``#VALUE!``."""
    return VALUE if isinstance(v, bool) else to_number(v, ctx)


# Multiples ---------------------------------------------------------------------------------------


def _to_multiple(x: float, m: float, up: bool) -> float:
    """The multiple of ``m`` above (``up``) or below ``x``, with a ratio's noise taken out."""
    ratio = x / m
    if not math.isfinite(ratio):
        return math.inf
    near = round(ratio)
    if abs(ratio - near) <= 1e-13 * abs(ratio):
        return flush(near * m)
    return flush((math.ceil(ratio) if up else math.floor(ratio)) * m)


@function("CEILING", "number", "number", category=C)
def CEILING(ctx: Context, x: float, m: float) -> float | ErrorValue:
    """Up to a multiple; a negative number with a negative multiple goes away from zero."""
    if x > 0 and m < 0:
        return NUM
    if m == 0 or x == 0:
        return 0.0
    return _to_multiple(x, m, True)


@function("FLOOR", "number", "number", category=C)
def FLOOR(ctx: Context, x: float, m: float) -> float | ErrorValue:
    """Down to a multiple; a negative number with a negative multiple goes toward zero."""
    if x > 0 and m < 0:
        return NUM
    if x == 0:
        return 0.0
    if m == 0:
        return DIV0
    return _to_multiple(x, m, False)


@function("_XLFN.CEILING.PRECISE", "number", "number", min_args=1, category=C)
def CEILING_PRECISE(ctx: Context, x: float, m: float = 1.0) -> float:
    """Up, whatever the signs."""
    return 0.0 if m == 0 or x == 0 else _to_multiple(x, abs(m), True)


@function("ISO.CEILING", "number", "number", min_args=1, category=C)
def ISO_CEILING(ctx: Context, x: float, m: float = 1.0) -> float:
    """As ``CEILING.PRECISE``."""
    return CEILING_PRECISE(ctx, x, m)  # type: ignore[no-any-return]


@function("_XLFN.FLOOR.PRECISE", "number", "number", min_args=1, category=C)
def FLOOR_PRECISE(ctx: Context, x: float, m: float = 1.0) -> float:
    """Down, whatever the signs."""
    return 0.0 if m == 0 or x == 0 else _to_multiple(x, abs(m), False)


@function("_XLFN.CEILING.MATH", "number", "number", "number", min_args=1, category=C)
def CEILING_MATH(ctx: Context, x: float, m: float = 1.0, mode: float = 0.0) -> float:
    """Up; a mode other than 0 takes a negative number away from zero."""
    if m == 0 or x == 0:
        return 0.0
    return _to_multiple(x, abs(m), not (x < 0 and mode != 0))


@function("_XLFN.FLOOR.MATH", "number", "number", "number", min_args=1, category=C)
def FLOOR_MATH(ctx: Context, x: float, m: float = 1.0, mode: float = 0.0) -> float:
    """Down; a mode other than 0 takes a negative number toward zero."""
    if m == 0 or x == 0:
        return 0.0
    return _to_multiple(x, abs(m), x < 0 and mode != 0)


@function("MROUND", "value", "value", errors=True, category=C)
def MROUND(ctx: Context, a: Any, b: Any) -> float | ErrorValue:
    """The nearest multiple; ``#NUM!`` when the signs differ."""
    x, m = _strict(a, ctx), _strict(b, ctx)
    if isinstance(x, ErrorValue):
        return x
    if isinstance(m, ErrorValue):
        return m
    if m == 0 or x == 0:
        return 0.0
    if (x < 0) != (m < 0):
        return NUM
    ratio = x / m + 0.5
    return flush(math.floor(ratio) * m) if math.isfinite(ratio) else NUM


@function("EVEN", "number", category=C)
def EVEN(ctx: Context, x: float) -> float | ErrorValue:
    """Away from zero to an even integer."""
    if x == 0:
        return 0.0
    return math.copysign(math.ceil(abs(x) / 2) * 2.0, x)


@function("ODD", "number", category=C)
def ODD(ctx: Context, x: float) -> float | ErrorValue:
    """Away from zero to an odd integer."""
    return math.copysign(math.ceil((abs(x) + 1) / 2) * 2.0 - 1.0, x if x else 1.0)


@function("QUOTIENT", "value", "value", errors=True, category=C)
def QUOTIENT(ctx: Context, a: Any, b: Any) -> float | ErrorValue:
    """The integer part of a division."""
    x, y = _strict(a, ctx), _strict(b, ctx)
    if isinstance(x, ErrorValue):
        return x
    if isinstance(y, ErrorValue):
        return y
    if y == 0:
        return DIV0
    ratio = x / y
    return float(math.trunc(ratio)) if math.isfinite(ratio) else NUM


# Integers ----------------------------------------------------------------------------------------


_PRODUCT = "the order in which Excel multiplies, which a few tried orders did not reproduce"
_PAST_DOUBLES = 2**1024
_LARGEST_N = 2147483647.0  # COMBIN, PERMUT: 2147483646 is computed, this is #NUM! (capture)


def _big(value: int) -> float | ErrorValue:
    return NUM if value >= _PAST_DOUBLES else float(value)


def _product_of_falling(n: int, k: int, dividing: bool) -> float | ErrorValue:
    """n(n-1)...(n-k+1), over k! when ``dividing``, stopping past a double's range.

    Stepwise, with no big factorial: a million steps cannot be reached before the range ends
    unless n is small, and then k is too, so the cost is bounded whatever a formula asks for.
    """
    if dividing:
        k = min(k, n - k)
    out = 1
    for i in range(k):
        out = out * (n - i)
        if dividing:
            out //= i + 1
        if out >= _PAST_DOUBLES:
            return NUM
    return float(out)


@function("FACT", "number", category=C, ulps=8, ulps_reason=_PRODUCT)
def FACT(ctx: Context, x: float) -> float | ErrorValue:
    """Factorial of the integer part, a product of doubles as Excel takes it."""
    if x < 0 or x >= 171:
        return NUM
    out = 1.0
    for i in range(2, math.trunc(x) + 1):
        out *= i
    return out


@function("FACTDOUBLE", "value", category=C, ulps=8, ulps_reason=_PRODUCT)
def FACTDOUBLE(ctx: Context, a: Any) -> float | ErrorValue:
    """Double factorial of the integer part, a product of doubles."""
    x = _strict(a, ctx)
    if isinstance(x, ErrorValue):
        return x
    if x < -1 or x >= 301:
        return NUM
    n = math.trunc(x)
    out = 1.0
    while n > 1:
        out *= n
        n -= 2
    return out


@function("COMBIN", "number", "number", category=C, ulps=4, ulps_reason=_PRODUCT)
def COMBIN(ctx: Context, n: float, k: float) -> float | ErrorValue:
    """Combinations."""
    a, b = math.trunc(n), math.trunc(k)
    if a < 0 or b < 0 or b > a or n >= _LARGEST_N:
        return NUM
    return _product_of_falling(a, b, True)


@function("_XLFN.COMBINA", "number", "number", category=C, ulps=4, ulps_reason=_PRODUCT)
def COMBINA(ctx: Context, n: float, k: float) -> float | ErrorValue:
    """Combinations with repetition."""
    a, b = math.trunc(n), math.trunc(k)
    if a < 0 or b < 0 or (a == 0 and b > 0) or n >= _LARGEST_N or k >= _LARGEST_N:
        return NUM
    return 1.0 if a == 0 else _product_of_falling(a + b - 1, b, True)


@function("PERMUT", "number", "number", category=C, ulps=4, ulps_reason=_PRODUCT)
def PERMUT(ctx: Context, n: float, k: float) -> float | ErrorValue:
    """Permutations."""
    a, b = math.trunc(n), math.trunc(k)
    if a < 0 or b < 0 or b > a or n >= _LARGEST_N:
        return NUM
    return _product_of_falling(a, b, False)


@function("_XLFN.PERMUTATIONA", "number", "number", category=C)
def PERMUTATIONA(ctx: Context, n: float, k: float) -> float | ErrorValue:
    """Permutations with repetition."""
    a, b = math.trunc(n), math.trunc(k)
    if a < 0 or b < 0 or a >= 2**31 or b >= 2**31:
        return NUM
    if a > 1 and b * math.log2(a) >= 1024:
        return NUM
    return _big(a**b) if a > 1 else float(a**b)


def _whole_numbers(args: tuple[Any, ...], ctx: Context) -> list[int] | ErrorValue:
    """The integer parts of the arguments, text in a cell converted, a logical ``#VALUE!``."""
    if args[0] is MISSING:  # an empty first argument, and later ones are passed over
        return ErrorValue.NA
    out = []
    negative = False
    for arg in args:
        if arg is MISSING:
            continue
        for v in arg.values() if isinstance(arg, Ref) else (arg,):
            x = _strict(v, ctx)
            if isinstance(x, ErrorValue):
                return x
            negative = negative or x < 0 or x > 2.0**53
            out.append(0 if negative else math.trunc(x))
    if negative:
        return NUM
    return out or VALUE  # nothing but blank cells


@function("GCD", "raw", min_args=1, variadic=True, category=C)
def GCD(ctx: Context, *args: Any) -> float | ErrorValue:
    """Greatest common divisor of the integer parts."""
    values = _whole_numbers(args, ctx)
    return values if isinstance(values, ErrorValue) else _big(math.gcd(*values))


@function("LCM", "raw", min_args=1, variadic=True, category=C)
def LCM(ctx: Context, *args: Any) -> float | ErrorValue:
    """Least common multiple of the integer parts."""
    values = _whole_numbers(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    out = 1
    for v in values:
        out = math.lcm(out, v)
        if out > 2**53:
            return NUM
    return float(out)


@function("MULTINOMIAL", "raw", min_args=1, variadic=True, category=C)
def MULTINOMIAL(ctx: Context, *args: Any) -> float | ErrorValue:
    """The factorial of a sum over the product of the factorials."""
    values = _whole_numbers(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    if sum(values) >= 171:
        return NUM
    out = math.factorial(sum(values))
    for v in values:
        out //= math.factorial(v)
    return float(out)


# Number forms ------------------------------------------------------------------------------------

_ROMAN_FORMS: tuple[tuple[tuple[int, str], ...], ...] = tuple(
    tuple(sorted(pairs, reverse=True))
    for pairs in (
        [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"),
         (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
        [(1000, "M"), (950, "LM"), (900, "CM"), (500, "D"), (450, "LD"), (400, "CD"), (100, "C"),
         (95, "VC"), (90, "XC"), (50, "L"), (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"),
         (4, "IV"), (1, "I")],
        [(1000, "M"), (990, "XM"), (950, "LM"), (900, "CM"), (500, "D"), (490, "XD"), (450, "LD"),
         (400, "CD"), (100, "C"), (99, "IC"), (95, "VC"), (90, "XC"), (50, "L"), (49, "IL"),
         (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")],
        [(1000, "M"), (995, "VM"), (990, "XM"), (950, "LM"), (900, "CM"), (500, "D"), (495, "VD"),
         (490, "XD"), (450, "LD"), (400, "CD"), (100, "C"), (99, "IC"), (95, "VC"), (90, "XC"),
         (50, "L"), (49, "IL"), (45, "VL"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"),
         (1, "I")],
        [(1000, "M"), (999, "IM"), (995, "VM"), (990, "XM"), (950, "LM"), (900, "CM"), (500, "D"),
         (499, "ID"), (495, "VD"), (490, "XD"), (450, "LD"), (400, "CD"), (100, "C"), (99, "IC"),
         (95, "VC"), (90, "XC"), (50, "L"), (49, "IL"), (45, "VL"), (40, "XL"), (10, "X"),
         (9, "IX"),
         (5, "V"), (4, "IV"), (1, "I")],
    )
)  # fmt: skip


@function("ROMAN", "number", "value", min_args=1, category=C)
def ROMAN(ctx: Context, x: float, form: Any = 0.0) -> str | ErrorValue:
    """A Roman numeral; forms 0 to 4 are ever more concise, TRUE as 0 and FALSE as 4."""
    if form is True:
        form = 0.0
    elif form is False:
        form = 4.0
    elif isinstance(form, ErrorValue):
        return form
    elif form is BLANK or form is None:
        form = 0.0
    if not isinstance(form, float):
        return VALUE
    n = math.trunc(x)
    if n < 0 or n > 3999 or not 0 <= form < 5:
        return VALUE
    out = []
    for value, text in _ROMAN_FORMS[int(form)]:
        while n >= value:
            out.append(text)
            n -= value
    return "".join(out)


_ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}


@function("_XLFN.ARABIC", "text", category=C)
def ARABIC(ctx: Context, text: str) -> float | ErrorValue:
    """A Roman numeral's value."""
    s = text.strip().upper()
    sign = 1.0
    if s.startswith("-"):
        sign, s = -1.0, s[1:]
        if not s:
            return NUM
    if not s:
        return 0.0
    if any(c not in _ROMAN_VALUES for c in s):
        return VALUE
    total = 0
    for i, c in enumerate(s):
        value = _ROMAN_VALUES[c]
        total += -value if i + 1 < len(s) and _ROMAN_VALUES[s[i + 1]] > value else value
    return sign * total


_DIGITS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


@function("_XLFN.BASE", "number", "number", "number", min_args=2, category=C)
def BASE(ctx: Context, x: float, radix: float, length: float = 0.0) -> str | ErrorValue:
    """A whole number written in a base from 2 to 36."""
    r, n = math.trunc(radix), math.trunc(length)
    if x < 0 or x >= 2.0**53 or not 2 <= r <= 36 or not 0 <= n <= 255:
        return NUM
    value, out = math.trunc(x), []
    while value:
        value, digit = divmod(value, r)
        out.append(_DIGITS[digit])
    return "".join(reversed(out)).rjust(n, "0")


@function("_XLFN.DECIMAL", "text", "number", category=C)
def DECIMAL(ctx: Context, text: str, radix: float) -> float | ErrorValue:
    """Text in a base from 2 to 36 as a number."""
    r = math.trunc(radix)
    if not 2 <= r <= 36 or len(text) > 255:
        return NUM
    total = 0
    for c in text.strip().upper():
        digit = _DIGITS.find(c)
        if digit < 0 or digit >= r:
            return NUM
        total = total * r + digit
    return _big(total)


# Sums of series and pairs ------------------------------------------------------------------------


@function("SERIESSUM", "number", "number", "number", "raw", category=C, ulps=4, ulps_reason=_LIBM)
def SERIESSUM(ctx: Context, x: float, n: float, m: float, coefficients: Any) -> float | ErrorValue:
    """A power series: the coefficients times x to n, n+m, n+2m and so on."""
    total = 0.0
    for k, v in enumerate(numbers((coefficients,), ctx)):
        if isinstance(v, ErrorValue):
            return v
        if x == 0 and n + k * m <= 0:
            return NUM
        try:
            power = x ** (n + k * m)
        except (ZeroDivisionError, OverflowError):
            return NUM
        if isinstance(power, complex):
            return NUM
        total += v * power
    return _finite(total)


def _cells_of(value: Any) -> list[list[Any]]:
    grid = to_array(value) if isinstance(value, Ref) else value if is_array(value) else ((value,),)
    rows, cols = shape(grid)
    return [[element(grid, rows, cols, i, j) for j in range(cols)] for i in range(rows)]


def _paired(a: Any, b: Any) -> list[tuple[float, float]] | ErrorValue:
    left, right = _cells_of(a), _cells_of(b)
    if len(left) != len(right) or len(left[0]) != len(right[0]):
        return ErrorValue.NA
    pairs = []
    for row_a, row_b in zip(left, right, strict=True):
        for u, v in zip(row_a, row_b, strict=True):
            if isinstance(u, ErrorValue):
                return u
            if isinstance(v, ErrorValue):
                return v
            if type(u) is float and type(v) is float:
                pairs.append((u, v))
    return pairs or DIV0


@function("SUMX2MY2", "array", "array", category=C)
def SUMX2MY2(ctx: Context, a: Any, b: Any) -> float | ErrorValue:
    """Sum of the differences of the squares of pairs."""
    pairs = _paired(a, b)
    return pairs if isinstance(pairs, ErrorValue) else _finite(sum(x * x - y * y for x, y in pairs))


@function("SUMX2PY2", "array", "array", category=C)
def SUMX2PY2(ctx: Context, a: Any, b: Any) -> float | ErrorValue:
    """Sum of the sums of the squares of pairs."""
    pairs = _paired(a, b)
    return pairs if isinstance(pairs, ErrorValue) else _finite(sum(x * x + y * y for x, y in pairs))


@function("SUMXMY2", "array", "array", category=C)
def SUMXMY2(ctx: Context, a: Any, b: Any) -> float | ErrorValue:
    """Sum of the squares of the differences of pairs."""
    pairs = _paired(a, b)
    return pairs if isinstance(pairs, ErrorValue) else _finite(sum((x - y) ** 2 for x, y in pairs))
