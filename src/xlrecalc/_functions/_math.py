"""Math and aggregation: SUM, ROUND, INT, MOD, SUMIF, SUMPRODUCT, SUBTOTAL and friends."""

from __future__ import annotations

import math
import random
import statistics
from decimal import ROUND_DOWN, ROUND_HALF_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal
from fractions import Fraction
from typing import Any

from xlrecalc._arith import flush, snap
from xlrecalc._arrays import check_area
from xlrecalc._coerce import Context, numbers, scalar
from xlrecalc._compare import criterion
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._numtext import DBL_MIN
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
    for v in numbers(args, ctx):
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


@function("SUMIF", "raw", "raw", "raw", min_args=2, category=C)
def SUMIF(ctx: Context, ref: Any, crit: Any, sums: Any = MISSING) -> float | ErrorValue:
    """Sum of the cells of ``sums`` (or ``ref``) whose cell in ``ref`` meets the criterion.

    ``sums`` takes the shape of ``ref`` from its top-left cell, as Excel does.
    """
    if not isinstance(ref, Ref):
        return VALUE
    target = ref
    if isinstance(sums, Ref):
        target = Ref(sums.source, sums.sheet, sums.r1, sums.c1,
                     sums.r1 + ref.rows - 1, sums.c1 + ref.columns - 1)  # fmt: skip
    elif sums is not MISSING:
        return VALUE
    test = criterion(scalar(crit, ctx), ctx.quirks)
    if test(BLANK):  # blank cells match: every position counts
        check_area(ref.rows, ref.columns, getattr(ctx, "max_array_cells", None))
        positions: Any = ((r, c) for r in range(ref.rows) for c in range(ref.columns))
        picked = (target.value(r, c) for r, c in positions if test(ref.value(r, c)))
    else:  # only non-blank cells can match: visit those
        picked = (target.value(r, c) for r, c, v in ref.cells() if test(v))
    return _total(
        (v for v in picked if isinstance(v, float | ErrorValue) and not isinstance(v, bool)), ctx
    )


def _multi(ctx: Context, pairs: tuple[Any, ...]) -> list[tuple[int, int]] | ErrorValue:
    """Positions meeting every (range, criterion) pair."""
    if not pairs or len(pairs) % 2:
        return VALUE
    first = pairs[0]
    if not isinstance(first, Ref):
        return VALUE
    rows, cols = first.rows, first.columns
    keep = [(r, c) for r in range(rows) for c in range(cols)]
    for k in range(0, len(pairs), 2):
        ref = pairs[k]
        if not isinstance(ref, Ref) or ref.rows != rows or ref.columns != cols:
            return VALUE
        test = criterion(scalar(pairs[k + 1], ctx), ctx.quirks)
        keep = [(r, c) for r, c in keep if test(ref.value(r, c))]
    return keep


@function("SUMIFS", "raw", "raw", "raw", min_args=3, variadic=True, category=C)
def SUMIFS(ctx: Context, sums: Any, *pairs: Any) -> float | ErrorValue:
    """Sum of ``sums`` where every range meets its criterion."""
    if not isinstance(sums, Ref):
        return VALUE
    keep = _multi(ctx, pairs)
    if isinstance(keep, ErrorValue):
        return keep
    if sums.rows != pairs[0].rows or sums.columns != pairs[0].columns:
        return VALUE
    values = (sums.value(r, c) for r, c in keep)
    return _total(
        (v for v in values if isinstance(v, float | ErrorValue) and not isinstance(v, bool)), ctx
    )


@function("COUNTIF", "raw", "raw", category=C)
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


@function("COUNTIFS", "raw", "raw", min_args=2, variadic=True, category=C)
def COUNTIFS(ctx: Context, *pairs: Any) -> float | ErrorValue:
    """Positions where every range meets its criterion."""
    keep = _multi(ctx, pairs)
    return keep if isinstance(keep, ErrorValue) else float(len(keep))


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
    status="partial",
    note="codes 101 to 111 act as 1 to 11: row visibility is not read",
    category=C,
)
def SUBTOTAL(ctx: Context, kind: float, *refs: Any) -> float | ErrorValue:
    """An aggregate over references, leaving out cells that hold SUBTOTAL themselves.

    Codes 101 to 111 also leave out hidden rows; row visibility is not read yet, so they act
    as 1 to 11.
    """
    code = int(kind) % 100
    name = _SUBTOTAL.get(code)
    if name is None or int(kind) not in (code, code + 100):
        return VALUE
    if not all(isinstance(r, Ref) for r in refs):
        return VALUE
    skip = getattr(ctx, "is_subtotal", lambda sheet, row, col: False)
    values: list[Any] = []
    for ref in refs:
        for r, c, v in ref.cells():
            if not skip(ref.sheet, ref.r1 + r, ref.c1 + c):
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
