"""Statistical functions: AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK."""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._arrays import to_array
from xlrecalc._coerce import Context, numbers, to_number
from xlrecalc._formula import MISSING
from xlrecalc._functions._math import _total
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Ref

DIV0, NA, NUM, VALUE = ErrorValue.DIV0, ErrorValue.NA, ErrorValue.NUM, ErrorValue.VALUE
S = "statistical"
_VARIANCE = (
    "Excel's variance is not reproduced in the last bits (two-pass with plain or compensated "
    "means, the textbook and corrected formulas tried); within 4 units in the last place"
)
_REGRESSION = (
    "Excel's sums of products and its regression (likely by QR decomposition) differ in the "
    "last bits; within 4 units in the last place except listed rows near cancellation"
)
_NORMAL = (
    "Excel's normal distribution is not math.erfc or Cody's algorithm in the last bits; "
    "within 4 units in the last place except listed far-tail rows"
)


def _numbers(args: tuple[Any, ...], ctx: Context) -> list[float] | ErrorValue:
    out = []
    for v in numbers(args, ctx):
        if isinstance(v, ErrorValue):
            return v
        out.append(v)
    return out


@function("AVERAGE", "raw", min_args=1, variadic=True, category=S)
def AVERAGE(ctx: Context, *args: Any) -> float | ErrorValue:
    """Mean of the numbers; ``#DIV/0!`` when there are none."""
    values = _numbers(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    if not values:
        return DIV0
    total = _total(values, ctx)
    return total if isinstance(total, ErrorValue) else total / len(values)


@function("MIN", "raw", min_args=1, variadic=True, category=S)
def MIN(ctx: Context, *args: Any) -> float | ErrorValue:
    """Smallest number; 0 when there are none."""
    values = _numbers(args, ctx)
    return values if isinstance(values, ErrorValue) else (min(values) if values else 0.0)


@function("MAX", "raw", min_args=1, variadic=True, category=S)
def MAX(ctx: Context, *args: Any) -> float | ErrorValue:
    """Largest number; 0 when there are none."""
    values = _numbers(args, ctx)
    return values if isinstance(values, ErrorValue) else (max(values) if values else 0.0)


@function("COUNT", "raw", min_args=1, variadic=True, category=S)
def COUNT(ctx: Context, *args: Any) -> float:
    """Numbers among the values; errors and other values are not counted."""
    return float(
        sum(1 for v in numbers(args, ctx, direct_text="skip") if not isinstance(v, ErrorValue))
    )


@function("COUNTA", "raw", min_args=1, variadic=True, category=S)
def COUNTA(ctx: Context, *args: Any) -> float:
    """Values that are not blank (an empty text from a formula counts)."""
    n = 0
    for arg in args:
        if isinstance(arg, Ref):
            n += sum(1 for _ in arg.cells())
        elif isinstance(arg, tuple):
            n += sum(len(row) for row in arg)
        elif arg is not MISSING:
            n += 1
    return float(n)


@function("COUNTBLANK", "raw", category=S)
def COUNTBLANK(ctx: Context, ref: Any) -> float | ErrorValue:
    """Blank cells and empty texts."""
    if not isinstance(ref, Ref):
        return ErrorValue.VALUE
    filled = sum(1 for _, _, v in ref.cells() if v is not BLANK and v != "")
    return float(ref.rows * ref.columns - filled)


# Spread, order and the A forms (capture ``statistics``) ----------------------------------------


def _values_a(
    args: tuple[Any, ...], ctx: Context, *, arrays_too: bool = False
) -> list[float] | ErrorValue:
    """The values of the A forms: in references text is 0 and logicals 1 or 0, blanks skipped."""
    out: list[float] = []
    for arg in args:
        if isinstance(arg, Ref):
            for _, _, v in arg.cells():
                if isinstance(v, ErrorValue):
                    return v
                if isinstance(v, bool):
                    out.append(float(v))
                elif isinstance(v, float):
                    out.append(v)
                elif isinstance(v, str):
                    out.append(0.0)
        elif isinstance(arg, tuple):
            # an array constant: numbers only, except for AVERAGEA, which counts text as 0 and
            # logicals as 1 or 0 there too (captured)
            for row in arg:
                for v in row:
                    if isinstance(v, ErrorValue):
                        return v
                    if type(v) is float:
                        out.append(v)
                    elif arrays_too and isinstance(v, bool | str):
                        out.append(float(v) if isinstance(v, bool) else 0.0)
        else:
            got = _numbers((arg,), ctx)
            if isinstance(got, ErrorValue):
                return got
            out.extend(got)
    return out


def _variance(values: list[float], sample: bool) -> float | ErrorValue:
    n = len(values)
    if n < (2 if sample else 1):
        return DIV0
    mean = 0.0
    for v in values:
        mean += v
    mean /= n
    ss = 0.0
    for v in values:
        d = v - mean
        ss += d * d
    result = ss / (n - 1 if sample else n)
    return result if math.isfinite(result) else NUM


def _spread(name: str, collect: Any, sample: bool, root: bool) -> None:
    @function(name, "raw", min_args=1, variadic=True, category=S, ulps=4, ulps_reason=_VARIANCE)
    def spread(ctx: Context, *args: Any) -> float | ErrorValue:
        values = collect(args, ctx)
        if isinstance(values, ErrorValue):
            return values
        var = _variance(values, sample)
        return math.sqrt(var) if root and not isinstance(var, ErrorValue) else var

    kind = ("standard deviation" if root else "variance") + (
        " of a sample" if sample else " of a population"
    )
    spread.__doc__ = kind[0].upper() + kind[1:] + "."


for _name, _collect, _sample, _root in (
    ("STDEV", _numbers, True, True),
    ("_xlfn.STDEV.S", _numbers, True, True),
    ("STDEVP", _numbers, False, True),
    ("_xlfn.STDEV.P", _numbers, False, True),
    ("VAR", _numbers, True, False),
    ("_xlfn.VAR.S", _numbers, True, False),
    ("VARP", _numbers, False, False),
    ("_xlfn.VAR.P", _numbers, False, False),
    ("STDEVA", _values_a, True, True),
    ("STDEVPA", _values_a, False, True),
    ("VARA", _values_a, True, False),
    ("VARPA", _values_a, False, False),
):
    _spread(_name, _collect, _sample, _root)


@function("AVERAGEA", "raw", min_args=1, variadic=True, category=S)
def AVERAGEA(ctx: Context, *args: Any) -> float | ErrorValue:
    """Mean counting text in references as 0 and logicals as 1 or 0."""
    values = _values_a(args, ctx, arrays_too=True)
    if isinstance(values, ErrorValue):
        return values
    if not values:
        return DIV0
    total = _total(values, ctx)
    return total if isinstance(total, ErrorValue) else total / len(values)


@function("MINA", "raw", min_args=1, variadic=True, category=S)
def MINA(ctx: Context, *args: Any) -> float | ErrorValue:
    """Smallest value counting text in references as 0 and logicals as 1 or 0."""
    values = _values_a(args, ctx)
    return values if isinstance(values, ErrorValue) else (min(values) if values else 0.0)


@function("MAXA", "raw", min_args=1, variadic=True, category=S)
def MAXA(ctx: Context, *args: Any) -> float | ErrorValue:
    """Largest value counting text in references as 0 and logicals as 1 or 0."""
    values = _values_a(args, ctx)
    return values if isinstance(values, ErrorValue) else (max(values) if values else 0.0)


@function("MEDIAN", "raw", min_args=1, variadic=True, category=S)
def MEDIAN(ctx: Context, *args: Any) -> float | ErrorValue:
    """The middle number, or the mean of the two middle ones."""
    values = _numbers(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    if not values:
        return NUM
    values.sort()
    n = len(values)
    if n % 2:
        return values[n // 2]
    return (values[n // 2 - 1] + values[n // 2]) / 2


@function("SUMSQ", "raw", min_args=1, variadic=True, category=S)
def SUMSQ(ctx: Context, *args: Any) -> float | ErrorValue:
    """Sum of the squares of the numbers."""
    values = _numbers(args, ctx)
    if isinstance(values, ErrorValue):
        return values
    return _total([v * v for v in values], ctx)


def _cells(arg: Any) -> list[Any]:
    if isinstance(arg, Ref):
        return list(to_array_values(arg))
    if isinstance(arg, tuple):
        return [v for row in arg for v in row]
    return [arg]


def to_array_values(ref: Ref) -> list[Any]:
    return [v for row in to_array(ref) for v in row]


def _plain(values: list[float]) -> float:
    s = 0.0
    for v in values:
        s += v
    return s


def _pairs(xs: Any, ys: Any) -> tuple[list[float], list[float]] | ErrorValue:
    """Paired numbers of two ranges of the same size; other pairs are left out."""
    a, b = _cells(xs), _cells(ys)
    if len(a) != len(b):
        return ErrorValue.NA
    px: list[float] = []
    py: list[float] = []
    for x, y in zip(a, b, strict=True):
        if isinstance(x, ErrorValue):
            return x
        if isinstance(y, ErrorValue):
            return y
        if type(x) is float and type(y) is float:
            px.append(x)
            py.append(y)
    return px, py


@function("CORREL", "array", "array", category=S, ulps=4, ulps_reason=_REGRESSION)
def CORREL(ctx: Context, xs: Any, ys: Any) -> float | ErrorValue:
    """Pearson correlation of two ranges."""
    pairs = _pairs(xs, ys)
    if isinstance(pairs, ErrorValue):
        return pairs
    px, py = pairs
    n = len(px)
    if n < 2:
        return DIV0
    mx, my = _plain(px) / n, _plain(py) / n
    sxy = sxx = syy = 0.0
    for x, y in zip(px, py, strict=True):
        dx, dy = x - mx, y - my
        sxy += dx * dy
        sxx += dx * dx
        syy += dy * dy
    if sxx == 0 or syy == 0:
        return DIV0
    return sxy / math.sqrt(sxx * syy)


@function(
    "LINEST", "array", "array", "logical", "logical", min_args=1, category=S,
    ulps=4, ulps_reason=_REGRESSION, status="partial", note="one regressor only",
)  # fmt: skip
def LINEST(ctx: Context, ys: Any, xs: Any = MISSING, const: Any = True, stats: Any = False) -> Any:
    """Least-squares line through the points: slope and intercept, with statistics."""
    yv = _cells(ys)
    xv = [float(i) for i in range(1, len(yv) + 1)] if xs is MISSING else _cells(xs)
    if len(xv) != len(yv):
        return ErrorValue.REF
    for v in (*yv, *xv):
        if isinstance(v, ErrorValue):
            return v
        if type(v) is not float:
            return VALUE
    const = True if const is MISSING else bool(const)
    stats = False if stats is MISSING else bool(stats)
    n = len(yv)
    if const:
        mx, my = _plain(xv) / n, _plain(yv) / n
    else:
        mx = my = 0.0
    sxx = sxy = 0.0
    for x, y in zip(xv, yv, strict=True):
        sxx += (x - mx) * (x - mx)
        sxy += (x - mx) * (y - my)
    if sxx == 0:
        slope, intercept = 0.0, my
    else:
        slope = sxy / sxx
        intercept = my - slope * mx if const else 0.0
    if not stats:
        return ((slope, intercept),)
    df = n - (2 if const else 1)
    ssreg = ssresid = 0.0
    for x, y in zip(xv, yv, strict=True):
        fit = slope * x + intercept
        ssresid += (y - fit) * (y - fit)
        ssreg += (fit - my) * (fit - my)
    na = ErrorValue.NA
    if df <= 0:
        return ((slope, intercept), (na, na), (1.0, 0.0), (NUM, df), (ssreg, ssresid))
    sey = math.sqrt(ssresid / df)
    se_slope = sey / math.sqrt(sxx) if sxx else NUM
    se_int: Any = sey * math.sqrt(1 / n + mx * mx / sxx) if const and sxx else na
    total = ssreg + ssresid
    r2 = ssreg / total if total else 1.0
    f: Any = ssreg / (ssresid / df) if ssresid else NUM  # one regressor: 1 degree of freedom
    return (
        (slope, intercept),
        (se_slope, se_int),
        (r2, sey),
        (f, float(df)),
        (ssreg, ssresid),
    )


def _rank(ctx: Context, x: float, ref: Any, order: float, average: bool) -> float | ErrorValue:
    if not isinstance(ref, Ref):
        return VALUE
    values = []
    for _, _, v in ref.cells():
        if isinstance(v, ErrorValue):
            return v
        if type(v) is float:
            values.append(v)
    if x not in values:
        return ErrorValue.NA
    before = sum(1 for v in values if (v < x if order else v > x))
    if not average:
        return float(before + 1)
    ties = sum(1 for v in values if v == x)
    return before + (ties + 1) / 2


@function("RANK", "number", "raw", "number", min_args=2, category=S)
@function("_xlfn.RANK.EQ", "number", "raw", "number", min_args=2, category=S)
def RANK(ctx: Context, x: float, ref: Any, order: float = 0.0) -> float | ErrorValue:
    """Rank of a number among the numbers of a range; ties share the best rank."""
    return _rank(ctx, x, ref, order, False)


@function("_xlfn.RANK.AVG", "number", "raw", "number", min_args=2, category=S)
def RANK_AVG(ctx: Context, x: float, ref: Any, order: float = 0.0) -> float | ErrorValue:
    """Rank of a number; ties share the mean of their ranks."""
    return _rank(ctx, x, ref, order, True)


# The normal distribution -----------------------------------------------------------------------

_SQRT2 = math.sqrt(2.0)
_SQRT2PI = math.sqrt(2.0 * math.pi)


def _phi(z: float) -> float:
    return 0.5 * math.erfc(-z / _SQRT2)


def _density(z: float) -> float:
    return math.exp(-0.5 * z * z) / _SQRT2PI


def _ppnd(p: float) -> float:
    """The standard normal quantile (Wichura, AS241 PPND16)."""
    q = p - 0.5
    if abs(q) <= 0.425:
        r = 0.180625 - q * q
        num = (((((((2509.0809287301226727 * r + 33430.575583588128105) * r
                 + 67265.770927008700853) * r + 45921.953931549871457) * r
                 + 13731.693765509461125) * r + 1971.5909503065514427) * r
                 + 133.14166789178437745) * r + 3.387132872796366608)  # fmt: skip
        den = (((((((5226.495278852545925 * r + 28729.085735721942674) * r
                 + 39307.89580009271061) * r + 21213.794301586595867) * r
                 + 5394.1960214247511077) * r + 687.1870074920579083) * r
                 + 42.313330701600911252) * r + 1.0)  # fmt: skip
        return q * num / den
    r = p if q < 0 else 1 - p
    r = math.sqrt(-math.log(r))
    if r <= 5:
        r -= 1.6
        num = (((((((7.7454501427834140764e-4 * r + 0.0227238449892691845833) * r
                 + 0.24178072517745061177) * r + 1.27045825245236838258) * r
                 + 3.64784832476320460504) * r + 5.7694972214606914055) * r
                 + 4.6303378461565452959) * r + 1.42343711074968357734)  # fmt: skip
        den = (((((((1.05075007164441684324e-9 * r + 5.475938084995344946e-4) * r
                 + 0.0151986665636164571966) * r + 0.14810397642748007459) * r
                 + 0.68976733498510000455) * r + 1.6763848301838038494) * r
                 + 2.05319162663775882187) * r + 1.0)  # fmt: skip
    else:
        r -= 5
        num = (((((((2.01033439929228813265e-7 * r + 2.71155556874348757815e-5) * r
                 + 0.0012426609473880784386) * r + 0.026532189526576123093) * r
                 + 0.29656057182850489123) * r + 1.7848265399172913358) * r
                 + 5.4637849111641143699) * r + 6.6579046435011037772)  # fmt: skip
        den = (((((((2.04426310338993978564e-15 * r + 1.4215117583164458887e-7) * r
                 + 1.8463183175100546818e-5) * r + 7.868691311456132591e-4) * r
                 + 0.0148753612908506148525) * r + 0.13692988092273580531) * r
                 + 0.59983220655588793769) * r + 1.0)  # fmt: skip
    x = num / den
    return -x if q < 0 else x


@function("NORMSDIST", "number", category=S, ulps=4, ulps_reason=_NORMAL)
def NORMSDIST(ctx: Context, z: float) -> float:
    """The standard normal cumulative distribution."""
    return _phi(z)


@function("_xlfn.NORM.S.DIST", "number", "logical", category=S, ulps=4, ulps_reason=_NORMAL)
def NORM_S_DIST(ctx: Context, z: float, cumulative: bool) -> float:
    """The standard normal distribution: cumulative or density."""
    return _phi(z) if cumulative else _density(z)


@function(
    "NORMDIST", "number", "number", "number", "logical", category=S, ulps=4, ulps_reason=_NORMAL
)
@function(
    "_xlfn.NORM.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    ulps=4,
    ulps_reason=_NORMAL,
)
def NORMDIST(ctx: Context, x: float, mean: float, sd: float, cumulative: bool) -> Any:
    """The normal distribution with a mean and standard deviation."""
    if sd <= 0:
        return NUM
    z = (x - mean) / sd
    return _phi(z) if cumulative else _density(z) / sd


@function("NORMSINV", "number", category=S, ulps=4, ulps_reason=_NORMAL)
@function("_xlfn.NORM.S.INV", "number", category=S, ulps=4, ulps_reason=_NORMAL)
def NORMSINV(ctx: Context, p: float) -> Any:
    """The inverse of the standard normal cumulative distribution."""
    if not 0 < p < 1:
        return NUM
    return _ppnd(p)


@function("NORMINV", "number", "number", "number", category=S, ulps=4, ulps_reason=_NORMAL)
@function("_xlfn.NORM.INV", "number", "number", "number", category=S, ulps=4, ulps_reason=_NORMAL)
def NORMINV(ctx: Context, p: float, mean: float, sd: float) -> Any:
    """The inverse of the normal cumulative distribution."""
    if not 0 < p < 1 or sd <= 0:
        return NUM
    return mean + sd * _ppnd(p)


def _array_numbers(arg: Any, ctx: Context) -> list[float] | ErrorValue:
    """The numbers of ``LARGE``'s and ``PERCENTILE``'s array: cells or elements (capture modern).

    In a range or an array only numbers count, and an error stops; a single value is converted,
    as an argument is, so ``"5"`` and TRUE count, and ``"a"`` is ``#VALUE!``.
    """
    if isinstance(arg, Ref):
        elements: Any = arg.values()
    elif isinstance(arg, tuple):
        elements = (v for row in arg for v in row)
    else:
        v = to_number(arg, ctx)
        return v if isinstance(v, ErrorValue) else [v]
    out = []
    for v in elements:
        if isinstance(v, ErrorValue):
            return v
        if isinstance(v, float):
            out.append(v)
    return out


def _at_rank(a: Any, k: float, ctx: Context, largest: bool) -> float | ErrorValue:
    values = _array_numbers(a, ctx)
    if isinstance(values, ErrorValue):
        return values
    n = math.ceil(k) if largest else math.floor(k)  # capture modern: LARGE up, SMALL down
    if n < 1 or n > len(values):
        return NUM
    return sorted(values, reverse=largest)[n - 1]


@function("LARGE", "raw", "number", category=S)
def LARGE(ctx: Context, a: Any, k: float) -> float | ErrorValue:
    """The ``k``th largest number, ``k`` rounded up."""
    return _at_rank(a, k, ctx, True)


@function("SMALL", "raw", "number", category=S)
def SMALL(ctx: Context, a: Any, k: float) -> float | ErrorValue:
    """The ``k``th smallest number, ``k`` rounded down."""
    return _at_rank(a, k, ctx, False)


def _percentile(a: Any, k: float, ctx: Context) -> float | ErrorValue:
    values = _array_numbers(a, ctx)
    if isinstance(values, ErrorValue):
        return values
    if not values or k < 0 or k > 1:
        return NUM
    values.sort()
    rank = k * (len(values) - 1) + 1  # 1-based, as Excel takes it: the fraction's last bits show
    low = math.floor(rank)
    if low == rank:
        return values[low - 1]
    return values[low - 1] + (rank - low) * (values[low] - values[low - 1])


@function("PERCENTILE", "raw", "number", category=S)
@function("_xlfn.PERCENTILE.INC", "raw", "number", category=S)
def PERCENTILE(ctx: Context, a: Any, k: float) -> float | ErrorValue:
    """The value at fraction ``k`` of the sorted numbers, interpolating between neighbors."""
    return _percentile(a, k, ctx)


@function("QUARTILE", "raw", "number", category=S)
@function("_xlfn.QUARTILE.INC", "raw", "number", category=S)
def QUARTILE(ctx: Context, a: Any, q: float) -> float | ErrorValue:
    """The ``q``th quartile, 0 to 4, ``q`` truncated."""
    n = math.trunc(q)
    return _percentile(a, n / 4, ctx) if 0 <= n <= 4 else NUM


@function("MODE", "array", min_args=1, variadic=True, errors=True, category=S)
@function("_xlfn.MODE.SNGL", "array", min_args=1, variadic=True, errors=True, category=S)
def MODE(ctx: Context, *args: Any) -> float | ErrorValue:
    """The number occurring most often, the first of those tied; ``#N/A`` if none repeats.

    Text given directly is skipped, as in a range, and an omitted argument is ``#VALUE!``
    (capture ``modern``).
    """
    counts: dict[float, int] = {}
    for arg in args:
        if arg is MISSING:
            return VALUE
        if isinstance(arg, Ref):
            elements: Any = arg.values()
        elif isinstance(arg, tuple):
            elements = (v for row in arg for v in row)
        else:
            elements = (float(arg) if isinstance(arg, bool) else arg,)
        for v in elements:
            if isinstance(v, ErrorValue):
                return v
            if isinstance(v, float):
                counts[v] = counts.get(v, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        return NA
    return next(v for v, n in counts.items() if n == best)
