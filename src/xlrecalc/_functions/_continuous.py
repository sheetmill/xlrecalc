"""The beta, gamma, chi-squared, F and t distributions, their inverses and tests.

Captured from Excel (``continuous_family``). The cumulative functions rest on the regularized
incomplete beta and gamma functions, each tail computed directly where it is the smaller; an
inverse is found by bisection on the smaller tail. At large shapes, where logarithms of the
gamma function cancel, Loader's saddle point, DiDonato and Morris's expansion (BGRAT) and
Temme's uniform expansion take over. Each function declares a relative tolerance, its largest
difference from Excel on the capture rounded up to a power of ten (``tools/relative_spread.py``).
"""

from __future__ import annotations

import functools
import math
import struct
import sys
from typing import Any

from xlrecalc._arrays import is_array, shape
from xlrecalc._coerce import numbers
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._functions._statistical import _phi
from xlrecalc._values import ErrorValue, Ref

S = "statistical"
_REL_TAILS = (
    "the incomplete beta and gamma functions, by continued fractions, series and asymptotic "
    "expansions whose last bits are not Excel's (capture continuous_family)"
)
_REL_SEARCH = (
    "an inverse found by bisection to the last bit on a distribution whose last bits are not "
    "Excel's, where Excel's search stops within its own tolerance (capture continuous_family)"
)
_REL_TEST = (
    "a statistic from sums of squares, then a distribution whose last bits are not Excel's "
    "(capture continuous_family)"
)
NUM, DIV0, NA = ErrorValue.NUM, ErrorValue.DIV0, ErrorValue.NA
_MAX_DF = 1e10


# The incomplete beta and gamma functions ----------------------------------------------------


def _beta_fraction(x: float, a: float, b: float) -> float:
    """The continued fraction of the incomplete beta function (modified Lentz)."""
    tiny = 1e-300
    c, d = 1.0, 1.0 - (a + b) * x / (a + 1)
    d = 1 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, _FRACTION_TERMS):
        for numerator in (
            m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m)),
            -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1)),
        ):
            d = 1 + numerator * d
            d = 1 / (d if abs(d) > tiny else tiny)
            c = 1 + numerator / c
            c = c if abs(c) > tiny else tiny
            h *= d * c
        if abs(d * c - 1) < 1e-16:
            return h
    return math.nan  # not converged within its terms: no value


def _lgamma_diff(a: float, b: float) -> float:
    """Lgamma(a + b) - lgamma(a), by Stirling's expansion where ``a`` is large (they cancel)."""
    if a > 1e5 and abs(b) < 10:
        c = b * (b - 1)
        return (
            b * math.log(a)
            + c / (2 * a)
            - c * (2 * b - 1) / (12 * a * a)
            + c * c / (12 * a * a * a)
        )
    if a >= _SADDLE and b > 0:  # Stirling's form, the two large terms taken together
        return (
            (a - 0.5) * math.log1p(b / a) + b * math.log(a + b) - b
            + _stirling_rest(a + b) - _stirling_rest(a)
        )  # fmt: skip
    return _lgamma(a + b) - _lgamma(a)


def _stirling_rest(x: float) -> float:
    """lgamma(x) less Stirling's approximation, for x of 10 and more."""
    xx = x * x
    return (1 / 12 - (1 / 360 - (1 / 1260 - 1 / (1680 * xx)) / xx) / xx) / x


def _log_beta(a: float, b: float) -> float:
    """Log B(a, b) without cancelling: by Stirling's form where both are large.

    That form never adds a and b, which may overflow: log(a + b) is log(big) + log1p(small / big).
    """
    big, small = (a, b) if a >= b else (b, a)
    if small >= 10:
        log_total = math.log(big) + math.log1p(small / big)
        return (
            0.5 * math.log(2 * math.pi) - 0.5 * log_total
            + (a - 0.5) * _log_share(a, b) + (b - 0.5) * _log_share(b, a)
            + _stirling_rest(a) + _stirling_rest(b) - _stirling_rest(a + b)
        )  # fmt: skip
    return _lgamma(small) - _lgamma_diff(big, small)


def _log_share(part: float, other: float) -> float:
    """Log(part / (part + other)), from their ratio alone: the sum may overflow."""
    if part >= other:
        return -math.log1p(other / part)
    return math.log(part / other) - math.log1p(part / other)


_LN_SQRT_2PI = 0.5 * math.log(2 * math.pi)
_FRACTION_TERMS = 20_000  # a continued fraction's terms, at most: past them, no value (#NUM!)
_SERIES_TERMS = 20_000  # a series' terms, at most, likewise
_EXP_MAX = math.log(sys.float_info.max)


def _exp(x: float) -> float:
    """Exp, infinite past the largest double instead of raising."""
    return math.inf if x > _EXP_MAX else math.exp(x)


def _lgamma(x: float) -> float:
    """Lgamma, infinite past the largest double instead of raising."""
    try:
        return math.lgamma(x)
    except OverflowError:
        return math.inf


def _finite(core: Any) -> Any:
    """A function's value, #NUM! where it is not a finite number (an overflow, no convergence)."""

    @functools.wraps(core)
    def wrapped(*args: Any) -> Any:
        value = core(*args)
        return NUM if isinstance(value, float) and not math.isfinite(value) else value

    return wrapped


_SADDLE = 1e3  # shapes from which the saddle-point form replaces the logarithms (they cancel)
_SMALL_SHAPE = 30.0  # the other shape, at most, for BGRAT and its steps up
_MAX_Z = 600.0  # BGRAT's z, at most, so that its exp(-z) does not underflow


def _stirlerr(n: float) -> float:
    """Log(n!) less Stirling's approximation log(sqrt(2 pi n) (n/e)^n) (Loader)."""
    if n <= 15:
        return _lgamma(n + 1) - (n + 0.5) * math.log(n) + n - _LN_SQRT_2PI
    nn = n * n
    s0, s1, s2, s3, s4 = 1 / 12, 1 / 360, 1 / 1260, 1 / 1680, 1 / 1188
    if n > 500:
        return (s0 - s1 / nn) / n
    if n > 80:
        return (s0 - (s1 - s2 / nn) / nn) / n
    if n > 35:
        return (s0 - (s1 - (s2 - s3 / nn) / nn) / nn) / n
    return (s0 - (s1 - (s2 - (s3 - s4 / nn) / nn) / nn) / nn) / n


def _bd0(x: float, mean: float) -> float:
    """The deviance x log(x / mean) + mean - x, not cancelling for x near ``mean`` (Loader)."""
    if abs(x - mean) < 0.1 * (x + mean):
        v = (x - mean) / (x + mean)
        s, term, v2 = (x - mean) * v, 2 * x * v, v * v
        for j in range(1, 1000):
            term *= v2
            nxt = s + term / (2 * j + 1)
            if nxt == s:
                break
            s = nxt
        return s
    if x == 0:
        return mean
    if mean == 0 or mean == math.inf:  # x is no mean's neighbor: the term is 0
        return math.inf
    ratio = x / mean
    if 0 < ratio < math.inf:
        return x * math.log(ratio) + mean - x
    return x * (math.log(x) - math.log(mean)) + mean - x  # the ratio under- or overflows


def _binom_term(k: float, n: float, p: float, q: float) -> float:
    """C(n, k) p^k q^(n-k) for real 0 <= k <= n, by Loader's saddle point; q is 1 - p exactly."""
    if k == 0:
        return _exp(-_bd0(n, n * q) - n * p if p < 0.1 else n * math.log(q))
    if k == n:
        return _exp(-_bd0(n, n * p) - n * q if q < 0.1 else n * math.log(p))
    log_c = _stirlerr(n) - _stirlerr(k) - _stirlerr(n - k) - _bd0(k, n * p) - _bd0(n - k, n * q)
    log_f = 2 * _LN_SQRT_2PI + math.log(k) + math.log(n - k) - math.log(n)  # n - k is exact
    return _exp(log_c - 0.5 * log_f)


def _bgrat(a: float, b: float, x: float, y: float) -> float:
    """I_x(a, b) for a large and 0 < b <= 1, x near 1 (DiDonato and Morris, TOMS 708 BGRAT).

    Driven by log x = log1p(-y), so the rounding of x, which a large ``a`` magnifies, never enters.
    """
    bm1 = b - 1
    nu = a + bm1 / 2
    lnx = math.log(x) if y > 0.375 else math.log1p(-y)
    z = -nu * lnx
    log_r = b * math.log(z) - z - _lgamma(b)
    if log_r < -_EXP_MAX:  # r underflows (a subnormal shape): no value here
        return math.nan
    log_u = log_r + _lgamma_diff(a, b) - b * math.log(nu)
    j = total = _igamma_upper(b, z) / _exp(log_r)
    v, t2 = 0.25 / (nu * nu), 0.25 * lnx * lnx
    t = cn = 1.0
    n2 = 0.0
    c: list[float] = []
    d: list[float] = []
    for n in range(1, 31):
        bp2n = b + n2
        j = (bp2n * (bp2n + 1) * j + (z + bp2n + 1) * t) * v
        n2 += 2
        t *= t2
        cn /= n2 * (n2 + 1)
        c.append(cn)
        s, coef = 0.0, b - n
        for i in range(1, n):
            s += coef * c[i - 1] * d[n - 1 - i]
            coef += b
        d.append(bm1 * cn + s / n)
        dj = d[-1] * j
        total += dj
        if abs(dj) <= 1e-17 * total:
            return _exp(log_u) * total
    return math.nan


def _ibeta_large(x: float, a: float, b: float, y: float) -> float:
    """I_x(a, b) for a large and b small, x near 1: BGRAT at b's fraction, then up to b exactly.

    I_x(a, b + 1) = I_x(a, b) + x^a y^b / (b B(a, b)), every term positive.
    """
    steps = math.ceil(b) - 1
    base = b - steps
    total = _bgrat(a, base, x, y)
    for i in range(steps):
        bi = base + i
        total += _binom_term(a, a + bi, x, y) * (a / (a + bi))
    return total


def _ibeta(x: float, a: float, b: float, y: float | None = None) -> float:
    """The regularized incomplete beta function I_x(a, b); ``y`` is 1 - x, given exactly."""
    y = 1 - x if y is None else y
    if x <= 0:
        return 0.0
    if y <= 0:
        return 1.0
    if max(a, b) >= _SADDLE:  # where x^a y^b / B(a, b) underflows, the tails are 0 and 1
        log_x = math.log1p(-y) if y < 0.5 else math.log(x)
        log_y = math.log1p(-x) if x < 0.5 else math.log(y)
        if a * log_x + b * log_y - _log_beta(a, b) < -_EXP_MAX - 50:
            return 0.0 if x < a / (a + b) else 1.0
    if a >= _SADDLE and b <= _SMALL_SHAPE and y < 0.3 and -a * math.log1p(-y) < _MAX_Z:
        value = _ibeta_large(x, a, b, y)
        if value == value:  # else the fraction below
            return value
    if b >= _SADDLE and a <= _SMALL_SHAPE and x < 0.3 and -b * math.log1p(-x) < _MAX_Z:
        other = _ibeta_large(y, b, a, x)
        if other <= 0.5:  # 1 - other keeps its digits; a small result is left to the fraction
            return 1 - other
    if max(a, b) >= _SADDLE:  # x^a y^b / B(a, b), whose logarithms cancel at large shapes
        front = _binom_term(a, a + b, x, y) * (a * b / (a + b))
    else:
        log_x = math.log1p(-y) if y < 0.5 else math.log(x)
        log_y = math.log1p(-x) if x < 0.5 else math.log(y)
        front = _exp(a * log_x + b * log_y - _log_beta(a, b))
    if x < (a + 1) / (a + b + 2):
        return front * _beta_fraction(x, a, b) / a
    return 1 - front * _beta_fraction(y, b, a) / b


def _igamma(a: float, x: float) -> float:
    """The regularized lower incomplete gamma function P(a, x)."""
    return _gamma_tails(a, x)[0]


def _igamma_upper(a: float, x: float) -> float:
    """The regularized upper incomplete gamma function Q(a, x), not as 1 - P in its tail."""
    return _gamma_tails(a, x)[1]


# Temme's coefficients C0, C1, C2 of the uniform expansion in z (DLMF 8.12), lowest power first.
_TEMME = (
    (-1 / 3, 1 / 12, -2 / 135, 1 / 864, 1 / 2835, -139 / 777600),
    (-1 / 540, -1 / 288, 1 / 378, -77 / 77760, 1 / 4860),
    (25 / 6048, -139 / 51840, 1 / 1296),
)
_TEMME_SHAPE = 1e5  # shapes from which the uniform expansion replaces the series and fraction


def _minus_log1p(s: float) -> float:
    """S - log1p(s), by its series where it would cancel."""
    if abs(s) >= 0.1:
        return s - math.log1p(s)
    total, power = 0.0, -s  # the sum of (-s)^k / k from k = 2
    for k in range(2, 60):
        power *= -s
        term = power / k
        total += term
        if abs(term) < 1e-17 * total:
            break
    return total


def _gamma_temme(a: float, x: float) -> tuple[float, float]:
    """P(a, x) and Q(a, x) for a large and x near a, by Temme's uniform expansion.

    The smaller tail is erfc(sqrt(a phi)) / 2 and a correction in powers of 1/a, with
    phi = sigma - log1p(sigma) and sigma = (x - a) / a: no lgamma of a and no long series.
    """
    phi = _minus_log1p((x - a) / a)
    y = a * phi
    z = math.sqrt(2 * phi) * (-1 if x < a else 1)
    c = [sum(k * z**i for i, k in enumerate(row)) for row in _TEMME]
    rest = (c[0] + (c[1] + c[2] / a) / a) * _exp(-y) / math.sqrt(2 * math.pi * a)
    if x < a:
        lower = math.erfc(math.sqrt(y)) / 2 - rest
        return lower, 1 - lower
    upper = math.erfc(math.sqrt(y)) / 2 + rest
    return 1 - upper, upper


def _gamma_tails(a: float, x: float) -> tuple[float, float]:
    """P(a, x) and Q(a, x), each computed directly where it is the smaller."""
    if x <= 0:
        return 0.0, 1.0
    if a >= _TEMME_SHAPE and abs(x - a) < 0.1 * a:
        return _gamma_temme(a, x)
    log_front = a * math.log(x) - x - _lgamma(a)
    if log_front < -_EXP_MAX - 50:  # the smaller tail underflows: the limits
        return (0.0, 1.0) if x < a else (1.0, 0.0)
    if x < a + 1:
        term = total = 1 / a
        n = a
        for _ in range(_SERIES_TERMS):
            n += 1
            term *= x / n
            total += term
            if abs(term) < abs(total) * 1e-17:
                break
        else:
            return math.nan, math.nan
        lower = min(1.0, total * _exp(log_front))
        return lower, 1 - lower
    tiny = 1e-300
    b = x + 1 - a
    c, d = 1 / tiny, 1 / b
    h = d
    for i in range(1, _FRACTION_TERMS):
        an = -i * (i - a)
        b += 2
        d = an * d + b
        d = 1 / (d if abs(d) > tiny else tiny)
        c = b + an / c
        c = c if abs(c) > tiny else tiny
        h *= d * c
        if abs(d * c - 1) < 1e-16:
            break
    else:
        return math.nan, math.nan
    upper = _exp(log_front) * h
    return max(0.0, 1 - upper), upper


def _bits(x: float) -> int:
    return int(struct.unpack("<q", struct.pack("<d", x))[0])


def _from_bits(n: int) -> float:
    return float(struct.unpack("<d", struct.pack("<q", n))[0])


def _invert(cdf: Any, p: float, lo: float, hi: float) -> float:
    """The x in [lo, ...) with ``cdf(x) == p``, to the last bit; NaN where ``cdf`` gives none.

    Bisection over the doubles' bit patterns, which order as the doubles do from 0 up: at most
    64 evaluations reach two adjacent doubles, wherever the root lies. Where ``cdf(hi)`` is
    below ``p`` the interval reaches to 1E300.
    """
    top = cdf(hi)
    if top != top:
        return math.nan
    if top < p:  # as far as 1E300, where products such as d1 * x stay finite
        lo, hi = hi, 1e300
    low, high = _bits(lo), _bits(hi)
    while high - low > 1:
        mid = (low + high) // 2
        value = cdf(_from_bits(mid))
        if value != value:
            return math.nan
        if value < p:
            low = mid
        else:
            high = mid
    return (_from_bits(low) + _from_bits(high)) / 2


# Beta ---------------------------------------------------------------------------------------


def _beta_args(a: float, b: float, lo: float, hi: float) -> bool:
    return a > 0 and b > 0 and lo < hi


@function("_xlfn.BETA.DIST", "number", "number", "number", "logical", "number", "number",
          min_args=4, category=S, rel=1e-13, rel_reason=_REL_TAILS)  # fmt: skip
@_finite
def BETA_DIST(
    ctx: Any, x: float, a: float, b: float, cumulative: bool, lo: float = 0.0, hi: float = 1.0
) -> Any:
    """The beta distribution over ``lo`` to ``hi``."""
    if not _beta_args(a, b, lo, hi) or not lo <= x <= hi:
        return NUM
    z = (x - lo) / (hi - lo)
    if cumulative:
        return _ibeta(z, a, b)
    if z == 0:  # Excel: 0 for a shape above 1, #NUM! at 1 and below
        return 0.0 if a > 1 else NUM
    if z == 1:  # the same at the other end (capture continuous_family)
        return 0.0 if b > 1 else NUM
    if (
        a > 1 and b > 1 and max(a, b) >= _SADDLE
    ):  # z^(a-1) (1-z)^(b-1) / B(a, b), as a binomial term
        return _binom_term(a - 1, a + b - 2, z, 1 - z) * (a + b - 1) / (hi - lo)
    big = max(a, b) >= _SADDLE  # where the logarithms of the gamma function cancel
    log_b = _log_beta(a, b) if big else _lgamma(a) + _lgamma(b) - _lgamma(a + b)
    density = _exp((a - 1) * math.log(z) + (b - 1) * math.log1p(-z) - log_b)
    return density / (hi - lo)


@function(
    "BETADIST",
    "number",
    "number",
    "number",
    "number",
    "number",
    min_args=3,
    category=S,
    rel=1e-14,
    rel_reason=_REL_TAILS,
)
@_finite
def BETADIST(ctx: Any, x: float, a: float, b: float, lo: float = 0.0, hi: float = 1.0) -> Any:
    """The cumulative beta distribution."""
    return BETA_DIST(ctx, x, a, b, True, lo, hi)


@function(
    "BETAINV",
    "number",
    "number",
    "number",
    "number",
    "number",
    min_args=3,
    category=S,
    rel=1e-14,
    rel_reason=_REL_SEARCH,
)
@function("_xlfn.BETA.INV", "number", "number", "number", "number", "number", min_args=3,
          category=S, rel=1e-13, rel_reason=_REL_SEARCH)  # fmt: skip
@_finite
def BETA_INV(ctx: Any, p: float, a: float, b: float, lo: float = 0.0, hi: float = 1.0) -> Any:
    """The inverse of the cumulative beta distribution."""
    if not _beta_args(a, b, lo, hi) or not 0 < p < 1:
        return NUM
    if p > 0.5:  # the upper tail I_{1-z}(b, a), where it is the smaller
        return hi - (hi - lo) * _invert(lambda y: _ibeta(y, b, a), 1 - p, 0.0, 1.0)
    return lo + (hi - lo) * _invert(lambda z: _ibeta(z, a, b), p, 0.0, 1.0)


# Gamma and chi-squared ----------------------------------------------------------------------


@function(
    "GAMMADIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    rel=1e-14,
    rel_reason=_REL_TAILS,
)
@function(
    "_xlfn.GAMMA.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    rel=1e-13,
    rel_reason=_REL_TAILS,
)
@_finite
def GAMMA_DIST(ctx: Any, x: float, alpha: float, beta: float, cumulative: bool) -> Any:
    """The gamma distribution with shape ``alpha`` and scale ``beta``."""
    if x < 0 or alpha <= 0 or beta <= 0:
        return NUM
    if cumulative:
        return _igamma(alpha, x / beta)
    if x == 0:  # Excel: 0 for a shape above 1, #NUM! at 1 and below
        return 0.0 if alpha > 1 else NUM
    if alpha >= _SADDLE:  # a Poisson term by Loader's saddle point (R's dgamma)
        k, mean = alpha - 1, x / beta
        log_term = -_stirlerr(k) - _bd0(k, mean)
        return _exp(log_term) / math.sqrt(2 * math.pi * k) / beta
    return _exp((alpha - 1) * math.log(x) - x / beta - _lgamma(alpha) - alpha * math.log(beta))


@function("GAMMAINV", "number", "number", "number", category=S, rel=1e-15, rel_reason=_REL_SEARCH)
@function(
    "_xlfn.GAMMA.INV", "number", "number", "number", category=S, rel=1e-13, rel_reason=_REL_SEARCH
)
@_finite
def GAMMA_INV(ctx: Any, p: float, alpha: float, beta: float) -> Any:
    """The inverse of the cumulative gamma distribution."""
    if not 0 <= p < 1 or alpha <= 0 or beta <= 0:
        return NUM
    if p == 0:
        return 0.0
    if p > 0.5:  # the upper tail, where it is the smaller
        return beta * _invert(lambda x: -_igamma_upper(alpha, x), -(1 - p), 0.0, max(1.0, alpha))
    return beta * _invert(lambda x: _igamma(alpha, x), p, 0.0, max(1.0, alpha))


def _df(df: float) -> int | None:
    k = math.trunc(df)
    return k if 1 <= k <= _MAX_DF else None


@function(
    "_xlfn.CHISQ.DIST", "number", "number", "logical", category=S, rel=1e-12, rel_reason=_REL_TAILS
)
@_finite
def CHISQ_DIST(ctx: Any, x: float, df: float, cumulative: bool) -> Any:
    """The chi-squared distribution."""
    k = _df(df)
    if k is None or x < 0:
        return NUM
    if x == 0 and not cumulative:  # 0.5 at 2 degrees of freedom, as Excel gives it
        return 0.5 if k == 2 else 0.0 if k > 2 else NUM
    return GAMMA_DIST(ctx, x, k / 2, 2.0, cumulative)


@function("CHIDIST", "number", "number", category=S, rel=1e-14, rel_reason=_REL_TAILS)
@function("_xlfn.CHISQ.DIST.RT", "number", "number", category=S, rel=1e-13, rel_reason=_REL_TAILS)
@_finite
def CHISQ_DIST_RT(ctx: Any, x: float, df: float) -> Any:
    """The right tail of the chi-squared distribution."""
    k = _df(df)
    if k is None or x < 0:
        return NUM
    return _igamma_upper(k / 2, x / 2)


@function("_xlfn.CHISQ.INV", "number", "number", category=S, rel=1e-13, rel_reason=_REL_SEARCH)
@_finite
def CHISQ_INV(ctx: Any, p: float, df: float) -> Any:
    """The inverse of the left tail of the chi-squared distribution."""
    k = _df(df)
    if k is None or not 0 <= p < 1:
        return NUM
    return GAMMA_INV(ctx, p, k / 2, 2.0)


@function("CHIINV", "number", "number", category=S, rel=1e-14, rel_reason=_REL_SEARCH)
@function("_xlfn.CHISQ.INV.RT", "number", "number", category=S, rel=1e-13, rel_reason=_REL_SEARCH)
@_finite
def CHISQ_INV_RT(ctx: Any, p: float, df: float) -> Any:
    """The inverse of the right tail of the chi-squared distribution."""
    k = _df(df)
    if k is None or not 0 < p <= 1:
        return NUM
    if p == 1:
        return 0.0
    if k == 1:  # Excel inverts the left tail at 1 - p here, rounded: 1E-17 is #NUM!, 1E-10 is
        p = 1 - (1 - p)  # read as 1.0000000827E-10; exact from 0.5 on (capture continuous_family)
        if p == 0:
            return NUM
    if p > 0.5:  # the left tail, where it is the smaller
        return 2 * _invert(lambda x: _igamma(k / 2, x), 1 - p, 0.0, max(1.0, k / 2))
    return 2 * _invert(lambda x: -_igamma_upper(k / 2, x), -p, 0.0, max(1.0, k / 2))


# F ------------------------------------------------------------------------------------------


def _f_cdf(x: float, d1: float, d2: float) -> float:
    if x <= 0:
        return 0.0
    return _ibeta(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2, d2 / (d1 * x + d2))


def _f_right(x: float, d1: float, d2: float) -> float:
    if x <= 0:
        return 1.0
    return _ibeta(d2 / (d2 + d1 * x), d2 / 2, d1 / 2, d1 * x / (d2 + d1 * x))


@function(
    "_xlfn.F.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    rel=1e-12,
    rel_reason=_REL_TAILS,
)
@_finite
def F_DIST(ctx: Any, x: float, d1: float, d2: float, cumulative: bool) -> Any:
    """The F distribution."""
    m, n = _df(d1), _df(d2)
    if m is None or n is None or x < 0:
        return NUM
    if cumulative:
        return _f_cdf(x, m, n)
    if x == 0:
        return 0.0 if m > 2 else (1.0 if m == 2 else NUM)
    xm = x * m
    p, q = xm / (n + xm), n / (n + xm)
    if max(m, n) >= _SADDLE and math.isfinite(xm) and p > 0 and q > 0:  # Loader; R's df
        if m >= 2:
            return (m * q / 2) * _binom_term((m - 2) / 2, (m + n - 2) / 2, p, q)
        return (m * m * q / (2 * p * (m + n))) * _binom_term(m / 2, (m + n) / 2, p, q)
    log_b = _lgamma(m / 2) + _lgamma(n / 2) - _lgamma((m + n) / 2)
    log_density = (
        (m / 2) * math.log(m / n) + (m / 2 - 1) * math.log(x)
        - ((m + n) / 2) * math.log1p(m * x / n) - log_b
    )  # fmt: skip
    return _exp(log_density)


@function("FDIST", "number", "number", "number", category=S, rel=1e-14, rel_reason=_REL_TAILS)
@function(
    "_xlfn.F.DIST.RT", "number", "number", "number", category=S, rel=1e-12, rel_reason=_REL_TAILS
)
@_finite
def F_DIST_RT(ctx: Any, x: float, d1: float, d2: float) -> Any:
    """The right tail of the F distribution."""
    m, n = _df(d1), _df(d2)
    if m is None or n is None or x < 0:
        return NUM
    return _f_right(x, m, n)


@function(
    "_xlfn.F.INV", "number", "number", "number", category=S, rel=1e-12, rel_reason=_REL_SEARCH
)
@_finite
def F_INV(ctx: Any, p: float, d1: float, d2: float) -> Any:
    """The inverse of the F distribution's left tail."""
    m, n = _df(d1), _df(d2)
    if m is None or n is None or not 0 <= p < 1:
        return NUM
    if p == 0:
        return 0.0
    if p > 0.5:  # the right tail, where it is the smaller
        return _invert(lambda x: -_f_right(x, m, n), -(1 - p), 0.0, 1.0)
    return _invert(lambda x: _f_cdf(x, m, n), p, 0.0, 1.0)


@function("FINV", "number", "number", "number", category=S, rel=1e-14, rel_reason=_REL_SEARCH)
@function(
    "_xlfn.F.INV.RT", "number", "number", "number", category=S, rel=1e-13, rel_reason=_REL_SEARCH
)
@_finite
def F_INV_RT(ctx: Any, p: float, d1: float, d2: float) -> Any:
    """The inverse of the F distribution's right tail."""
    m, n = _df(d1), _df(d2)
    if m is None or n is None or not 0 < p <= 1:
        return NUM
    if p == 1:
        return 0.0
    if p > 0.5:  # the left tail, where it is the smaller
        return _invert(lambda x: _f_cdf(x, m, n), 1 - p, 0.0, 1.0)
    return _invert(lambda x: -_f_right(x, m, n), -p, 0.0, 1.0)  # the right tail itself


# t ------------------------------------------------------------------------------------------


def _t_tail(t: float, df: float) -> float:
    """P(T > t) for t >= 0, with 1 - x kept exact and t * t never overflowing."""
    if t == 0:
        return 0.5
    r = math.sqrt(df) / t  # x = df / (df + t*t) = r*r / (1 + r*r)
    rr = r * r
    if rr < 1e-20:  # far out: the tail's leading term, in logarithms (x may underflow)
        a = df / 2
        return _exp(2 * a * math.log(r) - math.log(a) - _log_beta(a, 0.5)) / 2
    if rr > 1:
        x, y = 1 / (1 + 1 / rr), (1 / rr) / (1 + 1 / rr)
    else:
        x, y = rr / (1 + rr), 1 / (1 + rr)
    return _ibeta(x, df / 2, 0.5, y) / 2


def _t_center(t: float, df: float) -> float:
    """P(|T| < t) for t >= 0, directly, where it is the smaller of it and the two tails."""
    return _ibeta(t * t / (df + t * t), 0.5, df / 2, df / (df + t * t))


def _t_quantile(two_tails: float, df: float) -> float:
    """The t >= 0 whose two tails hold ``two_tails``, inverting the smaller of tails and center.

    Near 1 the center, 1 - ``two_tails`` (exact there), carries the digits the tails lose.
    """
    if two_tails > 0.5:
        return _invert(lambda u: _t_center(u, df), 1 - two_tails, 0.0, 1.0)
    return _invert(lambda u: -2 * _t_tail(u, df), -two_tails, 0.0, 1.0)


@function(
    "_xlfn.T.DIST", "number", "number", "logical", category=S, rel=1e-12, rel_reason=_REL_TAILS
)
@_finite
def T_DIST(ctx: Any, x: float, df: float, cumulative: bool) -> Any:
    """Student's t distribution's left tail, or its density."""
    k = _df(df)
    if k is None:
        return NUM if cumulative or df >= 1 else DIV0
    if cumulative:
        tail = _t_tail(abs(x), k)
        return 1 - tail if x >= 0 else tail
    log_density = (
        _lgamma_diff(k / 2, 0.5) - 0.5 * math.log(k * math.pi)
        - ((k + 1) / 2) * math.log1p(x * x / k)
    )  # fmt: skip
    return _exp(log_density)


@function("_xlfn.T.DIST.RT", "number", "number", category=S, rel=1e-13, rel_reason=_REL_TAILS)
@_finite
def T_DIST_RT(ctx: Any, x: float, df: float) -> Any:
    """Student's t distribution's right tail."""
    k = _df(df)
    if k is None:
        return NUM
    tail = _t_tail(abs(x), k)
    return tail if x >= 0 else 1 - tail


@function("_xlfn.T.DIST.2T", "number", "number", category=S, rel=1e-13, rel_reason=_REL_TAILS)
@_finite
def T_DIST_2T(ctx: Any, x: float, df: float) -> Any:
    """Student's t distribution's two tails."""
    k = _df(df)
    if k is None or x < 0:
        return NUM
    return 2 * _t_tail(x, k)


@function("TDIST", "number", "number", "number", category=S, rel=1e-14, rel_reason=_REL_TAILS)
@_finite
def TDIST(ctx: Any, x: float, df: float, tails: float) -> Any:
    """Student's t distribution's right tail or two tails."""
    k, n = _df(df), math.trunc(tails)
    if k is None or x < 0 or n not in (1, 2):
        return NUM
    return n * _t_tail(x, k)


@function("_xlfn.T.INV", "number", "number", category=S, rel=1e-12, rel_reason=_REL_SEARCH)
@_finite
def T_INV(ctx: Any, p: float, df: float) -> Any:
    """The inverse of Student's t distribution's left tail."""
    k = _df(df)
    if p == 0 and k is not None:  # #DIV/0! at 1 degree of freedom, #NUM! above
        return DIV0 if k == 1 else NUM
    if k is None or not 0 < p < 1:
        return NUM
    if p == 0.5:
        return 0.0
    t = _t_quantile(2 * min(p, 1 - p), k)
    return t if p > 0.5 else -t


@function("TINV", "number", "number", category=S, rel=1e-13, rel_reason=_REL_SEARCH)
@function("_xlfn.T.INV.2T", "number", "number", category=S, rel=1e-13, rel_reason=_REL_SEARCH)
@_finite
def T_INV_2T(ctx: Any, p: float, df: float) -> Any:
    """The inverse of Student's t distribution's two tails."""
    k = _df(df)
    if p == 0 and k is not None:  # #DIV/0! at 1 degree of freedom, #NUM! above
        return DIV0 if k == 1 else NUM
    if k is None or not 0 < p < 2:
        return NUM
    if p > 1:  # the other side: T.INV.2T(1.5,1) is -1
        return -T_INV_2T(ctx, 2 - p, df)
    if p == 1:
        return 0.0
    return _t_quantile(p, k)


# Tests --------------------------------------------------------------------------------------


def _sample(arg: Any, ctx: Any) -> list[float] | ErrorValue:
    out = []
    for v in numbers((arg,), ctx):
        if isinstance(v, ErrorValue):
            return v
        out.append(v)
    return out


def _mean_var(xs: list[float]) -> tuple[float, float]:
    mean = sum(xs) / len(xs)
    return mean, sum((v - mean) ** 2 for v in xs) / (len(xs) - 1)


@function("TTEST", "raw", "raw", "number", "number", category=S, rel=1e-14, rel_reason=_REL_TEST)
@function(
    "_xlfn.T.TEST", "raw", "raw", "number", "number", category=S, rel=1e-14, rel_reason=_REL_TEST
)
@_finite
def T_TEST(ctx: Any, array1: Any, array2: Any, tails: float, kind: float) -> Any:
    """Student's t-test: paired (1), equal variances (2) or unequal (3)."""
    t_, k_ = math.trunc(tails), math.trunc(kind)
    if t_ not in (1, 2) or k_ not in (1, 2, 3):
        return NUM
    a, b = _sample(array1, ctx), _sample(array2, ctx)
    for s in (a, b):
        if isinstance(s, ErrorValue):
            return s
    assert isinstance(a, list)
    assert isinstance(b, list)
    if k_ == 1:
        d = _paired(array1, array2)
        if d is None:  # neither a range nor an array: the numbers as given
            if len(a) != len(b):
                return NA
            d = [x - y for x, y in zip(a, b, strict=True)]
        elif isinstance(d, ErrorValue):
            return d
        if len(d) < 2:
            return DIV0
        mean, var = _mean_var(d)
        if var == 0:
            return DIV0
        t, df = mean / math.sqrt(var / len(d)), len(d) - 1.0
    else:
        if len(a) < 2 or len(b) < 2:
            return DIV0
        (ma, va), (mb, vb) = _mean_var(a), _mean_var(b)
        na, nb = len(a), len(b)
        if k_ == 2:
            pooled = ((na - 1) * va + (nb - 1) * vb) / (na + nb - 2)
            se, df = math.sqrt(pooled * (1 / na + 1 / nb)), float(na + nb - 2)
        else:
            qa, qb = va / na, vb / nb
            se = math.sqrt(qa + qb)
            df = (qa + qb) ** 2 / (qa * qa / (na - 1) + qb * qb / (nb - 1)) if qa + qb else 0.0
        if se == 0:
            return DIV0
        t = (ma - mb) / se
    return t_ * _t_tail(abs(t), df)


@function("FTEST", "raw", "raw", category=S, rel=1e-14, rel_reason=_REL_TEST)
@function("_xlfn.F.TEST", "raw", "raw", category=S, rel=1e-14, rel_reason=_REL_TEST)
@_finite
def F_TEST(ctx: Any, array1: Any, array2: Any) -> Any:
    """The two-tailed probability that two samples' variances do not differ."""
    a, b = _sample(array1, ctx), _sample(array2, ctx)
    for s in (a, b):
        if isinstance(s, ErrorValue):
            return s
    assert isinstance(a, list)
    assert isinstance(b, list)
    if len(a) < 2 or len(b) < 2:
        return DIV0
    va, vb = _mean_var(a)[1], _mean_var(b)[1]
    if va == 0 or vb == 0:
        return DIV0
    left = _f_cdf(va / vb, len(a) - 1, len(b) - 1)
    return 2 * min(left, 1 - left)


def _paired(x: Any, y: Any) -> list[float] | ErrorValue | None:
    """Differences of the pairs at the same position, in row-major order, as Excel pairs them.

    The two must hold as many elements, whatever their shapes (#N/A otherwise); a pair with
    anything but a number on either side is dropped (capture continuous_family). None where
    either is neither a range nor an array.
    """
    tx, ty = _table(x), _table(y)
    if tx is None or ty is None:
        return None
    if tx[0] * tx[1] != ty[0] * ty[1]:
        return NA
    cx, cy = tx[1], ty[1]
    both = {i * cx + j for i, j in tx[2]} & {i * cy + j for i, j in ty[2]}
    out = []
    for k in sorted(both):
        u, v = tx[3](*divmod(k, cx)), ty[3](*divmod(k, cy))
        if type(u) is float and type(v) is float:
            out.append(u - v)
    return out


def _table(arg: Any) -> tuple[int, int, set[tuple[int, int]], Any] | None:
    """A range or an array as rows, columns, the positions holding something, and a reader."""
    if isinstance(arg, Ref):
        return arg.rows, arg.columns, {(i, j) for i, j, _ in arg.cells()}, arg.value
    if is_array(arg):
        rows, cols = shape(arg)
        held = {(i, j) for i in range(rows) for j in range(cols)}
        return rows, cols, held, lambda i, j: arg[i][j]
    return None


@function("CHITEST", "raw", "raw", category=S, rel=1e-15, rel_reason=_REL_TEST)
@function("_xlfn.CHISQ.TEST", "raw", "raw", category=S, rel=1e-15, rel_reason=_REL_TEST)
@_finite
def CHISQ_TEST(ctx: Any, actual: Any, expected: Any) -> Any:
    """The chi-squared test of observed counts against expected ones, ranges or arrays."""
    a, e_ = _table(actual), _table(expected)
    if a is None or e_ is None or a[0] * a[1] != e_[0] * e_[1]:
        return NA
    chi = 0.0
    # Element by element in row-major order, whatever the two shapes, as Excel pairs them; only
    # where either holds something, so a whole sheet is no cost (capture continuous_family).
    ca, ce = a[1], e_[1]
    held = {i * ca + j for i, j in a[2]} | {i * ce + j for i, j in e_[2]}
    for k in sorted(held):
        o, e = a[3](*divmod(k, ca)), e_[3](*divmod(k, ce))
        if isinstance(o, ErrorValue):
            return o
        if isinstance(e, ErrorValue):
            return e
        if isinstance(o, float) and isinstance(e, float):
            if e == 0:
                return DIV0
            chi += (o - e) ** 2 / e
    r, c = a[0], a[1]  # the degrees of freedom are actual's
    df = (r - 1) * (c - 1) if r > 1 and c > 1 else max(r, c) - 1
    if df < 1:
        return NA
    return _igamma_upper(df / 2, chi / 2)


@function(
    "ZTEST", "raw", "number", "number", min_args=2, category=S, rel=1e-12, rel_reason=_REL_TEST
)
@function(
    "_xlfn.Z.TEST",
    "raw",
    "number",
    "number",
    min_args=2,
    category=S,
    rel=1e-12,
    rel_reason=_REL_TEST,
)
@_finite
def Z_TEST(ctx: Any, array: Any, x: float, sigma: float | None = None) -> Any:
    """The one-tailed probability of a z-test of a sample's mean against ``x``."""
    if array is MISSING:  # an empty sample, #DIV/0! in Excel (capture empty_args)
        return DIV0
    a = _sample(array, ctx)
    if isinstance(a, ErrorValue):
        return a
    if not a:
        return NA
    mean = sum(a) / len(a)
    if sigma is None:
        if len(a) < 2:
            return DIV0
        sigma = math.sqrt(_mean_var(a)[1])
        if sigma == 0:
            return DIV0
    elif sigma <= 0:
        return NUM
    return _phi(-(mean - x) / (sigma / math.sqrt(len(a))))


@function(
    "_xlfn.CONFIDENCE.T",
    "number",
    "number",
    "number",
    category=S,
    rel=1e-14,
    rel_reason=_REL_SEARCH,
)
@_finite
def CONFIDENCE_T(ctx: Any, alpha: float, sd: float, size: float) -> Any:
    """The half-width of a confidence interval for a mean, from Student's t distribution."""
    n = math.trunc(size)
    if not 0 < alpha < 1 or sd <= 0 or n < 1:
        return NUM
    if n == 1:
        return DIV0
    t = T_INV_2T(ctx, alpha, n - 1)  # #NUM! past 1E10 degrees of freedom
    return t if isinstance(t, ErrorValue) else t * sd / math.sqrt(n)
