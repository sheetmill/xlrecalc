"""The beta, gamma, chi-squared, F and t distributions, their inverses and tests.

Captured from Excel (``continuous_family``). The cumulative functions rest on the regularized
incomplete beta and gamma functions, each tail computed directly where it is the smaller; an
inverse is found by bisection on the smaller tail. Not imported by the package yet: it is
registered once each function's tolerance is declared from ``tools/relative_spread.py``.
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._coerce import numbers
from xlrecalc._functions._registry import function
from xlrecalc._functions._statistical import _phi
from xlrecalc._values import ErrorValue, Ref

S = "statistical"
NUM, DIV0, NA = ErrorValue.NUM, ErrorValue.DIV0, ErrorValue.NA
_MAX_DF = 1e10


# The incomplete beta and gamma functions ----------------------------------------------------


def _beta_fraction(x: float, a: float, b: float) -> float:
    """The continued fraction of the incomplete beta function (modified Lentz)."""
    tiny = 1e-300
    c, d = 1.0, 1.0 - (a + b) * x / (a + 1)
    d = 1 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, 10_000):
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
            break
    return h


def _lgamma_diff(a: float, b: float) -> float:
    """Lgamma(a + b) - lgamma(a), by Stirling's expansion where ``a`` is large (they cancel)."""
    if a > 1e5 and abs(b) < 10:
        c = b * (b - 1)
        return b * math.log(a) + c / (2 * a) - c * (2 * b - 1) / (12 * a * a) + c * c / (12 * a**3)
    return math.lgamma(a + b) - math.lgamma(a)


def _stirling_rest(x: float) -> float:
    """lgamma(x) less Stirling's approximation, for x of 10 and more."""
    xx = x * x
    return (1 / 12 - (1 / 360 - (1 / 1260 - 1 / (1680 * xx)) / xx) / xx) / x


def _log_beta(a: float, b: float) -> float:
    """Log B(a, b) without cancelling: by Stirling's form where both are large."""
    big, small = (a, b) if a >= b else (b, a)
    if small >= 10:
        total = a + b
        return (
            0.5 * math.log(2 * math.pi) - 0.5 * math.log(total)
            + (a - 0.5) * math.log1p(-b / total) + (b - 0.5) * math.log1p(-a / total)
            + _stirling_rest(a) + _stirling_rest(b) - _stirling_rest(total)
        )  # fmt: skip
    return math.lgamma(small) - _lgamma_diff(big, small)


def _ibeta(x: float, a: float, b: float, y: float | None = None) -> float:
    """The regularized incomplete beta function I_x(a, b); ``y`` is 1 - x, given exactly."""
    y = 1 - x if y is None else y
    if x <= 0:
        return 0.0
    if y <= 0:
        return 1.0
    log_x = math.log1p(-y) if y < 0.5 else math.log(x)
    log_y = math.log1p(-x) if x < 0.5 else math.log(y)
    log_front = a * log_x + b * log_y - _log_beta(a, b)
    if x < (a + 1) / (a + b + 2):
        return math.exp(log_front) * _beta_fraction(x, a, b) / a
    return 1 - math.exp(log_front) * _beta_fraction(y, b, a) / b


def _igamma(a: float, x: float) -> float:
    """The regularized lower incomplete gamma function P(a, x)."""
    return _gamma_tails(a, x)[0]


def _igamma_upper(a: float, x: float) -> float:
    """The regularized upper incomplete gamma function Q(a, x), not as 1 - P in its tail."""
    return _gamma_tails(a, x)[1]


def _gamma_tails(a: float, x: float) -> tuple[float, float]:
    """P(a, x) and Q(a, x), each computed directly where it is the smaller."""
    if x <= 0:
        return 0.0, 1.0
    log_front = a * math.log(x) - x - math.lgamma(a)
    if x < a + 1:
        term = total = 1 / a
        n = a
        for _ in range(100_000):
            n += 1
            term *= x / n
            total += term
            if abs(term) < abs(total) * 1e-17:
                break
        lower = min(1.0, total * math.exp(log_front))
        return lower, 1 - lower
    tiny = 1e-300
    b = x + 1 - a
    c, d = 1 / tiny, 1 / b
    h = d
    for i in range(1, 100_000):
        an = -i * (i - a)
        b += 2
        d = an * d + b
        d = 1 / (d if abs(d) > tiny else tiny)
        c = b + an / c
        c = c if abs(c) > tiny else tiny
        h *= d * c
        if abs(d * c - 1) < 1e-16:
            break
    upper = math.exp(log_front) * h
    return max(0.0, 1 - upper), upper


def _invert(cdf: Any, p: float, lo: float, hi: float) -> float:
    """The x with ``cdf(x) == p`` by bisection, to the last bit, ``hi`` doubled until past."""
    while cdf(hi) < p and hi < 1e300:
        lo, hi = hi, hi * 2
    for _ in range(2000):
        mid = (lo + hi) / 2
        if mid in (lo, hi):
            break
        if cdf(mid) < p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


# Beta ---------------------------------------------------------------------------------------


def _beta_args(a: float, b: float, lo: float, hi: float) -> bool:
    return a > 0 and b > 0 and lo < hi


@function("_xlfn.BETA.DIST", "number", "number", "number", "logical", "number", "number",
          min_args=4, category=S)  # fmt: skip
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
    if z == 1 and b < 1:
        return NUM
    log_b = math.lgamma(a) + math.lgamma(b) - math.lgamma(a + b)
    if z in (0.0, 1.0):
        return (1.0 if (a == 1 if z == 0 else b == 1) else 0.0) / math.exp(log_b) / (hi - lo)
    density = math.exp((a - 1) * math.log(z) + (b - 1) * math.log1p(-z) - log_b)
    return density / (hi - lo)


@function("BETADIST", "number", "number", "number", "number", "number", min_args=3, category=S)
def BETADIST(ctx: Any, x: float, a: float, b: float, lo: float = 0.0, hi: float = 1.0) -> Any:
    """The cumulative beta distribution."""
    return BETA_DIST(ctx, x, a, b, True, lo, hi)


@function("BETAINV", "number", "number", "number", "number", "number", min_args=3, category=S)
@function("_xlfn.BETA.INV", "number", "number", "number", "number", "number", min_args=3,
          category=S)  # fmt: skip
def BETA_INV(ctx: Any, p: float, a: float, b: float, lo: float = 0.0, hi: float = 1.0) -> Any:
    """The inverse of the cumulative beta distribution."""
    if not _beta_args(a, b, lo, hi) or not 0 < p < 1:
        return NUM
    if p > 0.5:  # the upper tail I_{1-z}(b, a), where it is the smaller
        return hi - (hi - lo) * _invert(lambda y: _ibeta(y, b, a), 1 - p, 0.0, 1.0)
    return lo + (hi - lo) * _invert(lambda z: _ibeta(z, a, b), p, 0.0, 1.0)


# Gamma and chi-squared ----------------------------------------------------------------------


@function("GAMMADIST", "number", "number", "number", "logical", category=S)
@function("_xlfn.GAMMA.DIST", "number", "number", "number", "logical", category=S)
def GAMMA_DIST(ctx: Any, x: float, alpha: float, beta: float, cumulative: bool) -> Any:
    """The gamma distribution with shape ``alpha`` and scale ``beta``."""
    if x < 0 or alpha <= 0 or beta <= 0:
        return NUM
    if cumulative:
        return _igamma(alpha, x / beta)
    if x == 0:  # Excel: 0 for a shape above 1, #NUM! at 1 and below
        return 0.0 if alpha > 1 else NUM
    return math.exp(
        (alpha - 1) * math.log(x) - x / beta - math.lgamma(alpha) - alpha * math.log(beta)
    )


@function("GAMMAINV", "number", "number", "number", category=S)
@function("_xlfn.GAMMA.INV", "number", "number", "number", category=S)
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


@function("_xlfn.CHISQ.DIST", "number", "number", "logical", category=S)
def CHISQ_DIST(ctx: Any, x: float, df: float, cumulative: bool) -> Any:
    """The chi-squared distribution."""
    k = _df(df)
    if k is None or x < 0:
        return NUM
    if x == 0 and not cumulative:  # 0.5 at 2 degrees of freedom, as Excel gives it
        return 0.5 if k == 2 else 0.0 if k > 2 else NUM
    return GAMMA_DIST(ctx, x, k / 2, 2.0, cumulative)


@function("CHIDIST", "number", "number", category=S)
@function("_xlfn.CHISQ.DIST.RT", "number", "number", category=S)
def CHISQ_DIST_RT(ctx: Any, x: float, df: float) -> Any:
    """The right tail of the chi-squared distribution."""
    k = _df(df)
    if k is None or x < 0:
        return NUM
    return _igamma_upper(k / 2, x / 2)


@function("_xlfn.CHISQ.INV", "number", "number", category=S)
def CHISQ_INV(ctx: Any, p: float, df: float) -> Any:
    """The inverse of the left tail of the chi-squared distribution."""
    k = _df(df)
    if k is None or not 0 <= p < 1:
        return NUM
    return GAMMA_INV(ctx, p, k / 2, 2.0)


@function("CHIINV", "number", "number", category=S)
@function("_xlfn.CHISQ.INV.RT", "number", "number", category=S)
def CHISQ_INV_RT(ctx: Any, p: float, df: float) -> Any:
    """The inverse of the right tail of the chi-squared distribution."""
    k = _df(df)
    if k is None or not 0 < p <= 1:
        return NUM
    if p == 1:
        return 0.0
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


@function("_xlfn.F.DIST", "number", "number", "number", "logical", category=S)
def F_DIST(ctx: Any, x: float, d1: float, d2: float, cumulative: bool) -> Any:
    """The F distribution."""
    m, n = _df(d1), _df(d2)
    if m is None or n is None or x < 0:
        return NUM
    if cumulative:
        return _f_cdf(x, m, n)
    if x == 0:
        return 0.0 if m > 2 else (1.0 if m == 2 else NUM)
    log_b = math.lgamma(m / 2) + math.lgamma(n / 2) - math.lgamma((m + n) / 2)
    log_density = (
        (m / 2) * math.log(m / n) + (m / 2 - 1) * math.log(x)
        - ((m + n) / 2) * math.log1p(m * x / n) - log_b
    )  # fmt: skip
    return math.exp(log_density)


@function("FDIST", "number", "number", "number", category=S)
@function("_xlfn.F.DIST.RT", "number", "number", "number", category=S)
def F_DIST_RT(ctx: Any, x: float, d1: float, d2: float) -> Any:
    """The right tail of the F distribution."""
    m, n = _df(d1), _df(d2)
    if m is None or n is None or x < 0:
        return NUM
    return _f_right(x, m, n)


@function("_xlfn.F.INV", "number", "number", "number", category=S)
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


@function("FINV", "number", "number", "number", category=S)
@function("_xlfn.F.INV.RT", "number", "number", "number", category=S)
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
    if df > 1e5:  # the normal and its 1/df and 1/df^2 terms, where the fraction is too slow
        density = math.exp(-t * t / 2) / math.sqrt(2 * math.pi)
        first = (t**3 + t) / (4 * df)
        second = (5 * t**7 + 16 * t**5 + 3 * t**3 - 3 * t) / (96 * df * df)
        return _phi(-t) + density * (first + second)
    r = math.sqrt(df) / t  # x = df / (df + t*t) = r*r / (1 + r*r)
    rr = r * r
    if rr < 1e-20:  # far out: the tail's leading term, in logarithms (x may underflow)
        a = df / 2
        return math.exp(2 * a * math.log(r) - math.log(a) - _log_beta(a, 0.5)) / 2
    if rr > 1:
        x, y = 1 / (1 + 1 / rr), (1 / rr) / (1 + 1 / rr)
    else:
        x, y = rr / (1 + rr), 1 / (1 + rr)
    return _ibeta(x, df / 2, 0.5, y) / 2


@function("_xlfn.T.DIST", "number", "number", "logical", category=S)
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
    return math.exp(log_density)


@function("_xlfn.T.DIST.RT", "number", "number", category=S)
def T_DIST_RT(ctx: Any, x: float, df: float) -> Any:
    """Student's t distribution's right tail."""
    k = _df(df)
    if k is None:
        return NUM
    tail = _t_tail(abs(x), k)
    return tail if x >= 0 else 1 - tail


@function("_xlfn.T.DIST.2T", "number", "number", category=S)
def T_DIST_2T(ctx: Any, x: float, df: float) -> Any:
    """Student's t distribution's two tails."""
    k = _df(df)
    if k is None or x < 0:
        return NUM
    return 2 * _t_tail(x, k)


@function("TDIST", "number", "number", "number", category=S)
def TDIST(ctx: Any, x: float, df: float, tails: float) -> Any:
    """Student's t distribution's right tail or two tails."""
    k, n = _df(df), math.trunc(tails)
    if k is None or x < 0 or n not in (1, 2):
        return NUM
    return n * _t_tail(x, k)


@function("_xlfn.T.INV", "number", "number", category=S)
def T_INV(ctx: Any, p: float, df: float) -> Any:
    """The inverse of Student's t distribution's left tail."""
    k = _df(df)
    if p == 0 and k is not None:  # #DIV/0! at 1 degree of freedom, #NUM! above
        return DIV0 if k == 1 else NUM
    if k is None or not 0 < p < 1:
        return NUM
    if p == 0.5:
        return 0.0
    tail = min(p, 1 - p)  # the right tail beyond |t|, inverted as it is, not as 1 - p
    t = _invert(lambda u: -_t_tail(u, k), -tail, 0.0, 1.0)
    return t if p > 0.5 else -t


@function("TINV", "number", "number", category=S)
@function("_xlfn.T.INV.2T", "number", "number", category=S)
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
    return _invert(lambda u: -2 * _t_tail(u, k), -p, 0.0, 1.0)


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


@function("TTEST", "raw", "raw", "number", "number", category=S)
@function("_xlfn.T.TEST", "raw", "raw", "number", "number", category=S)
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
        if len(a) != len(b):
            return NA
        d = [x - y for x, y in zip(a, b, strict=True)]
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


@function("FTEST", "raw", "raw", category=S)
@function("_xlfn.F.TEST", "raw", "raw", category=S)
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


@function("CHITEST", "raw", "raw", category=S)
@function("_xlfn.CHISQ.TEST", "raw", "raw", category=S)
def CHISQ_TEST(ctx: Any, actual: Any, expected: Any) -> Any:
    """The chi-squared test of observed counts against expected ones."""
    if not (isinstance(actual, Ref) and isinstance(expected, Ref)):
        return NA
    if (actual.rows, actual.columns) != (expected.rows, expected.columns):
        return NA
    chi = 0.0
    for i in range(actual.rows):
        for j in range(actual.columns):
            o, e = actual.value(i, j), expected.value(i, j)
            if isinstance(o, ErrorValue):
                return o
            if isinstance(e, ErrorValue):
                return e
            if isinstance(o, float) and isinstance(e, float):
                if e == 0:
                    return DIV0
                chi += (o - e) ** 2 / e
    r, c = actual.rows, actual.columns
    df = (r - 1) * (c - 1) if r > 1 and c > 1 else max(r, c) - 1
    if df < 1:
        return NA
    return _igamma_upper(df / 2, chi / 2)


@function("ZTEST", "raw", "number", "number", min_args=2, category=S)
@function("_xlfn.Z.TEST", "raw", "number", "number", min_args=2, category=S)
def Z_TEST(ctx: Any, array: Any, x: float, sigma: float | None = None) -> Any:
    """The one-tailed probability of a z-test of a sample's mean against ``x``."""
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


@function("_xlfn.CONFIDENCE.T", "number", "number", "number", category=S)
def CONFIDENCE_T(ctx: Any, alpha: float, sd: float, size: float) -> Any:
    """The half-width of a confidence interval for a mean, from Student's t distribution."""
    n = math.trunc(size)
    if not 0 < alpha < 1 or sd <= 0 or n < 1:
        return NUM
    if n == 1:
        return DIV0
    return T_INV_2T(ctx, alpha, n - 1) * sd / math.sqrt(n)
