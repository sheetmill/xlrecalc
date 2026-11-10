"""Discrete and closed-form distributions, ``GAMMA`` and its logarithm, ``ERF`` and ``ERFC``.

Captured from Excel (``distributions_family``). Counts are truncated toward zero; an argument
outside its domain is ``#NUM!``.
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._arith import flush
from xlrecalc._arrays import NotImplementedFormError
from xlrecalc._coerce import scalar, to_number
from xlrecalc._functions._registry import function
from xlrecalc._functions._statistical import _phi, _ppnd
from xlrecalc._values import ErrorValue

S = "statistical"
E = "engineering"
NUM, VALUE = ErrorValue.NUM, ErrorValue.VALUE
_DISCRETE = (
    "Excel's terms of the discrete distributions match neither the exact count of ways times the "
    "powers, nor lgamma, nor Loader's saddle point in the last bits; within 8 units in the last "
    "place but listed rows, most at a thousand trials and more"
)
_LIBRARY = (
    "Excel's own exp, pow, erf and gamma differ from the platform's in the last bits; within 4 "
    "units in the last place but listed rows"
)


_MAX_TERMS = 1_000_000  # terms of a sum computed one by one; past it, not computed here


def _log_choose(n: float, k: float) -> float:
    return math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1)


def _ways(n: int, k: int) -> float | None:
    """The number of ways to choose ``k`` of ``n``, exactly, or ``None`` past a double.

    Bounded first in logarithms, so that ``math.comb`` never builds an integer of millions of
    digits for a count such as 1E15 (a call that hung an unguarded comparison of ``COMBIN``).
    """
    if _log_choose(n, k) > 700:
        return None
    return float(math.comb(n, k))


def _terms(count: int, name: str) -> None:
    """Refuse a sum of more terms than ``_MAX_TERMS``: kept and reported, never a wrong value."""
    if count > _MAX_TERMS:
        raise NotImplementedFormError(f"{name} summed over more than a million terms")


def _binom_pmf(k: float, n: float, p: float) -> float:
    """The binomial density: the exact count of ways times the powers, closest to Excel."""
    if p == 0:
        return 1.0 if k == 0 else 0.0
    if p == 1:
        return 1.0 if k == n else 0.0
    ways = _ways(int(n), int(k))
    if ways is None:  # past a double: in logarithms
        return flush(math.exp(_log_choose(n, k) + k * math.log(p) + (n - k) * math.log1p(-p)))
    return flush(ways * p**k * (1 - p) ** (n - k))


def _binom_cdf(k: float, n: float, p: float) -> float:
    """The cumulative binomial probability, at most 1 (Excel's 1.0 where a sum passes it)."""
    if k >= n:
        return 1.0
    _terms(int(k) + 1, "BINOM.DIST")
    return min(1.0, sum(_binom_pmf(float(i), n, p) for i in range(int(k) + 1)))


# Binomial -----------------------------------------------------------------------------------


@function(
    "BINOMDIST", "number", "number", "number", "logical", category=S, ulps=8, ulps_reason=_DISCRETE
)
@function(
    "_xlfn.BINOM.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    ulps=8,
    ulps_reason=_DISCRETE,
)
def BINOM_DIST(ctx: Any, number_s: float, trials: float, p: float, cumulative: bool) -> Any:
    """The binomial distribution: the probability of ``number_s`` successes in ``trials``."""
    k, n = math.trunc(number_s), math.trunc(trials)
    if k < 0 or k > n or not 0 <= p <= 1:
        return NUM
    if p == 1 and k == 0 and n > 0 and not cumulative:  # as Excel gives it
        return NUM
    return _binom_cdf(k, n, p) if cumulative else _binom_pmf(k, n, p)


@function(
    "_xlfn.BINOM.DIST.RANGE",
    "number",
    "number",
    "number",
    "number",
    min_args=3,
    category=S,
    ulps=8,
    ulps_reason=_DISCRETE,
)
def BINOM_DIST_RANGE(
    ctx: Any, trials: float, p: float, number_s: float, number_s2: float | None = None
) -> Any:
    """The probability of between ``number_s`` and ``number_s2`` successes in ``trials``."""
    n, lo = math.trunc(trials), math.trunc(number_s)
    hi = lo if number_s2 is None else math.trunc(number_s2)
    if n < 0 or not 0 <= p <= 1 or lo < 0 or lo > n or hi < lo or hi > n:
        return NUM
    _terms(hi - lo + 1, "BINOM.DIST.RANGE")
    return sum(_binom_pmf(float(i), n, p) for i in range(lo, hi + 1))


@function("CRITBINOM", "number", "number", "number", category=S)
@function("_xlfn.BINOM.INV", "number", "number", "number", category=S)
def BINOM_INV(ctx: Any, trials: float, p: float, alpha: float) -> Any:
    """The smallest number of successes whose cumulative probability reaches ``alpha``."""
    n = math.trunc(trials)
    if n < 0 or not 0 < p < 1 or not 0 < alpha < 1:
        return NUM
    _terms(n + 1, "BINOM.INV")
    total = 0.0
    for k in range(n + 1):
        total += _binom_pmf(float(k), n, p)
        if total >= alpha:
            return float(k)
    return float(n)


@function("NEGBINOMDIST", "number", "number", "number", category=S, ulps=8, ulps_reason=_DISCRETE)
def NEGBINOMDIST(ctx: Any, number_f: float, number_s: float, p: float) -> Any:
    """The probability of ``number_f`` failures before the ``number_s``-th success."""
    return NEGBINOM_DIST(ctx, number_f, number_s, p, False)


@function(
    "_xlfn.NEGBINOM.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    ulps=8,
    ulps_reason=_DISCRETE,
)
def NEGBINOM_DIST(ctx: Any, number_f: float, number_s: float, p: float, cumulative: bool) -> Any:
    """The negative binomial distribution."""
    f, s = math.trunc(number_f), math.trunc(number_s)
    if f < 0 or s < 1 or not 0 < p < 1:
        return NUM

    def pmf(i: int) -> float:
        ways = _ways(i + s - 1, s - 1)
        if ways is None:
            return flush(
                math.exp(_log_choose(i + s - 1, s - 1) + s * math.log(p) + i * math.log1p(-p))
            )
        return flush(ways * p**s * (1 - p) ** i)

    if not cumulative:
        return pmf(f)
    _terms(f + 1, "NEGBINOM.DIST")
    return sum(pmf(i) for i in range(f + 1))


# Poisson and hypergeometric -------------------------------------------------------------------


@function("POISSON", "number", "number", "logical", category=S, ulps=8, ulps_reason=_DISCRETE)
@function(
    "_xlfn.POISSON.DIST", "number", "number", "logical", category=S, ulps=8, ulps_reason=_DISCRETE
)
def POISSON_DIST(ctx: Any, x: float, mean: float, cumulative: bool) -> Any:
    """The Poisson distribution."""
    k = math.trunc(x)
    if k < 0 or (mean < 0 and k > 0):
        return NUM
    if k == 0:  # exp(-mean), for a negative mean too
        return math.exp(-mean)

    def pmf(i: int) -> float:
        if mean == 0:
            return 1.0 if i == 0 else 0.0
        if i <= 170:  # i! within a double
            try:
                return flush(mean**i / math.factorial(i) * math.exp(-mean))
            except OverflowError:
                pass
        return flush(math.exp(i * math.log(mean) - mean - math.lgamma(i + 1)))

    if not cumulative:
        return pmf(k)
    _terms(k + 1, "POISSON.DIST")
    return sum(pmf(i) for i in range(k + 1))


@function(
    "HYPGEOMDIST", "number", "number", "number", "number", category=S, ulps=8, ulps_reason=_DISCRETE
)
def HYPGEOMDIST(ctx: Any, s: float, n: float, big_s: float, big_n: float) -> Any:
    """The hypergeometric distribution's probability of ``s`` successes in a sample of ``n``."""
    return HYPGEOM_DIST(ctx, s, n, big_s, big_n, False)


@function(
    "_xlfn.HYPGEOM.DIST",
    "number",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    ulps=8,
    ulps_reason=_DISCRETE,
)
def HYPGEOM_DIST(
    ctx: Any, sample_s: float, number_sample: float, population_s: float, number_pop: float,
    cumulative: bool,
) -> Any:  # fmt: skip
    """The hypergeometric distribution."""
    k, n = math.trunc(sample_s), math.trunc(number_sample)
    big_k, big_n = math.trunc(population_s), math.trunc(number_pop)
    if big_n <= 0 or n < 0 or n > big_n or big_k < 0 or big_k > big_n or k < 0:
        return NUM
    whole = _ways(big_n, n)
    low, high = max(0, n - big_n + big_k), min(n, big_k)  # the counts that can occur

    def pmf(i: int) -> float:
        if not low <= i <= high:
            return 0.0
        if whole is None:  # past a double: in logarithms
            logs = _log_choose(big_k, i) + _log_choose(big_n - big_k, n - i)
            return flush(math.exp(logs - _log_choose(big_n, n)))
        return math.comb(big_k, i) * math.comb(big_n - big_k, n - i) / whole

    if cumulative:
        _terms(min(k, high) - low + 1, "HYPGEOM.DIST")
        return sum((pmf(i) for i in range(low, min(k, high) + 1)), 0.0)
    return pmf(k)


# Continuous -----------------------------------------------------------------------------------


@function("EXPONDIST", "number", "number", "logical", category=S)
@function("_xlfn.EXPON.DIST", "number", "number", "logical", category=S)
def EXPON_DIST(ctx: Any, x: float, rate: float, cumulative: bool) -> Any:
    """The exponential distribution."""
    if x < 0 or rate <= 0:
        return NUM
    return -math.expm1(-rate * x) if cumulative else rate * math.exp(-rate * x)


@function(
    "WEIBULL", "number", "number", "number", "logical", category=S, ulps=4, ulps_reason=_LIBRARY
)
@function(
    "_xlfn.WEIBULL.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    ulps=4,
    ulps_reason=_LIBRARY,
)
def WEIBULL_DIST(ctx: Any, x: float, alpha: float, beta: float, cumulative: bool) -> Any:
    """The Weibull distribution."""
    if x < 0 or alpha <= 0 or beta <= 0:
        return NUM
    t = (x / beta) ** alpha
    if cumulative:
        return -math.expm1(-t)
    if x == 0:
        return 0.0
    return alpha / beta**alpha * x ** (alpha - 1) * math.exp(-t)


@function("LOGNORMDIST", "number", "number", "number", category=S, ulps=4, ulps_reason=_LIBRARY)
def LOGNORMDIST(ctx: Any, x: float, mean: float, sd: float) -> Any:
    """The cumulative lognormal distribution."""
    return LOGNORM_DIST(ctx, x, mean, sd, True)


@function(
    "_xlfn.LOGNORM.DIST",
    "number",
    "number",
    "number",
    "logical",
    category=S,
    ulps=4,
    ulps_reason=_LIBRARY,
)
def LOGNORM_DIST(ctx: Any, x: float, mean: float, sd: float, cumulative: bool) -> Any:
    """The lognormal distribution."""
    if x <= 0 or sd <= 0:
        return NUM
    z = (math.log(x) - mean) / sd
    if cumulative:
        return _phi(z)
    return math.exp(-z * z / 2) / (x * sd * math.sqrt(2 * math.pi))


@function("LOGINV", "number", "number", "number", category=S, ulps=4, ulps_reason=_LIBRARY)
@function(
    "_xlfn.LOGNORM.INV", "number", "number", "number", category=S, ulps=4, ulps_reason=_LIBRARY
)
def LOGNORM_INV(ctx: Any, p: float, mean: float, sd: float) -> Any:
    """The inverse of the cumulative lognormal distribution."""
    if not 0 < p < 1 or sd <= 0:
        return NUM
    return math.exp(mean + sd * _ppnd(p))


# GAMMA, ERF -----------------------------------------------------------------------------------


@function("_xlfn.GAMMA", "number", category=S, ulps=4, ulps_reason=_LIBRARY)
def GAMMA(ctx: Any, x: float) -> Any:
    """The gamma function; ``#NUM!`` at 0, a negative integer, or past a double."""
    if x == 0 or (x < 0 and x == math.floor(x)):
        return NUM
    try:
        value = math.gamma(x)
    except (OverflowError, ValueError):
        return NUM
    return value if math.isfinite(value) else NUM


@function("GAMMALN", "number", category=S, ulps=4, ulps_reason=_LIBRARY)
@function("_xlfn.GAMMALN.PRECISE", "number", category=S, ulps=4, ulps_reason=_LIBRARY)
def GAMMALN(ctx: Any, x: float) -> Any:
    """The natural logarithm of the gamma function, for ``x`` above 0."""
    if x <= 0:
        return NUM
    try:
        value = math.lgamma(x)
    except OverflowError:
        return NUM
    return value if math.isfinite(value) else NUM


def _no_logical(v: Any, ctx: Any) -> Any:
    """A number for ``ERF`` and ``ERFC``, which take no logical, given or in a cell."""
    v = scalar(v, ctx)
    return VALUE if isinstance(v, bool) else to_number(v, ctx)


@function("ERF", "value", "value", min_args=1, category=E, ulps=4, ulps_reason=_LIBRARY)
def ERF(ctx: Any, lower: Any, upper: Any = None) -> Any:
    """The error function from 0 to ``lower``, or between ``lower`` and ``upper``.

    Either argument empty is ``#N/A`` (``_registry._EMPTY_ANY_IS_NA``).
    """
    a = _no_logical(lower, ctx)
    b = None if upper is None else _no_logical(upper, ctx)
    for v in (a, b):
        if isinstance(v, ErrorValue):
            return v
    if b is None:
        return math.erf(a)
    return math.erf(b) - math.erf(a)


@function("_xlfn.ERF.PRECISE", "value", category=E, ulps=4, ulps_reason=_LIBRARY)
def ERF_PRECISE(ctx: Any, x: Any) -> Any:
    """The error function from 0 to ``x``."""
    v = _no_logical(x, ctx)
    return v if isinstance(v, ErrorValue) else math.erf(v)


@function("ERFC", "value", category=E, ulps=4, ulps_reason=_LIBRARY)
@function("_xlfn.ERFC.PRECISE", "value", category=E, ulps=4, ulps_reason=_LIBRARY)
def ERFC(ctx: Any, x: Any) -> Any:
    """The complementary error function."""
    v = _no_logical(x, ctx)
    return v if isinstance(v, ErrorValue) else math.erfc(v)
