"""BESSELI, BESSELJ, BESSELK, BESSELY (capture ``bessel_family``).

Excel computes these with the approximations of Numerical Recipes in C (Press, Teukolsky,
Vetterling and Flannery, 2nd edition, 1992, sections 6.5 and 6.6), themselves from Hart and from
Abramowitz and Stegun: for orders 0 and 1 a rational function of x^2 below a switch (x = 8 for J
and Y, 3.75 for I, 2 for K) and an asymptotic form with polynomials in (8/x)^2 or 1/x above it;
for higher orders a recurrence, upward for Y and K and for J when x > n, downward (Miller's
method, with rescaling) for I and for J when x <= n. This module evaluates the same
approximations, with their published coefficients, in doubles and in the same order of
operations, so that its values are Excel's bit for bit; the code is written for this package.

Excel's results differ from the printed approximations in four places, which this module follows
(QUIRKS.md): the J0 asymptotic coefficient -0.1098628627E-2 is -0.1098628267E-2; the J1
asymptotic polynomial repeats -0.3516396496E-4 (raising its degree by one, and its error to about
3E-5 relative from x = 8); Y0 below x = 8 adds 0.636619772 * (J0(x) * ln x); and Y0 and Y1 from
x = 8 group their second term as z * (cos * q). The approximations are accurate to about 1E-8
relative; Excel's values are reproduced, not the exact functions.

A result that is not finite is ``#NUM!``; one that underflows is 0.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from typing import Any

from xlrecalc._coerce import Context
from xlrecalc._functions._engineering import _strict
from xlrecalc._functions._registry import function
from xlrecalc._values import ErrorValue

NUM = ErrorValue.NUM
E = "engineering"
_TRIG = (
    "the platform's sine and cosine differ in the last bits; J and Y use them from x = 8, and the "
    "recurrences carry the difference (Linux x86_64: 62 of 4,359 captured rows)"
)
_TWO_OVER_PI = 0.636619772
_LARGEST_ORDER = 10_000_000
"""The largest n computed, for all four: 1E7 is computed, 10,000,001 is ``#NUM!``."""
_LARGEST_ANGLE = 2.0**27 - 1
"""The angle (x less 0.785398164 for order 0, 2.356194491 for order 1) from which J and Y are
``#NUM!``: ``BESSELJ(134217727.75,0)`` is computed, ``BESSELJ(134217727.8,0)`` is not, and
``BESSELJ(134217729.3,1)`` is computed, ``BESSELJ(134217729.4,1)`` is not."""
_LARGEST_I = 709.7
"""The largest x for which I is computed: 709.7 is, 709.71 is ``#NUM!``, though ``exp`` would not
overflow before 709.78; the limit is taken to be 709.7 itself."""
_START = 40.0
"""How far above n the downward recurrence starts: from 2 ((n + sqrt(40 n)) // 2) for J, from
2 (n + sqrt(40 n)) for I."""
_RESCALE_ABOVE, _RESCALE_BY = 1e10, 1e-10

Coefficients = tuple[float, ...]
"""A polynomial's coefficients, lowest degree first."""

# Orders 0 and 1 below the switch: numerator and denominator in y = x^2 (J1 and Y1 times x).
_J0_LOW = (
    (57568490574.0, -13362590354.0, 651619640.7, -11214424.18, 77392.33017, -184.9052456),
    (57568490411.0, 1029532985.0, 9494680.718, 59272.64853, 267.8532712, 1.0),
)
_J1_LOW = (
    (72362614232.0, -7895059235.0, 242396853.1, -2972611.439, 15704.48260, -30.16036606),
    (144725228442.0, 2300535178.0, 18583304.74, 99447.43394, 376.9991397, 1.0),
)
_Y0_LOW = (
    (-2957821389.0, 7062834065.0, -512359803.6, 10879881.29, -86327.92757, 228.4622733),
    (40076544269.0, 745249964.8, 7189466.438, 47447.26470, 226.1030244, 1.0),
)
_Y1_LOW = (
    (-0.4900604943e13, 0.1275274390e13, -0.5153438139e11, 0.7349264551e9, -0.4237922726e7,
     0.8511937935e4),
    (0.2499580570e14, 0.4244419664e12, 0.3733650367e10, 0.2245904002e8, 0.1020426050e6,
     0.3549632885e3, 1.0),
)  # fmt: skip

# Orders 0 and 1 from x = 8: P and Q in y = (8/x)^2; the phase is subtracted from x.
_J0_HIGH = (
    (1.0, -0.1098628267e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6),  # Excel's
    (-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934935152e-7),
)
_Y0_HIGH = (
    (1.0, -0.1098628627e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6),
    (-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934945152e-7),
)
_J1_HIGH = (
    (1.0, 0.183105e-2, -0.3516396496e-4, -0.3516396496e-4, 0.2457520174e-5,  # Excel's repeat
     -0.240337019e-6),
    (0.04687499995, -0.2002690873e-3, 0.8449199096e-5, -0.88228987e-6, 0.105787412e-6),
)  # fmt: skip
_Y1_HIGH = (
    (1.0, 0.183105e-2, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6),
    (0.04687499995, -0.2002690873e-3, 0.8449199096e-5, -0.88228987e-6, 0.105787412e-6),
)
_PHASE_0, _PHASE_1 = 0.785398164, 2.356194491

# I0 and I1: in (x/3.75)^2 below 3.75, in 3.75/x above it (times e^x / sqrt(x)).
_I0_LOW = (1.0, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.360768e-1, 0.45813e-2)
_I0_HIGH = (0.39894228, 0.1328592e-1, 0.225319e-2, -0.157565e-2, 0.916281e-2, -0.2057706e-1,
            0.2635537e-1, -0.1647633e-1, 0.392377e-2)  # fmt: skip
_I1_LOW = (0.5, 0.87890594, 0.51498869, 0.15084934, 0.2658733e-1, 0.301532e-2, 0.32411e-3)
_I1_HIGH = (0.39894228, -0.3988024e-1, -0.362018e-2, 0.163801e-2, -0.1031555e-1, 0.2282967e-1,
            -0.2895312e-1, 0.1787654e-1, -0.420059e-2)  # fmt: skip

# K0 and K1: in x^2/4 up to 2 (with the I term), in 2/x above it (times e^-x / sqrt(x)).
_K0_LOW = (-0.57721566, 0.42278420, 0.23069756, 0.3488590e-1, 0.262698e-2, 0.10750e-3, 0.74e-5)
_K0_HIGH = (1.25331414, -0.7832358e-1, 0.2189568e-1, -0.1062446e-1, 0.587872e-2, -0.251540e-2,
            0.53208e-3)  # fmt: skip
_K1_LOW = (1.0, 0.15443144, -0.67278579, -0.18156897, -0.1919402e-1, -0.110404e-2, -0.4686e-4)
_K1_HIGH = (1.25331414, 0.23498619, -0.3655620e-1, 0.1504268e-1, -0.780353e-2, 0.325614e-2,
            -0.68245e-3)  # fmt: skip


def _horner(y: float, coefficients: Coefficients) -> float:
    """c0 + y * (c1 + y * (c2 + ...)), innermost first, as the nested form evaluates it."""
    total = coefficients[-1]
    for c in reversed(coefficients[:-1]):
        total = c + y * total
    return total


def _angle(angle: float) -> None:
    """Refuse an angle as Excel does for J and Y (``ValueError``, which is ``#NUM!``)."""
    if abs(angle) >= _LARGEST_ANGLE:
        raise ValueError("angle too large")


def _scale_i(ax: float) -> float:
    """``exp(ax) / sqrt(ax)`` for I, refused past Excel's limit (``OverflowError``, ``#NUM!``)."""
    if ax > _LARGEST_I:
        raise OverflowError("I past its limit")
    return math.exp(ax) / math.sqrt(ax)


def _waves(
    ax: float, phase: float, pq: tuple[Coefficients, Coefficients]
) -> tuple[float, float, float, float, float]:
    """The asymptotic form's parts from x = 8: sqrt(2 / (pi x)), z = 8/x, P, Q and the angle."""
    z = 8.0 / ax
    y = z * z
    angle = ax - phase
    _angle(angle)
    return math.sqrt(_TWO_OVER_PI / ax), z, _horner(y, pq[0]), _horner(y, pq[1]), angle


def _j0(x: float) -> float:
    ax = abs(x)
    if ax < 8.0:
        y = x * x
        return _horner(y, _J0_LOW[0]) / _horner(y, _J0_LOW[1])
    amplitude, z, p, q, angle = _waves(ax, _PHASE_0, _J0_HIGH)
    return amplitude * (math.cos(angle) * p - z * math.sin(angle) * q)


def _j1(x: float) -> float:
    ax = abs(x)
    if ax < 8.0:
        y = x * x
        return x * _horner(y, _J1_LOW[0]) / _horner(y, _J1_LOW[1])
    amplitude, z, p, q, angle = _waves(ax, _PHASE_1, _J1_HIGH)
    value = amplitude * (math.cos(angle) * p - z * math.sin(angle) * q)
    return -value if x < 0.0 else value


def _y0(x: float) -> float:
    if x < 8.0:
        y = x * x
        rational = _horner(y, _Y0_LOW[0]) / _horner(y, _Y0_LOW[1])
        return rational + _TWO_OVER_PI * (_j0(x) * math.log(x))  # Excel's grouping
    amplitude, z, p, q, angle = _waves(x, _PHASE_0, _Y0_HIGH)
    return amplitude * (math.sin(angle) * p + z * (math.cos(angle) * q))  # Excel's grouping


def _y1(x: float) -> float:
    if x < 8.0:
        y = x * x
        rational = x * _horner(y, _Y1_LOW[0]) / _horner(y, _Y1_LOW[1])
        return rational + _TWO_OVER_PI * (_j1(x) * math.log(x) - 1.0 / x)
    amplitude, z, p, q, angle = _waves(x, _PHASE_1, _Y1_HIGH)
    return amplitude * (math.sin(angle) * p + z * (math.cos(angle) * q))  # Excel's grouping


def _i0(x: float) -> float:
    ax = abs(x)
    if ax < 3.75:
        t = x / 3.75
        return _horner(t * t, _I0_LOW)
    return _scale_i(ax) * _horner(3.75 / ax, _I0_HIGH)


def _i1(x: float) -> float:
    ax = abs(x)
    if ax < 3.75:
        t = x / 3.75
        value = ax * _horner(t * t, _I1_LOW)
    else:
        value = _horner(3.75 / ax, _I1_HIGH) * _scale_i(ax)
    return -value if x < 0.0 else value


def _k0(x: float) -> float:
    if x <= 2.0:
        return (-math.log(x / 2.0) * _i0(x)) + _horner(x * x / 4.0, _K0_LOW)
    return (math.exp(-x) / math.sqrt(x)) * _horner(2.0 / x, _K0_HIGH)


def _k1(x: float) -> float:
    if x <= 2.0:
        return (math.log(x / 2.0) * _i1(x)) + (1.0 / x) * _horner(x * x / 4.0, _K1_LOW)
    return (math.exp(-x) / math.sqrt(x)) * _horner(2.0 / x, _K1_HIGH)


def _spend(ctx: Context, steps: int) -> None:
    """Charge a loop of ``steps`` iterations to ``ReadLimits.max_function_steps``."""
    spend = getattr(ctx, "spend_steps", None)
    if spend is not None and steps > 0:
        spend(steps)


def _upward(
    first: Callable[[float], float],
    second: Callable[[float], float],
    n: int,
    x: float,
    sign: float,
    ctx: Context,
) -> float:
    """Order n from orders 0 and 1 by f(k+1) = (2k/x) f(k) + sign * f(k-1).

    Y, and J above n, take ``sign`` -1; K takes +1. The loop stops once the value is not finite,
    or is 0 with the one before it, since it then stays so.
    """
    if n == 0:
        return first(x)
    if n == 1:
        return second(x)
    _spend(ctx, n - 1)
    ratio = 2.0 / x
    before, current = first(x), second(x)
    for k in range(1, n):
        before, current = current, k * ratio * current + sign * before
        if not math.isfinite(current) or (current == 0.0 and before == 0.0):
            break
    return current


def _jn(n: int, x: float, ctx: Context) -> float:
    if n == 0:
        return _j0(x)
    if n == 1:
        return _j1(x)
    ax = abs(x)
    if ax == 0.0:
        return 0.0
    value = _upward(_j0, _j1, n, ax, -1.0, ctx) if ax > n else _downward_j(n, ax, ctx)
    return -value if x < 0.0 and n & 1 else value


def _downward_j(n: int, ax: float, ctx: Context) -> float:
    """J order n by Miller's method.

    The recurrence f(k-1) = (2k/x) f(k) - f(k+1) runs down from well above n, rescaled as it
    grows, and the result is normalized by J0 + 2 (J2 + J4 + ...) = 1.
    """
    top = 2 * ((n + int(math.sqrt(_START * n))) // 2)
    _spend(ctx, top)
    ratio = 2.0 / ax
    above, here = 0.0, 1.0
    wanted = evens = 0.0
    even = False
    for k in range(top, 0, -1):
        above, here = here, k * ratio * here - above
        if abs(here) > _RESCALE_ABOVE:
            here *= _RESCALE_BY
            above *= _RESCALE_BY
            wanted *= _RESCALE_BY
            evens *= _RESCALE_BY
        if even:
            evens += here
        even = not even
        if k == n:
            wanted = above
    return wanted / (2.0 * evens - here)


def _in(n: int, x: float, ctx: Context) -> float:
    """I order n: orders 0 and 1 directly, higher ones by Miller's method normalized by I0."""
    if n == 0:
        return _i0(x)
    if n == 1:
        return _i1(x)
    if x == 0.0:
        return 0.0
    top = 2 * (n + int(math.sqrt(_START * n)))
    _spend(ctx, top)
    ratio = 2.0 / abs(x)
    above, here, wanted = 0.0, 1.0, 0.0
    for k in range(top, 0, -1):
        above, here = here, above + k * ratio * here
        if abs(here) > _RESCALE_ABOVE:
            wanted *= _RESCALE_BY
            here *= _RESCALE_BY
            above *= _RESCALE_BY
        if k == n:
            wanted = above
    value = wanted * (_i0(x) / here)
    return -value if x < 0.0 and n & 1 else value


def _bessel(
    x: Any,
    n: Any,
    ctx: Context,
    compute: Callable[[int, float, Context], float],
    positive: bool,
) -> float | ErrorValue:
    a, b = _strict(x, ctx), _strict(n, ctx)
    if isinstance(a, ErrorValue):
        return a
    if isinstance(b, ErrorValue):
        return b
    if b < 0.0 or b >= _LARGEST_ORDER + 1 or (positive and a <= 0.0):
        return NUM  # a negative n is refused before it is truncated: -0.5 is #NUM!
    order = math.trunc(b)
    try:
        result = compute(order, a, ctx)
    except (OverflowError, ZeroDivisionError, ValueError):
        return NUM
    return result if math.isfinite(result) else NUM


@function("BESSELI", "value", "value", category=E)
def BESSELI(ctx: Context, x: Any, n: Any) -> float | ErrorValue:
    """The modified Bessel function In(x)."""
    return _bessel(x, n, ctx, _in, positive=False)


@function("BESSELJ", "value", "value", category=E, ulps=8, ulps_reason=_TRIG)
def BESSELJ(ctx: Context, x: Any, n: Any) -> float | ErrorValue:
    """The Bessel function Jn(x)."""
    return _bessel(x, n, ctx, _jn, positive=False)


@function("BESSELK", "value", "value", category=E)
def BESSELK(ctx: Context, x: Any, n: Any) -> float | ErrorValue:
    """The modified Bessel function Kn(x)."""
    return _bessel(
        x, n, ctx, lambda order, a, c: _upward(_k0, _k1, order, a, 1.0, c), positive=True
    )


@function("BESSELY", "value", "value", category=E, ulps=8, ulps_reason=_TRIG)
def BESSELY(ctx: Context, x: Any, n: Any) -> float | ErrorValue:
    """The Bessel function Yn(x), also called the Weber or Neumann function."""
    return _bessel(
        x, n, ctx, lambda order, a, c: _upward(_y0, _y1, order, a, -1.0, c), positive=True
    )
