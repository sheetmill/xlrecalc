"""The beta, gamma, chi-squared, F and t functions on extreme input: a value, never an exception.

Shapes, degrees of freedom and values from the smallest subnormal to the largest double, and
probabilities at their edges. Each call gives a finite number or an Excel error, and none takes
long: the series and fractions stop at a number of terms (``#NUM!`` past it) and every inverse is
a bisection over the doubles' bit patterns, at most 64 evaluations. Run under the watchdog.
"""

from __future__ import annotations

import itertools
import math
import time

from tests.watchdog import bounded

V = (0.0, 5e-324, 1e-300, 1e-10, 0.5, 1 - 1e-16, 1.0, 3.7, 1e5, 1e10, 1e15, 1e300, 1.7e308, -1.0)
P = (0.0, 5e-324, 1e-300, 1e-17, 0.5, 1 - 1e-16, 1.0, 1.5)
_GRIDS = {
    "BETA_DIST": (V, V, V, (True, False)),
    "BETA_INV": (P, V, V),
    "GAMMA_DIST": (V, V, V, (True, False)),
    "GAMMA_INV": (P, V, V),
    "CHISQ_DIST": (V, V, (True, False)),
    "CHISQ_DIST_RT": (V, V),
    "CHISQ_INV": (P, V),
    "CHISQ_INV_RT": (P, V),
    "F_DIST": (V, V, V, (True, False)),
    "F_DIST_RT": (V, V, V),
    "F_INV": (P, V, V),
    "F_INV_RT": (P, V, V),
    "T_DIST": (V, V, (True, False)),
    "T_DIST_RT": (V, V),
    "T_DIST_2T": (V, V),
    "TDIST": (V, V, (1.0, 2.0)),
    "T_INV": (P, V),
    "T_INV_2T": (P, V),
    "CONFIDENCE_T": (P, V, V),
}
_SLOWEST_SECONDS = 1.0


@bounded(seconds=300.0)
def test_every_call_is_a_value_and_quick():
    from xlrecalc._functions import _continuous
    from xlrecalc._values import ErrorValue

    bad, slow = [], []
    for name, axes in _GRIDS.items():
        function = getattr(_continuous, name)
        for args in itertools.product(*axes):
            start = time.perf_counter()
            try:
                value = function(None, *args)
            except Exception as error:
                bad.append((name, args, type(error).__name__))
                continue
            if time.perf_counter() - start > _SLOWEST_SECONDS:
                slow.append((name, args))
            if not (
                isinstance(value, ErrorValue) or (type(value) is float and math.isfinite(value))
            ):
                bad.append((name, args, repr(value)))
    assert not bad, bad[:10]
    assert not slow, slow[:10]


@bounded
def test_the_cases_that_raised_give_values():
    """Each from the first sweep: an overflow, a log of 0 or a division by an underflow."""
    from xlrecalc._functions import _continuous as c
    from xlrecalc._values import ErrorValue

    assert c.BETA_DIST(None, 1.0, 1e5, 1e5, False) == 0.0  # the density at 1, B(a, b) underflows
    assert c.GAMMA_DIST(None, 0.5, 1.7e308, 0.5, True) == 0.0  # lgamma of the shape overflows
    assert c.GAMMA_DIST(None, 0.5, 1.7e308, 0.5, False) == 0.0
    assert c.CONFIDENCE_T(None, 0.05, 1.0, 1e15) == ErrorValue.NUM  # past 1E10 degrees of freedom
    density = c.F_DIST(None, 5e-324, 1.0, 1e5, False)  # x subnormal, where p underflows
    assert isinstance(density, float)
    assert math.isclose(density, 1.794802441512813e161, rel_tol=1e-9)  # 50-digit arithmetic
