"""Excel's arithmetic on doubles: no subnormal results, and cancellation snapping (§4.4).

Captured in ``tests/captures/cancellation.tsv``. Every result smaller in magnitude than the
smallest normal double is 0 (``2.2250738585072014E-308/2`` is 0). With
``cancellation_snapping``, a formula whose last operation, outside any parentheses, is an
addition or subtraction gives 0 when the result is less than 8 units in the last place of the
left operand: ``=0.1+0.2-0.3`` and ``=(0.1+0.2)-0.3`` are 0, ``=(0.1+0.2-0.3)`` and
``=0.1+0.2-0.3+0`` are 5.55E-17. The evaluator applies ``snap`` to that last operation only.
"""

from __future__ import annotations

import math

from xlrecalc._numtext import DBL_MIN
from xlrecalc._quirks import Quirks

__all__ = ["flush", "snap"]


def flush(x: float) -> float:
    """A result as Excel keeps it: subnormal numbers become 0."""
    return 0.0 if -DBL_MIN < x < DBL_MIN else x


def snap(left: float, result: float, quirks: Quirks = Quirks.EXCEL) -> float:
    """The result of a formula's final ``left + right`` or ``left - right``."""
    if quirks.cancellation_snapping and result and abs(result) < 8 * math.ulp(left):
        return 0.0
    return flush(result)
