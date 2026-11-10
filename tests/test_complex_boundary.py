"""The complex rows whose text another platform's math library can change.

A complex result is text: each part rounded to 15 significant digits. Where a part computed through
``sin``, ``cos``, ``exp``, ``log``, ``sinh``, ``cosh`` or ``atan2`` lies within one unit in the last
place of a midpoint between two 15-digit decimals, a last-bit difference in another platform's
library rounds the text the other way (Linux did on two rows). And an imaginary part of exactly 1
is written ``i``, the double below it ``1i``. Those rows match Excel here and are listed for other
platforms (``OTHER_PLATFORMS`` in ``tests/test_captures.py``); this test derives them, on the
platform ``docs/coverage.md`` records, where the parts are the ones the captures were matched with.
"""

from __future__ import annotations

import math
import re
from decimal import Decimal, localcontext
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from tests import test_captures as tc
from tests.captures import load

LIBRARY = frozenset({
    "imexp", "imln", "imlog10", "imlog2", "imsin", "imcos", "imtan", "imcot", "imsec", "imcsc",
    "imsinh", "imcosh", "imsech", "imcsch", "imsqrt", "impower",
})  # fmt: skip
POLAR = frozenset({"imsqrt", "impower"})  # an imaginary part of 1 from sin, not set exactly
WITHIN_ULPS = 1.0


def midpoint_ulps(x: float) -> float:
    """Units in the last place from x to the nearest midpoint between 15-digit decimals."""
    if x == 0 or not math.isfinite(x):
        return math.inf
    with localcontext() as context:
        context.prec = 80
        d = Decimal(x).copy_abs()
        q = Decimal(1).scaleb(d.adjusted() - 14)
        low = (d / q).to_integral_value(rounding="ROUND_FLOOR") * q
        gap = min(abs(d - (low + q / 2)), abs(d - (low - q / 2)))
        return float(gap / Decimal(math.ulp(x)))


def boundary_rows(monkeypatch: Any) -> set[int]:
    """The complex_family cases near a rounding boundary, from the parts the engine writes."""
    import xlrecalc._functions._complex as cx

    parts: list[tuple[float, float]] = []
    text = cx._text

    def recording(real: float, imaginary: float, suffix: str) -> Any:
        parts.append((real, imaginary))
        return text(real, imaginary, suffix)

    monkeypatch.setattr(cx, "_text", recording)
    suite = load("complex_family")
    near = set()
    for row in suite.rows:
        if row.tag not in LIBRARY or not isinstance(row.expected, str):
            continue
        parts.clear()
        tc.run(SimpleNamespace(fixed={}, rows=[row], date1904=False, captured=suite.captured))
        if not parts:
            continue
        real, imaginary = parts[-1]
        if min(midpoint_ulps(real), midpoint_ulps(imaginary)) <= WITHIN_ULPS or (
            row.tag in POLAR and abs(imaginary) == 1.0
        ):
            near.add(row.case)
    return near


def test_rows_near_a_rounding_boundary_are_listed_for_other_platforms(monkeypatch):
    if not tc._reference_platform():
        pytest.skip("the parts are the reference platform's")
    listed = {
        case
        for (suite, case), reason in tc.KNOWN_DEVIATIONS.items()
        if suite == "complex_family" and reason == tc._COMPLEX_TEXT_BOUNDARY
    }
    assert listed == boundary_rows(monkeypatch)


def test_quirks_states_their_number():
    quirks = (Path(__file__).parents[1] / "QUIRKS.md").read_text(encoding="utf-8")
    found = re.search(r"([\d,]+) rows of `complex_family` match Excel here", quirks)
    assert found, "QUIRKS.md: the complex boundary rows are not stated"
    listed = [r for r in tc.KNOWN_DEVIATIONS.values() if r == tc._COMPLEX_TEXT_BOUNDARY]
    assert int(found[1].replace(",", "")) == len(listed)
