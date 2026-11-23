"""XIRR's scan and bisection on the cases its captures do not reach (QUIRKS.md, ``XIRR``)."""

from __future__ import annotations

import math

import pytest

from xlrecalc._engine import Engine
from xlrecalc._functions._financial import (
    _XIRR_TOP,
    _xirr_points,
    _xirr_solve,
    _xirr_value,
)
from xlrecalc._values import ErrorValue

NUM = ErrorValue.NUM


def _first(g: float, n: int) -> list[float]:
    points = _xirr_points(g)
    return [next(points) for _ in range(n)]


def test_points_from_a_positive_guess():
    assert _first(0.1, 6) == [-1.0, -0.1, 0.0, 0.1, 0.2, 0.4]
    assert _first(0.0, 4) == [-1.0, -1e-5, 0.0, 1e-5]
    assert max(_xirr_points(0.1)) <= _XIRR_TOP < 2 * max(_xirr_points(0.1))


def test_points_from_a_guess_of_one_or_more_have_no_negatives():
    assert _first(1.0, 3) == [0.0, 1.0, 2.0]
    assert _first(2.5, 2) == [0.0, 2.5]


def test_points_from_a_negative_guess_stop_at_zero():
    assert list(_xirr_points(-0.3)) == [-1.0, -0.3, 0.0]


def test_value_at_minus_one():
    assert _xirr_value([-100.0, 50.0], [0.0, 1.0], -1.0) == math.inf
    assert math.isnan(_xirr_value([-100.0, 300.0, -100.0], [0.0, 1.0, 2.0], -1.0))
    # a date before the first: 0 to a negative power is infinite, the flow's value 0
    assert _xirr_value([-100.0, 50.0], [0.0, -1.0], -1.0) == -100.0


def test_value_past_overflow_is_zero_for_that_flow():
    assert _xirr_value([-1.0, 1.0], [0.0, 400.0], 1e30) == -1.0


def test_a_point_that_is_not_a_number_is_skipped():
    """(-100, 300, -100) is NaN at -1; counted, it would make a change toward the wrong sign."""
    got = _xirr_solve([-100.0, 300.0, -100.0], [0.0, 1.0, 2.0], 0.1)
    assert isinstance(got, float)
    assert got == pytest.approx((1 + math.sqrt(5)) / 2, rel=1e-7)


def test_a_zero_in_the_scan_counts_as_negative():
    """The root 1 of (-1, 2) is the scan point 1.0 from the guess 0.25: [0.5, 1.0] is bisected."""
    got = _xirr_solve([-1.0, 2.0], [0.0, 1.0], 0.25)
    assert isinstance(got, float)
    assert 1 - 1e-7 < got < 1


def test_a_zero_midpoint_replaces_b():
    """(1, -2): XNPV is negative below the root 1, so a zero at m taken as a's side moves up."""
    got = _xirr_solve([1.0, -2.0], [0.0, 1.0], 0.1)
    assert isinstance(got, float)
    assert 1 - 1e-7 < got < 1


def test_the_first_change_must_run_toward_the_first_flow():
    # (-100, 230, -132) has roots 0.1 and 0.2; XNPV rises through 0 at 0.1, away from -100's sign
    assert _xirr_solve([-100.0, 230.0, -132.0], [0.0, 1.0, 2.0], 0.05) == NUM
    assert _xirr_solve([100.0, -230.0, 132.0], [0.0, 1.0, 2.0], 0.05) == NUM


def test_no_change_is_num():
    assert _xirr_solve([100.0, 50.0], [0.0, 1.0], 0.1) == NUM


def test_a_negative_guess_finds_a_negative_root_only():
    got = _xirr_solve([-100.0, 50.0], [0.0, 1.0], -0.1)
    assert isinstance(got, float)
    assert got == pytest.approx(-0.5, rel=1e-7)
    assert _xirr_solve([-100.0, 110.0], [0.0, 1.0], -0.1) == NUM
    got = _xirr_solve([-100.0, 110.0], [0.0, 1.0], 0.5)
    assert isinstance(got, float)
    assert got == pytest.approx(0.1, rel=1e-7)


def test_a_guess_of_one_or_more_misses_a_negative_root():
    assert _xirr_solve([-100.0, 50.0], [0.0, 1.0], 1.0) == NUM
    got = _xirr_solve([-100.0, 50.0], [0.0, 1.0], 0.99)
    assert isinstance(got, float)
    assert got == pytest.approx(-0.5, rel=1e-7)


def test_the_top_of_the_scan():
    below, above = _XIRR_TOP / 4, _XIRR_TOP * 4
    got = _xirr_solve([-1.0, 1.0 + below], [0.0, 1.0], 0.1)
    assert isinstance(got, float)
    assert got == pytest.approx(below, rel=1e-7)
    assert _xirr_solve([-1.0, 1.0 + above], [0.0, 1.0], 0.1) == NUM


def test_a_scan_through_overflow():
    got = _xirr_solve([-1.0, 1e300], [0.0, 100.0], 0.1)
    assert isinstance(got, float)
    assert got == pytest.approx(999.0, rel=1e-7)


@pytest.mark.parametrize("guess", [-1.0, -2.5])
def test_a_guess_of_minus_one_or_below_is_num(guess):
    assert _formula(f"XIRR(A1:A2,B1:B2,{guess})") == NUM


def _formula(text: str, **more: object) -> object:
    cells = {"A1": -100.0, "A2": 110.0, "B1": 39814.0, "B2": 40179.0, "C1": f"={text}", **more}
    engine = Engine.from_cells({"S": cells})
    return engine.get(0, 0, 2)


def test_an_empty_guess_is_the_default_and_a_blank_cell_zero():
    default, zero = _formula("XIRR(A1:A2,B1:B2)"), _formula("XIRR(A1:A2,B1:B2,0)")
    assert default != zero  # the guess shows in a result's last digits
    assert _formula("XIRR(A1:A2,B1:B2,)") == default
    assert _formula("XIRR(A1:A2,B1:B2,D1)") == zero
