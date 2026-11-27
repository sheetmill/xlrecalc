"""``tools/exact_regression.py``: the exact least-squares solution it compares captures with."""

from __future__ import annotations

from fractions import Fraction

from tools.exact_regression import exact


def test_an_exact_line_is_found_exactly():
    m, b = exact([3.0, 5.0, 7.0], [[1.0, 2.0, 3.0]], True)
    assert (m, b) == ([Fraction(2)], Fraction(1))


def test_a_fit_without_the_constant_and_two_columns():
    # y = 0.5 x1 - 2 x2 exactly, no constant: the solution is those coefficients, as fractions
    x1, x2 = [1.0, 2.0, 3.0, 4.0], [1.0, 0.0, 2.0, 5.0]
    y = [0.5 * a - 2 * c for a, c in zip(x1, x2, strict=True)]
    m, b = exact(y, [x1, x2], False)
    assert (m, b) == ([Fraction(1, 2), Fraction(-2)], Fraction(0))


def test_a_least_squares_fit_is_the_normal_equations_solution():
    m, b = exact([1.0, 2.0, 2.0], [[0.0, 1.0, 2.0]], True)  # slope 1/2, intercept 7/6
    assert (m, b) == ([Fraction(1, 2)], Fraction(7, 6))
