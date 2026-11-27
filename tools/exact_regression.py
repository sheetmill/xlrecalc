"""Capture ``linest_ill``'s fit solved exactly, beside Excel's coefficients and the package's.

Solves the normal equations of the fit in rational arithmetic (``fractions.Fraction``, no
rounding anywhere), rounds the solution to doubles once, and prints, with the constant and
without, the largest relative distance of Excel's captured coefficients and of the package's
from it, and of the package's from Excel's. QUIRKS.md and ``tests/test_captures.py`` state these
figures.

    uv run python tools/exact_regression.py
"""

from __future__ import annotations

import sys
from fractions import Fraction
from pathlib import Path


def exact(y: list[float], xs: list[list[float]], const: bool) -> tuple[list[Fraction], Fraction]:
    """The least-squares coefficients of ``y`` on ``xs`` (and a constant), exactly.

    Returns:
        The coefficients in the order of ``xs``, and the constant (0 without one).
    """
    n = len(y)
    rows = [([Fraction(1)] if const else []) + [Fraction(x[i]) for x in xs] for i in range(n)]
    m = len(rows[0])
    normal = [[sum((r[p] * r[q] for r in rows), Fraction(0)) for q in range(m)] for p in range(m)]
    right = [sum((rows[i][p] * Fraction(y[i]) for i in range(n)), Fraction(0)) for p in range(m)]
    for p in range(m):  # Gauss-Jordan elimination; the normal matrix is positive definite
        for r in range(m):
            if r != p and normal[r][p]:
                f = normal[r][p] / normal[p][p]
                normal[r] = [a - f * b for a, b in zip(normal[r], normal[p], strict=True)]
                right[r] -= f * right[p]
    solution = [right[p] / normal[p][p] for p in range(m)]
    return (solution[1:], solution[0]) if const else (solution, Fraction(0))


def relative(a: float, b: float) -> float:
    """How far ``a`` is from ``b``, relative to ``b``."""
    return abs(a - b) / abs(b)


def main() -> None:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    # the capture generator needs openpyxl, which exact() does not
    from capture.suites import _ill_columns
    from tests.captures import load

    from xlrecalc._functions._regression import fit

    columns = _ill_columns()
    y, xs = columns[0], columns[1:]
    rows = load("linest_ill").rows
    for const, first in ((True, 0), (False, 21)):
        m, b = exact(y, xs, const)
        # LINEST's first row: the last x column's coefficient first, then the constant
        want = [float(v) for v in reversed(m)] + ([float(b)] if const else [])
        excel = [row.expected for row in rows[first : first + len(want)]]
        f = fit(list(y), [list(x) for x in xs], const)
        package = list(reversed(f.m)) + ([f.b] if const else [])
        pairs = [(e, p, w) for e, p, w in zip(excel, package, want, strict=True) if w]
        print(
            f"constant={const}: Excel within {max(relative(e, w) for e, _, w in pairs):.2g} of the"
            f" exact solution, the package within {max(relative(p, w) for _, p, w in pairs):.2g},"
            f" the package within {max(relative(p, e) for e, p, _ in pairs):.2g} of Excel"
        )


if __name__ == "__main__":
    main()
