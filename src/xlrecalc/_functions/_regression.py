"""Least squares over several x columns: LINEST's and LOGEST's fit, TREND and GROWTH.

The fit is a Householder QR decomposition, as capture ``linest_multi`` shows Excel computing it on
Mac and Windows alike: the column of ones first when there is a constant, then the x columns,
each step taking the column with the largest remaining norm. A column left with less than
``COLLINEAR`` of its own norm is collinear with the columns before it: its coefficient and
standard error are 0, and it adds to the residual degrees of freedom. A column that is all zero,
or all one value beside the constant, has coefficient 0 too but keeps its place in the count of
columns: the degrees of freedom do not grow, and the residual sum of squares is read from the
entries after every column's place, as Excel reads it (``ssresid`` too small and ``ssreg`` too
large by the same amount; QUIRKS.md).
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._arith import plain_sum
from xlrecalc._arrays import to_array
from xlrecalc._coerce import Context
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import ErrorValue, Ref

NA, NUM, REF, VALUE = ErrorValue.NA, ErrorValue.NUM, ErrorValue.REF, ErrorValue.VALUE

COLLINEAR = 1e-6
"""A column's remaining norm, relative to its own, at or below which it is dropped (capture
``linest_multi``: a column 1E-6 off a sum of two others is dropped; independent columns keep far
more)."""


def grid(arg: Any) -> list[list[Any]]:
    """An argument as rows of values."""
    if isinstance(arg, Ref):
        return [list(row) for row in to_array(arg)]
    if isinstance(arg, tuple):
        return [list(row) for row in arg]
    return [[arg]]


class Fit:
    """A least-squares fit: coefficients (one per x column), the constant, and the statistics."""

    __slots__ = ("b", "df", "k", "m", "regressors", "se", "se_b", "ssreg", "ssresid")

    def __init__(self, k: int) -> None:
        self.k = k
        self.m = [0.0] * k
        self.se = [0.0] * k
        self.b = 0.0
        self.se_b: float | ErrorValue = NA
        self.df = 0
        self.ssreg = self.ssresid = 0.0
        self.regressors = 0  # the regression's degrees of freedom (F's numerator)


def columns(ys: Any, xs: Any) -> tuple[list[float], list[list[float]]] | ErrorValue:
    """The y values and the x columns.

    x is given in columns beside a y column or in rows beside a y row; an error in either is the
    result, any other value that is not a number ``#VALUE!``, and shapes that do not fit ``#REF!``.
    """
    y_rows = grid(ys)
    flat = [v for row in y_rows for v in row]
    n = len(flat)
    if xs is MISSING:
        x_cols = [[float(i) for i in range(1, n + 1)]]
    else:
        x_rows = grid(xs)
        if len(y_rows[0]) == 1 and len(x_rows) == n:  # y a column: x in columns
            x_cols = [[row[j] for row in x_rows] for j in range(len(x_rows[0]))]
        elif len(y_rows) == 1 and len(x_rows[0]) == n:  # y a row: x in rows
            x_cols = [list(row) for row in x_rows]
        else:
            return REF
    for v in (*flat, *(v for col in x_cols for v in col)):
        if isinstance(v, ErrorValue):
            return v
    for v in (*flat, *(v for col in x_cols for v in col)):
        if type(v) is not float:
            return VALUE
    return flat, x_cols


def fit(y: list[float], xs: list[list[float]], const: bool) -> Fit:
    """The least-squares fit of ``y`` on the columns ``xs`` (and a constant)."""
    n, k = len(y), len(xs)
    out = Fit(k)
    null = [all(v == 0.0 for v in col) or (const and all(v == col[0] for v in col)) for col in xs]
    cols = ([[1.0] * n] if const else []) + [list(col) for j, col in enumerate(xs) if not null[j]]
    names = ([-1] if const else []) + [j for j in range(k) if not null[j]]  # -1: the constant
    own = [math.sqrt(plain_sum(v * v for v in col)) for col in cols]
    qy = list(y)
    order: list[int] = []  # positions in cols, in the order they are factored
    diag: list[float] = []
    # The columns not yet factored, in the order a swap of the chosen column into place leaves
    # them: on a tie the first in this order is taken.
    left = list(range(len(cols)))
    step = 0
    while left and step < n:
        if const and step == 0:
            at = 0
        else:
            norms = [math.sqrt(plain_sum(cols[j][i] ** 2 for i in range(step, n))) for j in left]
            at = norms.index(max(norms))
        pick = left[at]
        left[0], left[at] = left[at], left[0]
        left.pop(0)
        col = cols[pick]
        a = math.sqrt(plain_sum(col[i] ** 2 for i in range(step, n)))
        if own[pick] == 0.0 or a <= COLLINEAR * own[pick]:
            continue  # collinear with the columns before it: dropped
        v = col[step:]
        v[0] += math.copysign(a, v[0])
        vv = plain_sum(t * t for t in v)
        for j in [*left, pick]:
            c = cols[j]
            s = 2.0 * plain_sum(v[i] * c[step + i] for i in range(len(v))) / vv
            for i in range(len(v)):
                c[step + i] -= s * v[i]
        s = 2.0 * plain_sum(v[i] * qy[step + i] for i in range(len(v))) / vv
        for i in range(len(v)):
            qy[step + i] -= s * v[i]
        order.append(pick)
        diag.append(col[step])
        step += 1
    r = len(order)
    # back substitution: R c = Q'y over the kept columns
    coef = [0.0] * r
    for i in range(r - 1, -1, -1):
        s = qy[i]
        for j in range(i + 1, r):
            s -= cols[order[j]][i] * coef[j]
        coef[i] = s / cols[order[i]][i]
    p = r + sum(null)
    out.df = n - p
    out.ssresid = plain_sum(t * t for t in qy[p:]) if p < n else 0.0
    if const:
        mean = plain_sum(y) / n
        total = plain_sum((t - mean) ** 2 for t in y)
    else:
        total = plain_sum(t * t for t in y)
    out.ssreg = total - out.ssresid
    out.regressors = p - (1 if const else 0)
    # standard errors: sey * the norm of each row of R's inverse
    inv = [[0.0] * r for _ in range(r)]
    for last in range(r):
        for i in range(last, -1, -1):
            s = 1.0 if i == last else 0.0
            for j in range(i + 1, last + 1):
                s -= cols[order[j]][i] * inv[j][last]
            inv[i][last] = s / cols[order[i]][i]
    sey = math.sqrt(out.ssresid / out.df) if out.df > 0 else 0.0
    for i, pos in enumerate(order):
        name = names[pos]
        se = sey * math.sqrt(plain_sum(t * t for t in inv[i])) if out.df > 0 else 0.0
        if name == -1:
            out.b, out.se_b = coef[i], se
        else:
            out.m[name], out.se[name] = coef[i], se
    return out


def table(f: Fit, const: bool, stats: bool, log: bool = False) -> tuple[tuple[Any, ...], ...]:
    """LINEST's (or LOGEST's) result.

    The coefficients, last column first, and the constant; with ``stats``, four more rows
    (standard errors; r2 and sey; F and df; ssreg and ssresid), ``#N/A`` past the second column.
    """
    k = f.k
    width = k + 1
    m = [math.exp(v) if log else v for v in reversed(f.m)]
    b = math.exp(f.b) if log else f.b
    first = (*m, b if const else (1.0 if log else 0.0))
    if not stats:
        return (first,)
    pad = (NA,) * (width - 2)
    total = f.ssreg + f.ssresid
    r2 = f.ssreg / total if total else 1.0
    if f.df > 0:
        sey: Any = math.sqrt(f.ssresid / f.df)
        fstat: Any = (
            (f.ssreg / f.regressors) / (f.ssresid / f.df) if f.ssresid and f.regressors else NUM
        )
    else:
        sey, fstat = 0.0, NUM
    return (
        first,
        (*reversed(f.se), f.se_b if const else NA),
        (r2, sey, *pad),
        (fstat, float(f.df), *pad),
        (f.ssreg, f.ssresid, *pad),
    )


_FIT = (
    "Excel's QR decomposition is not reproduced step for step, and its last bits differ "
    "(capture linest_multi, Mac and Windows alike): the coefficients within 31 units in the last "
    "place, the standard errors, r2, sey, F and the sums of squares within 67, and up to 195 "
    "(2.5E-14 relative) beside a zero or constant column or with as many rows as unknowns; "
    "TREND within 14 and GROWTH within 6. Past this, listed rows"
)


def _flags(const: Any, stats: Any = False) -> tuple[bool, bool]:
    return (True if const is MISSING else bool(const)), (False if stats is MISSING else bool(stats))


def _logs(y: list[float]) -> list[float] | ErrorValue:
    if any(v <= 0 for v in y):
        return NUM
    return [math.log(v) for v in y]


def several(ys: Any, xs: Any, const: Any, stats: Any, log: bool = False) -> Any:
    """LINEST's or LOGEST's result over several x columns, or ``None`` for one column."""
    data = columns(ys, xs)
    if isinstance(data, ErrorValue):
        return data
    y, cols = data
    if len(cols) < 2 and not log:
        return None
    if log:
        logs = _logs(y)
        if isinstance(logs, ErrorValue):
            return logs
        y = logs
    c, s = _flags(const, stats)
    return table(fit(y, cols, c), c, s, log)


@function(
    "LOGEST", "array", "array", "logical", "logical", min_args=1, category="statistical",
    rel=5e-14, rel_reason=_FIT,
)  # fmt: skip
def LOGEST(ctx: Context, ys: Any, xs: Any = MISSING, const: Any = True, stats: Any = False) -> Any:
    """The exponential curve through the points.

    LINEST of the logarithms of ``ys``, its coefficients raised back (capture ``linest_multi``).
    """
    return several(ys, xs, const, stats, log=True)


def _predict(ys: Any, xs: Any, new: Any, const: Any, log: bool) -> Any:
    data = columns(ys, xs)
    if isinstance(data, ErrorValue):
        return data
    y, cols = data
    if log:
        logs = _logs(y)
        if isinstance(logs, ErrorValue):
            return logs
        y = logs
    c, _ = _flags(const)
    f = fit(y, cols, c)
    k = len(cols)
    as_row = len(grid(ys)) == 1 and len(grid(ys)[0]) > 1
    if new is MISSING:
        points = [[col[i] for col in cols] for i in range(len(y))]
        shape = (1, len(points)) if as_row else (len(points), 1)
    else:
        rows = grid(new)
        for v in (v for row in rows for v in row):
            if isinstance(v, ErrorValue):
                return v
            if type(v) is not float:
                return VALUE
        if k == 1:  # one x column: each new value on its own, in the new values' shape
            fitted = [[f.b + f.m[0] * v for v in row] for row in rows]
            return tuple(tuple(math.exp(v) if log else v for v in row) for row in fitted)
        if as_row:
            if len(rows) != k:
                return REF
            points = [[rows[j][i] for j in range(k)] for i in range(len(rows[0]))]
            shape = (1, len(points))
        else:
            if len(rows[0]) != k:
                return REF
            points = rows
            shape = (len(points), 1)
    values = [f.b + plain_sum(f.m[j] * pt[j] for j in range(k)) for pt in points]
    values = [math.exp(v) if log else v for v in values]
    if shape[0] == 1:
        return (tuple(values),)
    return tuple((v,) for v in values)


@function(
    "TREND", "array", "array", "array", "logical", min_args=1, category="statistical",
    rel=5e-14, rel_reason=_FIT,
)  # fmt: skip
def TREND(ctx: Context, ys: Any, xs: Any = MISSING, new: Any = MISSING, const: Any = True) -> Any:
    """The least-squares fit's values at the new x values, or at the known ones."""
    return _predict(ys, xs, new, const, log=False)


@function(
    "GROWTH", "array", "array", "array", "logical", min_args=1, category="statistical",
    rel=5e-14, rel_reason=_FIT,
)  # fmt: skip
def GROWTH(ctx: Context, ys: Any, xs: Any = MISSING, new: Any = MISSING, const: Any = True) -> Any:
    """The exponential fit's values at the new x values, or at the known ones."""
    return _predict(ys, xs, new, const, log=True)
