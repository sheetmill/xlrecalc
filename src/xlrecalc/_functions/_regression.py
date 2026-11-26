"""Least squares over several x columns: LINEST's and LOGEST's fit, TREND and GROWTH.

The fit is a Householder QR decomposition with column pivoting (LAPACK's ``dgeqp3``), as captures
``linest_multi``, ``linest_ill`` and ``linest_shift`` show Excel computing it on Mac and Windows
alike. With a constant, the x columns and y lose their means first (Excel's slopes are the same to
the bit with a column shifted by 1E8), and a column of ones, last in the design, is swapped to the
front and factored first; then each step takes the column with the largest remaining norm, the
first in the current order on a tie, up to the last row. Columns are kept up to the first whose
diagonal entry is at most ``COLLINEAR`` of its own norm: it, and every column after it, has
coefficient and standard error 0, and each adds to the residual degrees of freedom. A column that
is all zero, or all one value beside the constant, has coefficient 0 too but keeps its place in the
count of columns: the degrees of freedom do not grow, and the residual sum of squares is read from
the entries after every column's place, as Excel reads it (``ssresid`` too small and ``ssreg`` too
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
"""A diagonal entry, relative to its column's own norm, at or below which the column and every
column factored after it are dropped (captures ``linest_multi``: a column 1E-6 off a sum of two
others is dropped; ``linest_shift``: so is an independent column scaled by 1E-8 factored after
it, and one scaled by 1E-4 factored before it is kept)."""

_RECOMPUTE = math.sqrt(2.0**-53)
"""A downdated remaining norm is recomputed when its square has lost this share of the column's
norm squared since the last computation (LAPACK's ``tol3z``)."""


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


def _norm(col: list[float], start: int = 0) -> float:
    """The Euclidean norm of ``col[start:]``: the squares added in order, then the root."""
    return math.sqrt(plain_sum(v * v for v in col[start:]))


def _reflector(col: list[float], i: int) -> float:
    """A Householder reflector for ``col[i:]`` (LAPACK ``dlarfg``); returns tau.

    ``col[i]`` becomes the diagonal entry and ``col[i + 1:]`` the reflector's tail (its head is 1).
    """
    if len(col) - i <= 1:
        return 0.0
    alpha = col[i]
    rest = _norm(col, i + 1)
    if rest == 0.0:
        return 0.0
    w, z = max(abs(alpha), rest), min(abs(alpha), rest)
    beta = -math.copysign(w * math.sqrt(1.0 + (z / w) ** 2), alpha)
    tau = (beta - alpha) / beta
    r = 1.0 / (alpha - beta)
    for t in range(i + 1, len(col)):
        col[t] = r * col[t]
    col[i] = beta
    return tau


def _reflect(v: list[float], i: int, tau: float, targets: list[list[float]]) -> None:
    """Apply ``I - tau v v'`` to rows ``i:`` of each target (LAPACK ``dlarf``).

    The caller sets ``v[i]`` to 1.
    """
    if tau == 0.0:
        return
    for c in targets:
        temp = -tau * plain_sum(c[t] * v[t] for t in range(i, len(v)))
        if temp != 0.0:
            for t in range(i, len(v)):
                c[t] += v[t] * temp


def _factor(cols: list[list[float]], fixed: int) -> tuple[list[int], list[float]]:
    """A QR decomposition with column pivoting (LAPACK ``dgeqp3``), in place.

    R and the reflectors overwrite ``cols``. Column ``fixed`` (or none, -1) is swapped to the
    front and factored first; then each step takes the column with the largest remaining norm,
    the first in the current order on a tie, at every step to the last row, and swaps it into
    place. The remaining norms are downdated, and recomputed when downdating has lost half the
    digits.

    Returns:
        The original position of each column in its new place, and the reflectors' taus.
    """
    n, k = len(cols[0]), len(cols)
    perm = list(range(k))
    if fixed > 0:
        cols[0], cols[fixed] = cols[fixed], cols[0]
        perm[0], perm[fixed] = perm[fixed], perm[0]
    steps = min(n, k)
    taus = [0.0] * steps
    start = 0
    if fixed >= 0 and steps:
        col = cols[0]
        taus[0] = _reflector(col, 0)
        head, col[0] = col[0], 1.0
        _reflect(col, 0, taus[0], cols[1:])
        col[0] = head
        start = 1
    vn1 = [0.0] * k
    vn2 = [0.0] * k
    for j in range(start, k):
        vn1[j] = vn2[j] = _norm(cols[j], start)
    for i in range(start, steps):
        pvt, best = i, -1.0
        for j in range(i, k):
            if vn1[j] > best:
                pvt, best = j, vn1[j]
        if pvt != i:
            cols[pvt], cols[i] = cols[i], cols[pvt]
            perm[pvt], perm[i] = perm[i], perm[pvt]
            vn1[pvt], vn2[pvt] = vn1[i], vn2[i]
        col = cols[i]
        taus[i] = _reflector(col, i) if i < n - 1 else 0.0
        if i < k - 1:
            head, col[i] = col[i], 1.0
            _reflect(col, i, taus[i], cols[i + 1 :])
            col[i] = head
        for j in range(i + 1, k):
            if vn1[j] != 0.0:
                temp = max(1.0 - (abs(cols[j][i]) / vn1[j]) ** 2, 0.0)
                if temp * (vn1[j] / vn2[j]) ** 2 <= _RECOMPUTE:
                    vn1[j] = vn2[j] = _norm(cols[j], i + 1) if i < n - 1 else 0.0
                else:
                    vn1[j] *= math.sqrt(temp)
    return perm, taus


def fit(y: list[float], xs: list[list[float]], const: bool) -> Fit:
    """The least-squares fit of ``y`` on the columns ``xs`` (and a constant)."""
    n, k = len(y), len(xs)
    out = Fit(k)
    null = [all(v == 0.0 for v in col) or (const and all(v == col[0] for v in col)) for col in xs]
    names = [j for j in range(k) if not null[j]]
    means = [0.0] * k
    qy = list(y)
    if const:  # centered: x and y lose their means, and a column of ones comes last
        for j in names:
            means[j] = plain_sum(xs[j]) / n
        mean_y = plain_sum(y) / n
        cols = [[v - means[j] for v in xs[j]] for j in names] + [[1.0] * n]
        names.append(-1)  # -1: the constant
        qy = [v - mean_y for v in y]
    else:
        cols = [list(xs[j]) for j in names]
    own = [_norm(col) for col in cols]
    perm, taus = _factor(cols, len(cols) - 1 if const else -1)
    # the kept columns: up to the first whose diagonal entry is small beside its own norm
    r = 0
    while r < min(n, len(cols)) and abs(cols[r][r]) > COLLINEAR * own[perm[r]]:
        r += 1
    for i in range(r):
        col = cols[i]
        head, col[i] = col[i], 1.0
        _reflect(col, i, taus[i], [qy])
        col[i] = head
    # back substitution: R c = Q'y over the kept columns
    coef = qy[:r]
    for j in range(r - 1, -1, -1):
        if coef[j] != 0.0:
            coef[j] /= cols[j][j]
            for i in range(j - 1, -1, -1):
                coef[i] -= coef[j] * cols[j][i]
    p = r + sum(null)
    out.df = n - p
    out.ssresid = plain_sum(t * t for t in qy[p:]) if p < n else 0.0
    total = plain_sum((t - mean_y) ** 2 for t in y) if const else plain_sum(t * t for t in y)
    out.ssreg = total - out.ssresid
    out.regressors = p - (1 if const else 0)
    # standard errors: sey * the norm of each row of R's inverse
    inv = [[0.0] * r for _ in range(r)]
    for last in range(r):
        for i in range(last, -1, -1):
            s = 1.0 if i == last else 0.0
            for j in range(i + 1, last + 1):
                s -= cols[j][i] * inv[j][last]
            inv[i][last] = s / cols[i][i]
    sey = math.sqrt(out.ssresid / out.df) if out.df > 0 else 0.0
    shift = 0.0  # the slopes times the means, taken off the constant
    for i in range(r):
        name = names[perm[i]]
        if name == -1:
            out.b = coef[i]
        else:
            out.m[name] = coef[i]
            out.se[name] = sey * math.sqrt(plain_sum(t * t for t in inv[i])) if out.df > 0 else 0.0
    if const:
        for j in names:
            if j != -1:
                shift += out.m[j] * means[j]
        out.b = mean_y + out.b - shift
        # the constant's variance through R's inverse: g = 1 for the ones column, -mean for x
        g = [1.0 if names[perm[i]] == -1 else -means[names[perm[i]]] for i in range(r)]
        v = plain_sum(
            t * t for t in (plain_sum(inv[i][j] * g[i] for i in range(r)) for j in range(r))
        )
        out.se_b = sey * math.sqrt(v) if out.df > 0 else 0.0
    return out


def line(y: list[float], x: list[float], const: bool, stats: bool) -> tuple[tuple[Any, ...], ...]:
    """LINEST's result over one x column, from sums about the means (capture ``statistics``).

    LOGEST and GROWTH with one x column take it over the logarithms of y (capture ``empty_args``:
    their results with x omitted, and with x given, are Excel's to the bit this way and not by the
    QR decomposition of ``fit``).
    """
    n = len(y)
    if const:
        mx, my = plain_sum(x) / n, plain_sum(y) / n
    else:
        mx = my = 0.0
    sxx = sxy = 0.0
    for u, v in zip(x, y, strict=True):
        sxx += (u - mx) * (u - mx)
        sxy += (u - mx) * (v - my)
    if sxx == 0:
        slope, intercept = 0.0, my
    else:
        slope = sxy / sxx
        intercept = my - slope * mx if const else 0.0
    if not stats:
        return ((slope, intercept),)
    df = n - (2 if const else 1)
    ssreg = ssresid = 0.0
    for u, v in zip(x, y, strict=True):
        fitted = slope * u + intercept
        ssresid += (v - fitted) * (v - fitted)
        ssreg += (fitted - my) * (fitted - my)
    if df <= 0:
        return ((slope, intercept), (NA, NA), (1.0, 0.0), (NUM, df), (ssreg, ssresid))
    sey = math.sqrt(ssresid / df)
    se_slope = sey / math.sqrt(sxx) if sxx else NUM
    se_int: Any = sey * math.sqrt(1 / n + mx * mx / sxx) if const and sxx else NA
    total = ssreg + ssresid
    r2 = ssreg / total if total else 1.0
    f: Any = ssreg / (ssresid / df) if ssresid else NUM  # one regressor: 1 degree of freedom
    return (
        (slope, intercept),
        (se_slope, se_int),
        (r2, sey),
        (f, float(df)),
        (ssreg, ssresid),
    )


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
    "Excel's QR steps are reproduced in their order (pivots, layout, centering), not in their last "
    "bits (captures linest_multi, linest_ill and linest_shift, Mac and Windows alike): the "
    "coefficients within 65 units in the last place, the standard errors, r2, sey, F and the sums "
    "of squares within 203 (2.8E-14 relative), and up to 312 (4.1E-14) beside a zero or constant "
    "column or with as many rows as unknowns; TREND within 24 and GROWTH within 10; the "
    "ill-conditioned fit within 226 (3.8E-14). Past this, listed rows"
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
    if len(cols) == 1:  # LOGEST over one column: LINEST's line through the logarithms
        first, *rest = line(y, cols[0], c, s)
        return (tuple(math.exp(v) for v in first), *rest)
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
    k = len(cols)
    if log and k == 1:  # GROWTH over one column: LINEST's line through the logarithms
        f = Fit(1)
        f.m[0], f.b = line(y, cols[0], c, False)[0]
    else:
        f = fit(y, cols, c)
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
