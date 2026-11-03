"""Financial functions: annuities, cash flows and depreciation.

PV, FV, PMT, IPMT, PPMT, NPER, NPV, XNPV, SLN, VDB (capture ``finance``); IRR, XIRR, RATE (capture
``iterative``).

Excel computes the growth factor ``(1 + rate) ** nper`` by repeated squaring when ``nper`` is
an integer and as ``exp(nper * log(1 + rate))`` otherwise, and PV and FV use ``q - 1`` as it
falls out of that; PMT instead works with ``log1p`` and ``expm1`` and is accurate for tiny
rates (``PMT(1E-10,10,1000)`` is -100.000000055).
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._arith import flush
from xlrecalc._arrays import to_array
from xlrecalc._coerce import Context, numbers, to_number
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Ref

NUM, DIV0, VALUE = ErrorValue.NUM, ErrorValue.DIV0, ErrorValue.VALUE
C = "financial"
_LIBM = (
    "Excel's exp, log1p and expm1 differ from the platform's in the last bits; within 4 units "
    "in the last place on every captured row"
)
_ITERATION = (
    "Excel's Newton iterates rest on its own NPV, whose last bits are not reproduced, and it "
    "guarantees its result only to its documented stopping tolerance: within 1E-9 relative"
)
_VDB_SUM = (
    "the period-by-period depreciation sums round differently from Excel's in the last bits; "
    "within 4 units in the last place except listed rows"
)
_NPV_SUM = (
    "Excel's summation of the discounted flows is not reproduced in the last bits (no tried "
    "order, power or compensated sum matches every row); within 4 units in the last place"
)


def _result(x: float) -> float | ErrorValue:
    return flush(x) if math.isfinite(x) else NUM


def growth(rate: float, n: float) -> float | None:
    """``(1 + rate) ** n`` as Excel computes it; None where it is undefined."""
    b = 1.0 + rate
    if n.is_integer():
        k = int(abs(n))
        q = 1.0
        while k:
            if k & 1:
                q *= b
            b *= b
            k >>= 1
        if n < 0:
            return math.inf if q == 0 else 1.0 / q
        return q
    if b == 0:
        return 0.0 if n > 0 else None
    if b < 0:
        return None
    try:
        return math.exp(n * math.log(b))
    except OverflowError:
        return math.inf


def _pv(rate: float, n: float, pmt: float, fv: float, due: bool) -> float | ErrorValue:
    if rate == 0:
        return _result(-(fv + pmt * n))
    if rate == -1:
        return NUM if n == 0 else DIV0
    q = growth(rate, n)
    if q is None:
        return NUM
    try:
        return _result(-(fv + pmt * ((q - 1) / rate * (1 + rate * due))) / q)
    except ZeroDivisionError:
        return NUM


def _fv(rate: float, n: float, pmt: float, pv: float, due: bool) -> float | ErrorValue:
    if rate == 0:
        return _result(-(pv + pmt * n))
    if rate == -1 and n <= 0:
        return NUM if n == 0 else DIV0
    q = growth(rate, n)
    if q is None:
        return NUM
    return _result(-(pv * q + pmt * ((q - 1) / rate * (1 + rate * due))))


def _pmt(rate: float, n: float, pv: float, fv: float, due: bool) -> float | ErrorValue:
    if n == 0:
        return NUM
    if rate == 0:
        return _result(-(pv + fv) / n)
    if rate <= -1:
        return NUM
    try:
        log = math.log1p(rate)
        if n * log > 709:  # q overflows: the payment is the interest on pv
            return _result(-(pv * rate) / (1 + rate * due))
        q = math.exp(n * log)
        qm = math.expm1(n * log)
        return _result(-(pv * q / qm * rate + fv / qm * rate) / (1 + rate * due))
    except (OverflowError, ZeroDivisionError):
        return NUM


@function("PV", "number", "number", "number", "number", "number", min_args=3, category=C)
def PV(ctx: Context, rate: float, n: float, pmt: float, fv: float = 0.0, due: float = 0.0) -> Any:
    """Present value of an annuity; any nonzero type means payments at the start."""
    return _pv(rate, n, pmt, fv, due != 0)


@function("FV", "number", "number", "number", "number", "number", min_args=3, category=C)
def FV(ctx: Context, rate: float, n: float, pmt: float, pv: float = 0.0, due: float = 0.0) -> Any:
    """Future value of an annuity."""
    return _fv(rate, n, pmt, pv, due != 0)


@function(
    "PMT",
    "number",
    "number",
    "number",
    "number",
    "number",
    min_args=3,
    category=C,
    ulps=4,
    ulps_reason=_LIBM,
)
def PMT(ctx: Context, rate: float, n: float, pv: float, fv: float = 0.0, due: float = 0.0) -> Any:
    """Payment per period of an annuity."""
    return _pmt(rate, n, pv, fv, due != 0)


@function("NPER", "number", "number", "number", "number", "number", min_args=3, category=C)
def NPER(
    ctx: Context, rate: float, pmt: float, pv: float, fv: float = 0.0, due: float = 0.0
) -> Any:
    """Number of periods of an annuity."""
    if rate == 0:
        if pmt == 0:
            return DIV0
        return _result(-(pv + fv) / pmt)
    if rate <= -1:
        return NUM
    log = math.log(1 + rate)
    if log == 0:
        return DIV0
    a = pmt * (1 + rate * (due != 0))
    den = a + pv * rate
    if den == 0:
        return NUM
    ratio = (a - fv * rate) / den
    if ratio <= 0:
        return NUM
    return _result(math.log(ratio) / log)


def _balance(pv: float, fv: float, qn: float, qj: float, qm: float) -> float:
    """The balance after ``j`` periods, from ``q**n``, ``q**j`` and ``q**n - 1``."""
    return (pv * (qn - qj) - fv * (qj - 1)) / qm


def _parts(
    rate: float, per: float, n: float, pv: float, fv: float, due: bool
) -> tuple[float, float | ErrorValue] | ErrorValue:
    """The interest and principal parts of the payment of period ``per``."""
    if per < 1 or per >= n + 1:
        return NUM
    if due and per == 1:
        pmt = _pmt(rate, n, pv, fv, due)
        return (0.0, pmt)  # the principal part may be an error
    if rate == 0:
        pmt = -(pv + fv) / n
        return (0.0, pmt)
    if rate <= -1:
        return NUM
    pmt = _pmt(rate, n, pv, fv, due)
    if isinstance(pmt, ErrorValue):
        return pmt
    qn = growth(rate, n)
    if qn is None or qn == 1:
        return NUM
    try:
        qm = math.expm1(n * math.log1p(rate))
    except OverflowError:
        qm = math.inf
    if qm == 0:
        return NUM
    if due:
        qj = growth(rate, per - 2)
        if qj is None:
            return NUM
        ipmt = -rate * (_balance(pv, fv, qn, qj, qn - 1) + pmt)
        return (ipmt, -rate * (pv + fv) * math.exp((per - 2) * math.log1p(rate)) / qm)
    qj = growth(rate, per - 1)
    if qj is None:
        return NUM
    ipmt = -rate * _balance(pv, fv, qn, qj, qn - 1)
    ppmt = -rate * (pv + fv) * math.exp((per - 1) * math.log1p(rate)) / qm
    return (ipmt, ppmt)


@function(
    "IPMT",
    "number",
    "number",
    "number",
    "number",
    "number",
    "number",
    min_args=4,
    category=C,
    ulps=4,
    ulps_reason=_LIBM,
)
def IPMT(
    ctx: Context, rate: float, per: float, n: float, pv: float, fv: float = 0.0, due: float = 0.0
) -> Any:
    """Interest part of the payment of period ``per``."""
    parts = _parts(rate, per, n, pv, fv, due != 0)
    return parts if isinstance(parts, ErrorValue) else _result(parts[0])


@function(
    "PPMT",
    "number",
    "number",
    "number",
    "number",
    "number",
    "number",
    min_args=4,
    category=C,
    ulps=4,
    ulps_reason=_LIBM,
)
def PPMT(
    ctx: Context, rate: float, per: float, n: float, pv: float, fv: float = 0.0, due: float = 0.0
) -> Any:
    """Principal part of the payment of period ``per``."""
    parts = _parts(rate, per, n, pv, fv, due != 0)
    if isinstance(parts, ErrorValue):
        return parts
    return parts[1] if isinstance(parts[1], ErrorValue) else _result(parts[1])


@function("NPV", "number", "raw", variadic=True, category=C, ulps=4, ulps_reason=_NPV_SUM)
def NPV(ctx: Context, rate: float, *values: Any) -> Any:
    """Net present value of flows at the end of periods 1, 2, ... (numbers as SUM takes them)."""
    total = 0.0
    flows = tuple(BLANK if v is MISSING else v for v in values)
    for k, v in enumerate(numbers(flows, ctx), 1):
        if isinstance(v, ErrorValue):
            return v
        q = growth(rate, float(k))
        if q is None or q == 0:
            return DIV0
        total += v / q
    return _result(total)


def _flows(arg: Any, ctx: Context) -> list[Any] | ErrorValue:
    """The cells of a range or array, row by row, as values."""
    if isinstance(arg, Ref):
        arg = to_array(arg)
    if isinstance(arg, tuple):
        return [v for row in arg for v in row]
    if isinstance(arg, ErrorValue):
        return arg
    return [arg]


@function("XNPV", "value", "raw", "raw", category=C)
def XNPV(ctx: Context, rate: Any, values: Any, dates: Any) -> Any:
    """Net present value of flows on dates, discounted over 365-day years."""
    if isinstance(rate, bool):
        return VALUE
    rate = to_number(rate, ctx)
    if isinstance(rate, ErrorValue):
        return rate
    vs, ds = _flows(values, ctx), _flows(dates, ctx)
    if isinstance(vs, ErrorValue):
        return vs
    if isinstance(ds, ErrorValue):
        return ds
    if len(vs) != len(ds):
        return NUM
    nums: list[float] = []
    for v in vs:
        if isinstance(v, ErrorValue):
            return v
        if not isinstance(v, float) or isinstance(v, bool):
            return NUM
        nums.append(v)
    days: list[float] = []
    for d in ds:
        if isinstance(d, ErrorValue):
            return d
        if not isinstance(d, float) or isinstance(d, bool) or d < 0:
            return NUM
        days.append(float(math.trunc(d)))
    if rate <= 0:
        return NUM
    first = days[0]
    total = 0.0
    for v, d in zip(nums, days, strict=True):
        if d < first:
            return NUM
        q = growth(rate, (d - first) / 365)
        if q is None or q == 0:
            return NUM
        total += v / q
    return _result(total)


@function("SLN", "number", "number", "number", category=C)
def SLN(ctx: Context, cost: float, salvage: float, life: float) -> Any:
    """Straight-line depreciation per period."""
    if life == 0:
        return DIV0
    return _result((cost - salvage) / life)


def _declining(
    book: float, salvage: float, life: float, t0: float, t1: float, factor: float, switch: bool
) -> float:
    """Depreciation from time ``t0`` to ``t1`` of an asset worth ``book`` at ``t0``.

    Period by period from ``t0``: the declining balance at ``factor / life``, or straight line
    over the remaining life once that is larger (unless ``switch`` is off), never below the
    salvage value, the last period taken by its fraction after that limit.
    """
    rate = factor / life
    total, t, straight = 0.0, t0, False
    while t1 - t > 1e-12:
        step = min(1.0, t1 - t)
        full = book * rate
        if switch:
            remaining = life - t
            line = (book - salvage) / remaining if remaining > 0 else book - salvage
            if straight or line > full:
                straight, full = True, line
        full = min(full, max(book - salvage, 0.0))
        dep = full * step
        total += dep
        book -= dep
        t += step
    return total


@function(
    "VDB", "number", "number", "number", "number", "number", "number", "logical",
    min_args=5, category=C, ulps=4, ulps_reason=_VDB_SUM,
)  # fmt: skip
def VDB(
    ctx: Context,
    cost: float,
    salvage: float,
    life: float,
    start: float,
    end: float,
    factor: float = 2.0,
    no_switch: bool = False,
) -> Any:
    """Declining-balance depreciation between two periods, switching to straight line.

    Excel depreciates from ``start`` an asset worth what is left at ``start`` (capture
    ``finance``); a fractional last period is that fraction of the period's depreciation.
    """
    if start < 0 or end < start or end > life or cost < 0 or salvage < 0 or factor < 0:
        return NUM
    if life <= 0:
        return NUM
    before = _declining(cost, salvage, life, 0.0, start, factor, not no_switch)
    return _result(_declining(cost - before, salvage, life, start, end, factor, not no_switch))


# Iterative functions -------------------------------------------------------------------------
#
# Excel finds IRR, XIRR and RATE by Newton's method from a guess (0.1 by default). Its iterates
# are not reproduced bit for bit (they rest on its own NPV, whose last bits differ, §8.1), so
# the method below is the one that agrees best with the captured outcomes: Newton's method on
# 1 + rate, halving instead of stepping to 1 + rate <= 0, stopping when a step is below 1E-7,
# #NUM! when it does not settle in time.


def _newton(fn: Any, guess: float, steps: int, tol: float = 1e-7) -> float | ErrorValue:
    """A root of ``fn(t) -> (f, f')`` near ``1 + guess``, as a rate (``t - 1``)."""
    t = 1.0 + guess
    for _ in range(steps):
        if t <= 0:
            return NUM
        try:
            f, d = fn(t)
            if d == 0 or not math.isfinite(f) or not math.isfinite(d):
                return NUM
            nt = t - f / d
        except (ZeroDivisionError, OverflowError, ValueError):
            return NUM
        if nt <= 0:
            nt = t / 2
        if not math.isfinite(nt):
            return NUM
        done = abs(nt - t) < tol
        t = nt
        if done:
            return _result(t - 1.0) if t > 0 else NUM
    return NUM


def _irr_npv(flows: list[float]) -> Any:
    def fn(t: float) -> tuple[float, float]:
        f = d = 0.0
        for i, v in enumerate(flows):
            q = t**i
            f += v / q
            d -= i * v / (q * t)
        return f, d

    return fn


@function("IRR", "raw", "value", min_args=1, category=C, rel=1e-9, rel_reason=_ITERATION)
def IRR(ctx: Context, values: Any, guess: Any = 0.1) -> Any:
    """Internal rate of return of flows at regular periods (numbers in ranges only)."""
    g = _guess(guess, ctx)
    if isinstance(g, ErrorValue):
        return g
    flows: list[float] = []
    for v in numbers((values,), ctx, direct_text="skip"):
        if isinstance(v, ErrorValue):
            return VALUE
        flows.append(v)
    if g <= -1:
        return VALUE
    return _newton(_irr_npv(flows), g, 20)


def _guess(guess: Any, ctx: Context) -> float | ErrorValue:
    if guess is BLANK or guess is MISSING:
        return 0.1
    return to_number(guess, ctx)


@function("XIRR", "raw", "raw", "value", min_args=2, category=C, rel=1e-9, rel_reason=_ITERATION)
def XIRR(ctx: Context, values: Any, dates: Any, guess: Any = MISSING) -> Any:
    """Internal rate of return of flows on dates, over 365-day years.

    Unlike XNPV, dates after the first may come in any order and text dates are read; a
    blank guess is 0 and a logical one ``#VALUE!``.
    """
    if isinstance(guess, bool):
        return VALUE
    g = 0.1 if guess is MISSING else (0.0 if guess is BLANK else to_number(guess, ctx))
    if isinstance(g, ErrorValue):
        return g
    if (isinstance(values, Ref) and values.is_cell()) or not isinstance(values, Ref | tuple):
        return ErrorValue.NA
    vs, ds = _flows(values, ctx), _flows(dates, ctx)
    if isinstance(vs, ErrorValue):
        return vs
    if isinstance(ds, ErrorValue):
        return ds
    if len(vs) != len(ds):
        return NUM
    nums: list[float] = []
    for v in vs:
        if isinstance(v, ErrorValue):
            return v
        if not isinstance(v, float) or isinstance(v, bool):
            return VALUE
        nums.append(v)
    days: list[float] = []
    for d in ds:
        if isinstance(d, ErrorValue):
            return d
        n = to_number(d, ctx) if isinstance(d, str) else d
        if isinstance(n, ErrorValue):
            return n
        if not isinstance(n, float) or isinstance(n, bool) or n < 0:
            return NUM
        days.append(float(math.trunc(n)))
    if g < 0:
        return NUM
    years = [(d - days[0]) / 365 for d in days]

    def fn(t: float) -> tuple[float, float]:
        f = dd = 0.0
        for v, y in zip(nums, years, strict=True):
            q = t**y
            f += v / q
            dd -= y * v / (q * t)
        return f, dd

    return _newton(fn, g, 100)


@function(
    "RATE", "number", "number", "number", "number", "number", "value",
    min_args=3, category=C, rel=1e-9, rel_reason=_ITERATION,
)  # fmt: skip
def RATE(
    ctx: Context,
    n: float,
    pmt: float,
    pv: float,
    fv: float = 0.0,
    due: float = 0.0,
    guess: Any = 0.1,
) -> Any:
    """Interest rate per period of an annuity."""
    g = _guess(guess, ctx)
    if isinstance(g, ErrorValue):
        return g
    if g <= -1:
        return VALUE
    if n <= 0:
        return NUM
    k = 1.0 if due != 0 else 0.0

    def fn(t: float) -> tuple[float, float]:
        # the present-value form, flat for large rates: pv + pmt*(1+r*k)*(1-t**-n)/r + fv*t**-n
        r = t - 1.0
        u = t**-n
        du = -n * u / t
        if r == 0:
            return pv + pmt * n + fv, -pmt * n * (n + 1) / 2 + pmt * k * n - fv * n
        a = (1 - u) / r
        da = (-du * r - (1 - u)) / (r * r)
        return pv + pmt * (1 + r * k) * a + fv * u, pmt * (k * a + (1 + r * k) * da) + fv * du

    return _newton(fn, g, 20)
