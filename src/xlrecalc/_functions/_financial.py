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
from typing import Any, cast

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


# Closed forms (capture finance_family) ------------------------------------------------------

_WALK_PERIODS = 1024
_POWERS = "Excel's powers and products differ in the last bits; within 4 units in the last place"


def _strict(v: Any, ctx: Context) -> float | ErrorValue:
    """A number as the old Analysis ToolPak functions take one: a logical is ``#VALUE!``."""
    return VALUE if isinstance(v, bool) else to_number(v, ctx)


def _power(base: float, exponent: float) -> float:
    """A power, infinite where Python would raise; ``_result`` makes that ``#NUM!``."""
    try:
        result = base**exponent
    except (OverflowError, ZeroDivisionError):
        return math.inf
    return result if isinstance(result, float) else math.nan


def _long_span(
    r: float, n: float, pv: float, start: int, end: int, due: int, interest: bool
) -> float | ErrorValue:
    """Interest or principal over many periods, in closed form as Excel's speed implies.

    Balances go through discount factors, which stay finite where ``(1 + r) ** n`` would not.
    """
    v = 1 / (1 + r)
    whole = 1 - _power(v, n)
    payment = pv * r / whole * (v if due else 1.0)

    def balance(k: int) -> float:
        if k <= 0:
            return pv
        return pv * (v if due else 1.0) * (1 - _power(v, n - k)) / whole

    principal = balance(start - 1) - balance(end)
    return _result(principal - payment * (end - start + 1) if interest else -principal)


def _cumulative(
    ctx: Context, rate: Any, n: Any, pv: Any, first: Any, last: Any, due: Any, part: Any
) -> float | ErrorValue:
    values = [_strict(v, ctx) for v in (rate, n, pv, first, last, due)]
    for v in values:
        if isinstance(v, ErrorValue):
            return v
    r, nper, present, a, b, kind = cast(list[float], values)
    start, end, code = math.ceil(a), math.floor(b), math.trunc(kind)
    if r <= 0 or nper <= 0 or present <= 0 or start < 1 or end < start or end > nper:
        return NUM
    if code not in (0, 1):
        return NUM
    if nper >= _WALK_PERIODS:  # a formula must not be able to ask for 1E300 steps
        return _long_span(r, nper, present, start, end, code, part is IPMT)
    total = 0.0
    for period in range(start, end + 1):
        step = part(ctx, r, float(period), nper, present, 0.0, float(code))
        if isinstance(step, ErrorValue):
            return step
        total += step
    return _result(total)


@function(
    "CUMIPMT", "value", "value", "value", "value", "value", "value",
    errors=True, category=C, ulps=8, ulps_reason=_POWERS,
)  # fmt: skip
def CUMIPMT(ctx: Context, rate: Any, n: Any, pv: Any, a: Any, b: Any, due: Any) -> Any:
    """Interest paid between two periods of a loan."""
    return _cumulative(ctx, rate, n, pv, a, b, due, IPMT)


@function(
    "CUMPRINC", "value", "value", "value", "value", "value", "value",
    errors=True, category=C, ulps=8, ulps_reason=_POWERS,
)  # fmt: skip
def CUMPRINC(ctx: Context, rate: Any, n: Any, pv: Any, a: Any, b: Any, due: Any) -> Any:
    """Principal paid between two periods of a loan."""
    return _cumulative(ctx, rate, n, pv, a, b, due, PPMT)


@function("SYD", "number", "number", "number", "number", category=C)
def SYD(
    ctx: Context, cost: float, salvage: float, life: float, period: float
) -> float | ErrorValue:
    """Depreciation by the sum of the years' digits."""
    if life <= 0 or period <= 0 or period > life or salvage < 0:
        return NUM
    return _result((cost - salvage) * (life - period + 1) * 2 / (life * (life + 1)))


@function(
    "DDB", "number", "number", "number", "number", "number",
    min_args=4, category=C, ulps=8, ulps_reason=_POWERS,
)  # fmt: skip
def DDB(
    ctx: Context, cost: float, salvage: float, life: float, period: float, factor: float = 2.0
) -> float | ErrorValue:
    """Depreciation by a declining balance."""
    if cost < 0 or salvage < 0 or life <= 0 or period <= 0 or period > life or factor <= 0:
        return NUM
    rate = min(factor / life, 1.0)
    left = cost * _power(1 - rate, period - 1)
    return _result(max(min(left * rate, left - salvage), 0.0))


@function(
    "DB", "number", "number", "number", "number", "number",
    min_args=4, category=C, ulps=8, ulps_reason=_POWERS,
)  # fmt: skip
def DB(
    ctx: Context, cost: float, salvage: float, life: float, period: float, month: float = 12.0
) -> Any:
    """Depreciation by a fixed declining balance, whose rate has three decimals."""
    m = math.trunc(month)
    per = math.trunc(period)
    if cost < 0 or salvage < 0 or life <= 0 or per < 1 or m < 1 or m > 12 or life < 1:
        return NUM
    if per > life + (m < 12):
        return NUM
    if cost == 0:
        return 0.0
    rate = round(1 - _power(salvage / cost, 1 / life), 3)
    if per > _WALK_PERIODS:  # the balance only shrinks: past the first year in closed form
        return _result((cost - cost * rate * m / 12) * _power(1 - rate, per - 2) * rate)
    taken = cost * rate * m / 12
    if per == 1:
        return taken
    total = taken
    step = taken
    for k in range(2, per + 1):
        step = (cost - total) * rate * ((12 - m) / 12 if k == life + 1 else 1)
        total += step
    return step


def _dollar_scale(fraction: float) -> tuple[int, float] | ErrorValue:
    f = math.trunc(fraction)
    if fraction < 0:
        return NUM
    if f == 0:
        return DIV0
    return f, _power(10.0, math.ceil(math.log10(f)))


@function("DOLLARDE", "value", "value", errors=True, category=C)
def DOLLARDE(ctx: Context, price: Any, fraction: Any) -> float | ErrorValue:
    """A price in the form 1.02 meaning 1 and 2/16, as a decimal."""
    x, f = _strict(price, ctx), _strict(fraction, ctx)
    if isinstance(x, ErrorValue):
        return x
    if isinstance(f, ErrorValue):
        return f
    scale = _dollar_scale(f)
    if isinstance(scale, ErrorValue):
        return scale
    whole = math.trunc(x)
    return whole + (x - whole) * scale[1] / scale[0]


@function("DOLLARFR", "value", "value", errors=True, category=C)
def DOLLARFR(ctx: Context, price: Any, fraction: Any) -> float | ErrorValue:
    """A decimal price in the form 1.02 meaning 1 and 2/16."""
    x, f = _strict(price, ctx), _strict(fraction, ctx)
    if isinstance(x, ErrorValue):
        return x
    if isinstance(f, ErrorValue):
        return f
    scale = _dollar_scale(f)
    if isinstance(scale, ErrorValue):
        return scale
    whole = math.trunc(x)
    return whole + (x - whole) * scale[0] / scale[1]


@function("EFFECT", "value", "value", errors=True, category=C, ulps=4, ulps_reason=_POWERS)
def EFFECT(ctx: Context, rate: Any, periods: Any) -> float | ErrorValue:
    """The effective annual rate of a nominal one."""
    r, k = _strict(rate, ctx), _strict(periods, ctx)
    if isinstance(r, ErrorValue):
        return r
    if isinstance(k, ErrorValue):
        return k
    n = math.trunc(k)
    if r <= 0 or n < 1:
        return NUM
    grown = growth(r / n, float(n))  # as Excel takes a whole power (capture finance)
    return NUM if grown is None else _result(grown - 1)


@function("NOMINAL", "value", "value", errors=True, category=C, ulps=4, ulps_reason=_POWERS)
def NOMINAL(ctx: Context, rate: Any, periods: Any) -> float | ErrorValue:
    """The nominal annual rate of an effective one."""
    r, k = _strict(rate, ctx), _strict(periods, ctx)
    if isinstance(r, ErrorValue):
        return r
    if isinstance(k, ErrorValue):
        return k
    n = math.trunc(k)
    if r <= 0 or n < 1:
        return NUM
    return _result((_power(1 + r, 1 / n) - 1) * n)


@function("FVSCHEDULE", "value", "raw", errors=True, category=C)
def FVSCHEDULE(ctx: Context, present: Any, schedule: Any) -> float | ErrorValue:
    """A present value grown by each rate of a schedule in turn."""
    total = _strict(present, ctx)
    if isinstance(total, ErrorValue):
        return total
    if isinstance(schedule, Ref):
        rates: Any = [v for row in to_array(schedule) for v in row]
    elif isinstance(schedule, tuple):
        rates = [v for row in schedule for v in row]
    else:
        rates = [schedule]
    for v in rates:
        rate = 0.0 if v is BLANK else _strict(v, ctx)
        if isinstance(rate, ErrorValue):
            return rate
        total *= 1 + rate
    return _result(total)


@function("ISPMT", "number", "number", "number", "number", category=C, ulps=4, ulps_reason=_POWERS)
def ISPMT(ctx: Context, rate: float, period: float, n: float, pv: float) -> float | ErrorValue:
    """Interest of a period of a loan repaid in equal parts of its principal."""
    return DIV0 if n == 0 else _result(pv * rate * (period - n) / n)


@function("MIRR", "raw", "number", "number", category=C, ulps=8, ulps_reason=_POWERS)
def MIRR(ctx: Context, flows: Any, finance: float, reinvest: float) -> float | ErrorValue:
    """The modified internal rate of return of a series of cash flows."""
    if isinstance(flows, Ref):
        values: Any = flows.values()
    elif isinstance(flows, tuple):
        values = [v for row in flows for v in row]
    else:
        return DIV0
    numbers_ = []
    for v in values:
        if isinstance(v, ErrorValue):
            return v
        if type(v) is float:
            numbers_.append(v)
    n = len(numbers_)
    if n < 2 or reinvest == -1:
        return DIV0
    if finance == -1 and any(v < 0 for v in numbers_[1:]):
        return DIV0
    costs = _plain_sum(v / _power(1 + finance, i) for i, v in enumerate(numbers_) if v < 0)
    gains = _plain_sum(v * _power(1 + reinvest, n - 1 - i) for i, v in enumerate(numbers_) if v > 0)
    if not costs:
        return DIV0
    return -1.0 if not gains else _result(_power(-gains / costs, 1 / (n - 1)) - 1)


def _plain_sum(values: Any) -> float:
    total = 0.0
    for v in values:
        total += v
    return total


@function("_xlfn.PDURATION", "number", "number", "number", category=C, ulps=4, ulps_reason=_POWERS)
def PDURATION(ctx: Context, rate: float, present: float, future: float) -> float | ErrorValue:
    """Periods for an investment to reach a value."""
    if rate <= 0 or present <= 0 or future <= 0:
        return NUM
    return (math.log(future) - math.log(present)) / math.log(1 + rate)


@function("_xlfn.RRI", "number", "number", "number", category=C, ulps=4, ulps_reason=_POWERS)
def RRI(ctx: Context, n: float, present: float, future: float) -> float | ErrorValue:
    """The rate for an investment to grow to a value."""
    if n <= 0:
        return NUM
    if present == 0:
        return 0.0 if future == 0 else NUM
    ratio = future / present
    return NUM if ratio < 0 else _result(_power(ratio, 1 / n) - 1)
