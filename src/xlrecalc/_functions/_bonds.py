"""Day counts and the bond functions built on them.

DAYS360, YEARFRAC (date and time); COUPPCD, COUPNCD, COUPNUM, COUPDAYBS, COUPDAYS,
COUPDAYSNC, PRICE, YIELD, DURATION, MDURATION (financial). Capture ``bonds``.

Day-count bases: 0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360.
Coupon dates step back from maturity by ``12 / frequency`` months, keeping the maturity's
day of month (the last day of the month when maturity is one).
"""

from __future__ import annotations

import calendar
import math
from typing import Any, cast

from xlrecalc._coerce import Context, to_number
from xlrecalc._dates import date_serial, serial_parts
from xlrecalc._functions._registry import function
from xlrecalc._values import ErrorValue

NUM, VALUE = ErrorValue.NUM, ErrorValue.VALUE
D = "date and time"
C = "financial"
_PRICE_SUM = (
    "Excel's summation of the discounted coupons is not reproduced in the last bits (two "
    "orders tried); within 8 units in the last place on every captured row (7 at quarterly "
    "coupons)"
)
_YIELD_ITERATION = (
    "with more than one period left Excel iterates on PRICE, whose last bits differ, to its "
    "own stopping tolerance: within 1E-9 relative"
)
_DURATION_SUM = (
    "the weighted sums of discounted cash flows differ from Excel's in the last bits; within 4 "
    "units in the last place on every captured row"
)

Ymd = tuple[int, int, int]


def _ymd(x: float, ctx: Context) -> Ymd | None:
    if x < 0:
        return None
    return serial_parts(x, date1904=ctx.date1904, leap=ctx.quirks.leap_year_1900)


def _month_days(y: int, m: int, ctx: Context) -> int:
    if y == 1900 and m == 2 and ctx.quirks.leap_year_1900 and not ctx.date1904:
        return 29
    return calendar.monthrange(y, m)[1]


def _last_of_feb(p: Ymd, ctx: Context) -> bool:
    return p[1] == 2 and p[2] == _month_days(p[0], 2, ctx)


def _days360_us(a: Ymd, b: Ymd, ctx: Context) -> int:
    """``DAYS360`` with the US method, as the worksheet function computes it."""
    y1, m1, d1 = a
    y2, m2, d2 = b
    if d1 == 31 or _last_of_feb(a, ctx):
        d1 = 30
    if d2 == 31:
        if d1 < 30:
            d2 = 1
            m2 += 1
        else:
            d2 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def _days360_nasd(a: Ymd, b: Ymd, ctx: Context) -> int:
    """The US (NASD) 30/360 count of ``YEARFRAC`` basis 0 and the coupon functions."""
    y1, m1, d1 = a
    y2, m2, d2 = b
    feb1, feb2 = _last_of_feb(a, ctx), _last_of_feb(b, ctx)
    if d1 == 31 and d2 == 31:
        d1 = d2 = 30
    elif d1 == 31:
        d1 = 30
    elif d1 == 30 and d2 == 31:
        d2 = 30
    elif feb1 and feb2:
        d1 = d2 = 30
    elif feb1:
        d1 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def _days360_eu(a: Ymd, b: Ymd) -> int:
    y1, m1, d1 = a
    y2, m2, d2 = b
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (min(d2, 30) - min(d1, 30))


@function("DAYS360", "number", "number", "logical", min_args=2, category=D)
def DAYS360(ctx: Context, start: float, end: float, european: bool = False) -> Any:
    """Days between two dates in a 360-day year (US method unless ``european``)."""
    a, b = _ymd(start, ctx), _ymd(end, ctx)
    if a is None or b is None:
        return NUM
    return float(_days360_eu(a, b) if european else _days360_us(a, b, ctx))


def _isleap(y: int) -> bool:
    """Year lengths for actual/actual follow the calendar: 1900 has 365 days here."""
    return calendar.isleap(y)


def _serial(y: int, m: int, d: int, ctx: Context) -> int:
    s = date_serial(y, m, d, date1904=ctx.date1904, leap=ctx.quirks.leap_year_1900)
    assert s is not None
    return s


def _actual_actual(s1: int, s2: int, a: Ymd, b: Ymd, ctx: Context) -> float:
    y1, m1, d1 = a
    y2, m2, d2 = b
    if y1 == y2 or (y2 == y1 + 1 and (m1 > m2 or (m1 == m2 and d1 >= d2))):
        if y1 == y2:
            length = 366 if _isleap(y1) else 365
        else:
            feb29 = (_isleap(y1) and (m1, d1) <= (2, 29)) or (_isleap(y2) and (m2, d2) >= (2, 29))
            length = 366 if feb29 else 365
        return (s2 - s1) / length
    total = sum(366 if _isleap(y) else 365 for y in range(y1, y2 + 1))
    return (s2 - s1) / (total / (y2 - y1 + 1))


def _basis(basis: float) -> int | None:
    b = math.trunc(basis)
    return b if 0 <= b <= 4 else None


def yearfrac(start: float, end: float, basis: int, ctx: Context) -> float | ErrorValue:
    s1, s2 = math.trunc(start), math.trunc(end)
    if s1 > s2:
        s1, s2 = s2, s1
    a, b = _ymd(s1, ctx), _ymd(s2, ctx)
    if a is None or b is None:
        return NUM
    if basis == 0:
        return _days360_nasd(a, b, ctx) / 360
    if basis == 1:
        return _actual_actual(s1, s2, a, b, ctx)
    if basis == 2:
        return (s2 - s1) / 360
    if basis == 3:
        return (s2 - s1) / 365
    return _days360_eu(a, b) / 360


@function("YEARFRAC", "number", "number", "number", min_args=2, category=D)
def YEARFRAC(ctx: Context, start: float, end: float, basis: float = 0.0) -> Any:
    """The fraction of a year between two dates under a day-count basis."""
    b = _basis(basis)
    if b is None:
        return NUM
    return yearfrac(start, end, b, ctx)


# Coupons --------------------------------------------------------------------------------------


class _Bond:
    """Coupon dates around settlement for a bond, or the error its arguments give."""

    __slots__ = ("basis", "ctx", "freq", "mature", "n", "ncd", "pcd", "settle")

    def __init__(self, settle: int, mature: int, freq: int, basis: int, ctx: Context) -> None:
        self.settle, self.mature, self.freq, self.basis, self.ctx = (
            settle, mature, freq, basis, ctx,
        )  # fmt: skip
        my, mm, md = _ymd(mature, ctx) or (0, 0, 0)
        eom = md == _month_days(my, mm, ctx)
        step = 12 // freq
        k = 0
        while True:
            k += 1
            date = self._shift(my, mm, md, -k * step, eom)
            if date <= settle:
                break
        self.pcd = date
        self.ncd = self._shift(my, mm, md, -(k - 1) * step, eom)
        self.n = k

    def _shift(self, y: int, m: int, d: int, months: int, eom: bool) -> int:
        total = y * 12 + (m - 1) + months
        y, m0 = divmod(total, 12)
        last = _month_days(y, m0 + 1, self.ctx)
        return _serial(y, m0 + 1, last if eom else min(d, last), self.ctx)

    def _parts(self, serial: int) -> Ymd:
        p = _ymd(serial, self.ctx)
        assert p is not None  # coupon dates lie between valid dates
        return p

    def days(self) -> float:
        """COUPDAYS: the days of the coupon period holding settlement."""
        if self.basis == 1:  # the period from pcd, without the end-of-month rule
            y, m, d = _ymd(self.pcd, self.ctx) or (0, 0, 0)
            return float(self._shift(y, m, d, 12 // self.freq, False) - self.pcd)
        return (365.0 if self.basis == 3 else 360.0) / self.freq

    def daybs(self) -> float:
        """COUPDAYBS: days from the period's start to settlement."""
        return float(self._count(self.pcd, self.settle))

    def daysnc(self) -> float:
        """COUPDAYSNC: days from settlement to the next coupon."""
        if self.basis == 0:
            return self.days() - self.daybs()
        if self.basis == 4:
            return float(_days360_eu(self._parts(self.settle), self._parts(self.ncd)))
        return float(self.ncd - self.settle)

    def _count(self, s1: int, s2: int) -> int:
        a, b = self._parts(s1), self._parts(s2)
        if self.basis == 0:
            return _days360_nasd(a, b, self.ctx)
        if self.basis == 4:
            return _days360_eu(a, b)
        return s2 - s1


def _bond(settle: float, mature: float, freq: float, basis: float, ctx: Context) -> Any:
    s, m = math.trunc(settle), math.trunc(mature)
    f, b = math.trunc(freq), _basis(basis)
    if b is None or f not in (1, 2, 4) or s >= m or s < 0:
        return NUM
    if _ymd(s, ctx) is None or _ymd(m, ctx) is None:
        return NUM
    return _Bond(s, m, f, b, ctx)


def _coupon(name: str, method: Any) -> None:
    @function(name, "number", "number", "number", "number", min_args=3, category=C)
    def coup(ctx: Context, settle: float, mature: float, freq: float, basis: float = 0.0) -> Any:
        bond = _bond(settle, mature, freq, basis, ctx)
        return bond if isinstance(bond, ErrorValue) else method(bond)

    coup.__doc__ = method.__doc__


_coupon("COUPPCD", lambda b: float(b.pcd))
_coupon("COUPNCD", lambda b: float(b.ncd))
_coupon("COUPNUM", lambda b: float(b.n))
_coupon("COUPDAYS", _Bond.days)
_coupon("COUPDAYBS", _Bond.daybs)
_coupon("COUPDAYSNC", _Bond.daysnc)


# Prices, yields, durations ---------------------------------------------------------------------


def _price(b: _Bond, rate: float, yld: float, redemption: float) -> float:
    # Excel takes the days to the next coupon as COUPDAYS - COUPDAYBS here, in every basis
    e, a, n, f = b.days(), b.daybs(), b.n, b.freq
    dsc = e - a
    coupon = 100 * rate / f
    if n == 1:
        return (coupon + redemption) / (1 + dsc / e * yld / f) - coupon * a / e
    base = 1 + yld / f
    x = dsc / e
    total: float = redemption / base ** (n - 1 + x)
    for k in range(1, n + 1):
        total += coupon / base ** (k - 1 + x)
    return total - coupon * a / e


@function(
    "PRICE", "number", "number", "number", "number", "number", "number", "number",
    min_args=6, category=C, ulps=8, ulps_reason=_PRICE_SUM,
)  # fmt: skip
def PRICE(
    ctx: Context,
    settle: float,
    mature: float,
    rate: float,
    yld: float,
    redemption: float,
    freq: float,
    basis: float = 0.0,
) -> Any:
    """Price per 100 of face value of a bond paying periodic interest."""
    b = _bond(settle, mature, freq, basis, ctx)
    if isinstance(b, ErrorValue):
        return b
    if rate < 0 or yld < 0 or redemption <= 0:
        return NUM
    return _price(b, rate, yld, redemption)


@function(
    "YIELD", "number", "number", "number", "number", "number", "number", "number",
    min_args=6, category=C, rel=1e-9, rel_reason=_YIELD_ITERATION,
)  # fmt: skip
def YIELD(
    ctx: Context,
    settle: float,
    mature: float,
    rate: float,
    pr: float,
    redemption: float,
    freq: float,
    basis: float = 0.0,
) -> Any:
    """Yield of a bond paying periodic interest, from its price."""
    b = _bond(settle, mature, freq, basis, ctx)
    if isinstance(b, ErrorValue):
        return b
    if rate < 0 or pr <= 0 or redemption <= 0:
        return NUM
    f = b.freq
    if b.n == 1:
        # Excel's closed form for one period left (capture bonds): A is COUPDAYBS, E the
        # actual period in bases 1 to 3 (360/f in 0 and 4), DSR the basis's own count from
        # settlement to redemption.
        a = b.daybs()
        e = float(b.ncd - b.pcd) if b.basis in (1, 2, 3) else b.days()
        dsr = float(b._count(b.settle, b.mature))
        num = (redemption / 100 + rate / f) - (pr / 100 + a / e * rate / f)
        return num / (pr / 100 + a / e * rate / f) * f * e / dsr
    y = rate if rate > 0 else 0.1
    for _ in range(100):
        p = _price(b, rate, y, redemption) - pr
        h = 1e-7 * max(abs(y), 1e-7)
        dp = (_price(b, rate, y + h, redemption) - pr - p) / h
        if dp == 0:
            return NUM
        ny = y - p / dp
        if abs(ny - y) < 1e-12:
            return ny
        y = ny if ny > -f else (y - f) / 2
    return NUM


def _duration(b: _Bond, coupon: float, yld: float) -> float:
    e, f, n = b.days(), b.freq, b.n
    x = (e - b.daybs()) / e  # the fraction of a period to the next coupon, as in PRICE
    c = coupon * 100 / f
    base = yld / f + 1
    dur: float = 0.0
    price: float = 0.0
    for k in range(1, n + 1):
        t = k - 1 + x
        v: float = c / base**t
        dur += t * v
        price += v
    t = n - 1 + x
    v = 100 / base**t
    dur += t * v
    price += v
    return dur / price / f


def _durations(name: str, modified: bool) -> None:
    @function(
        name, "number", "number", "number", "number", "number", "number",
        min_args=5, category=C, ulps=4, ulps_reason=_DURATION_SUM,
    )  # fmt: skip
    def dur(
        ctx: Context,
        settle: float,
        mature: float,
        coupon: float,
        yld: float,
        freq: float,
        basis: float = 0.0,
    ) -> Any:
        b = _bond(settle, mature, freq, basis, ctx)
        if isinstance(b, ErrorValue):
            return b
        if coupon < 0 or yld < 0:
            return NUM
        d = _duration(b, coupon, yld)
        return d / (1 + yld / b.freq) if modified else d

    dur.__doc__ = (
        "Modified Macaulay duration of a bond." if modified else "Macaulay duration of a bond."
    )


_durations("DURATION", False)
_durations("MDURATION", True)


# Discounted securities (capture finance_family) ------------------------------------------------


def _on_calendar(v: Any, ctx: Context) -> float | ErrorValue:
    """A date as a whole day: an argument's problems come before the next argument's."""
    x = _analysis(v, ctx)
    if isinstance(x, ErrorValue):
        return x
    if serial_parts(x, date1904=ctx.date1904, leap=ctx.quirks.leap_year_1900) is None:
        return NUM
    return float(math.trunc(x))


def _analysis(v: Any, ctx: Context) -> float | ErrorValue:
    """A number as the old Analysis ToolPak functions take one: a logical is ``#VALUE!``."""
    return VALUE if isinstance(v, bool) else to_number(v, ctx)


def _discount(name: str, doc: str, compute: Any) -> None:
    @function(
        name, "value", "value", "value", "value", "value", min_args=4, errors=True, category=C,
        ulps=4, ulps_reason=_DISCOUNT,
    )  # fmt: skip
    def security(ctx: Context, settle: Any, mature: Any, a: Any, b: Any, basis: Any = 0.0) -> Any:
        values = [
            _on_calendar(settle, ctx),
            _on_calendar(mature, ctx),
            _analysis(a, ctx),
            _analysis(b, ctx),
            _analysis(basis, ctx),
        ]
        for v in values:
            if isinstance(v, ErrorValue):
                return v
        start, end, x, y, code = cast(list[float], values)
        kind = _basis(code)
        if kind is None or start >= end or x <= 0 or y <= 0:
            return NUM
        fraction = yearfrac(start, end, kind, ctx)
        return fraction if isinstance(fraction, ErrorValue) else compute(x, y, fraction)

    security.__doc__ = doc


def _positive(x: float) -> float | ErrorValue:
    return x if x > 0 and math.isfinite(x) else NUM


_DISCOUNT = "Excel's day fractions differ in the last bits; within 4 units in the last place"
_discount("DISC", "The discount rate of a security.", lambda p, r, t: (1 - p / r) / t)
_discount("INTRATE", "The interest rate of a security.", lambda i, r, t: (r - i) / i / t)
_discount(
    "RECEIVED", "The amount received at maturity.", lambda i, d, t: _positive(i / (1 - d * t))
)
_discount("PRICEDISC", "The price of a discounted security.", lambda d, r, t: r * (1 - d * t))
_discount("YIELDDISC", "The yield of a discounted security.", lambda p, r, t: (r - p) / p / t)


def _treasury(ctx: Context, settle: Any, mature: Any, x: Any) -> Any:
    values = [_on_calendar(settle, ctx), _on_calendar(mature, ctx), _analysis(x, ctx)]
    for v in values:
        if isinstance(v, ErrorValue):
            return v
    start, end, value = cast(list[float], values)
    days = end - start
    return NUM if value <= 0 or days <= 0 or days > 366 else (value, days)


@function(
    "TBILLPRICE", "value", "value", "value", errors=True, category=C, ulps=4, ulps_reason=_DISCOUNT
)
def TBILLPRICE(ctx: Context, settle: Any, mature: Any, discount: Any) -> Any:
    """The price of a Treasury bill."""
    t = _treasury(ctx, settle, mature, discount)
    if isinstance(t, ErrorValue):
        return t
    return _positive(100 * (1 - t[0] * t[1] / 360))


@function(
    "TBILLYIELD", "value", "value", "value", errors=True, category=C, ulps=4, ulps_reason=_DISCOUNT
)
def TBILLYIELD(ctx: Context, settle: Any, mature: Any, price: Any) -> Any:
    """The yield of a Treasury bill."""
    t = _treasury(ctx, settle, mature, price)
    if isinstance(t, ErrorValue):
        return t
    return (100 - t[0]) / t[0] * 360 / t[1]


@function(
    "TBILLEQ", "value", "value", "value", errors=True, category=C, ulps=4, ulps_reason=_DISCOUNT
)
def TBILLEQ(ctx: Context, settle: Any, mature: Any, discount: Any) -> Any:
    """The bond-equivalent yield of a Treasury bill."""
    t = _treasury(ctx, settle, mature, discount)
    if isinstance(t, ErrorValue):
        return t
    d, days = t
    price = 1 - d * days / 360
    if price <= 0:
        return NUM
    if days <= 182:
        return 365 * d / (360 - d * days)
    span = days / (366 if days == 366 else 365)  # capture finance_family
    return (-span + math.sqrt(span * span - (2 * span - 1) * (1 - 1 / price))) / (span - 0.5)
