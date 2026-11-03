"""Date and time functions.

DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, EDATE, TODAY, NOW, DATEVALUE, TIMEVALUE, WORKDAY,
NETWORKDAYS and their .INTL forms, HOUR, MINUTE, SECOND.
"""

from __future__ import annotations

import calendar
import datetime as dt
import math
from typing import Any

from xlrecalc._coerce import Context, scalar, to_number
from xlrecalc._dates import MAX_SERIAL_1900, MAX_SERIAL_1904, date_serial, serial_parts, weekday
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._numtext import text_to_parts
from xlrecalc._values import BLANK, ErrorValue, Ref

NUM, VALUE = ErrorValue.NUM, ErrorValue.VALUE
D = "date and time"


def _leap(ctx: Context) -> bool:
    return ctx.quirks.leap_year_1900


def _serial(y: float, m: float, d: float, ctx: Context) -> float | ErrorValue:
    s = date_serial(y, m, d, date1904=ctx.date1904, leap=_leap(ctx))
    return NUM if s is None else float(s)


def _parts(x: float, ctx: Context) -> tuple[int, int, int] | ErrorValue:
    p = serial_parts(x, date1904=ctx.date1904, leap=_leap(ctx))
    return NUM if p is None else p


@function("DATE", "number", "number", "number", category=D)
def DATE(ctx: Context, y: float, m: float, d: float) -> float | ErrorValue:
    """The serial of a date; months and days may overflow (§4.3)."""
    return _serial(y, m, d, ctx)


@function("YEAR", "number", category=D)
def YEAR(ctx: Context, x: float) -> float | ErrorValue:
    """Year of a serial."""
    p = _parts(x, ctx)
    return p if isinstance(p, ErrorValue) else float(p[0])


@function("MONTH", "number", category=D)
def MONTH(ctx: Context, x: float) -> float | ErrorValue:
    """Month of a serial."""
    p = _parts(x, ctx)
    return p if isinstance(p, ErrorValue) else float(p[1])


@function("DAY", "number", category=D)
def DAY(ctx: Context, x: float) -> float | ErrorValue:
    """Day of a serial."""
    p = _parts(x, ctx)
    return p if isinstance(p, ErrorValue) else float(p[2])


_WEEK_STARTS = {1: 1, 2: 2, 3: 2, 11: 2, 12: 3, 13: 4, 14: 5, 15: 6, 16: 7, 17: 1}


@function("WEEKDAY", "number", "number", min_args=1, category=D)
def WEEKDAY(ctx: Context, x: float, kind: float = 1.0) -> float | ErrorValue:
    """Day of the week: type 1 Sunday=1, 2 Monday=1, 3 Monday=0, 11-17 other first days."""
    w = weekday(x, date1904=ctx.date1904, leap=_leap(ctx))
    if w is None:
        return NUM
    k = math.trunc(kind)
    start = _WEEK_STARTS.get(k)
    if start is None:
        return NUM
    n = (w - start) % 7 + 1
    return float(n - 1 if k == 3 else n)


def _month_days(y: int, m: int, ctx: Context) -> int:
    """Days in a month by the real calendar.

    EOMONTH and EDATE end February 1900 on the 28th (serial 59), not on Excel's fictitious
    29th (captured).
    """
    return calendar.monthrange(y, m)[1]


def _shift(start: float, months: float, ctx: Context) -> tuple[int, int, int] | ErrorValue:
    p = _parts(start, ctx)
    if isinstance(p, ErrorValue):
        return p
    y, m, d = p
    total = y * 12 + (m - 1) + math.trunc(months)
    y, m = divmod(total, 12)
    if not 1900 <= y <= 9999:
        return NUM
    return y, m + 1, d


@function("EDATE", "number", "number", category=D)
def EDATE(ctx: Context, start: float, months: float) -> float | ErrorValue:
    """The same day ``months`` later, at most the month's last day."""
    p = _shift(start, months, ctx)
    if isinstance(p, ErrorValue):
        return p
    y, m, d = p
    return _serial(y, m, min(max(d, 1), _month_days(y, m, ctx)), ctx)


@function("EOMONTH", "number", "number", category=D)
def EOMONTH(ctx: Context, start: float, months: float) -> float | ErrorValue:
    """The last day of the month ``months`` later."""
    p = _shift(start, months, ctx)
    if isinstance(p, ErrorValue):
        return p
    y, m, _ = p
    return _serial(y, m, _month_days(y, m, ctx), ctx)


def _now_serial(ctx: Context, now: dt.datetime) -> float:
    days = date_serial(now.year, now.month, now.day, date1904=ctx.date1904, leap=_leap(ctx))
    seconds = now.hour * 3600 + now.minute * 60 + now.second + now.microsecond / 1e6
    return float(days or 0) + seconds / 86400


@function("TODAY", volatile=True, category=D)
def TODAY(ctx: Context) -> float:
    """Today's serial."""
    return float(math.floor(_now_serial(ctx, dt.datetime.now())))


@function("NOW", volatile=True, category=D)
def NOW(ctx: Context) -> float:
    """The serial of this moment, in local time."""
    return _now_serial(ctx, dt.datetime.now())


def _text_date_time(text: Any, ctx: Context) -> float | ErrorValue:
    if not isinstance(text, str):
        return VALUE
    parsed = text_to_parts(text, ctx.quirks, date1904=ctx.date1904, today=ctx.today)
    if parsed is None or parsed.number is not None:
        return VALUE
    return parsed.value


@function("DATEVALUE", "value", category=D)
def DATEVALUE(ctx: Context, text: Any) -> float | ErrorValue:
    """The date in a text, as a serial."""
    v = _text_date_time(text, ctx)
    return v if isinstance(v, ErrorValue) else float(math.floor(v))


@function("TIMEVALUE", "value", category=D)
def TIMEVALUE(ctx: Context, text: Any) -> float | ErrorValue:
    """The time in a text, as a fraction of a day."""
    v = _text_date_time(text, ctx)
    return v if isinstance(v, ErrorValue) else v - math.floor(v)


# Working days (capture ``workdays``) ------------------------------------------------------------


def _limit(ctx: Context) -> int:
    return MAX_SERIAL_1904 if ctx.date1904 else MAX_SERIAL_1900


def _weekend(code: Any, ctx: Context) -> frozenset[int] | ErrorValue:
    """The weekend days (0 Monday to 6 Sunday) a WORKDAY.INTL weekend argument names."""
    if code is MISSING:
        return frozenset({5, 6})
    if isinstance(code, str):
        if len(code) != 7 or set(code) - {"0", "1"}:
            return VALUE
        return frozenset(i for i, ch in enumerate(code) if ch == "1")
    n = to_number(code, ctx)
    if isinstance(n, ErrorValue):
        return n
    k = math.trunc(n)
    if 1 <= k <= 7:  # 1: Saturday and Sunday, 2: Sunday and Monday, ...
        return frozenset({(k + 4) % 7, (k + 5) % 7})
    if 11 <= k <= 17:  # 11: Sunday only, 12: Monday only, ...
        return frozenset({(k - 12) % 7})
    return NUM


def _code(weekend: Any, ctx: Context) -> Any:
    return weekend if weekend is MISSING else scalar(weekend, ctx)


def _holidays(arg: Any, ctx: Context) -> frozenset[int] | ErrorValue:
    if arg is MISSING or arg is BLANK:
        return frozenset()
    if isinstance(arg, Ref):
        values = [v for _, _, v in arg.cells()]
    elif isinstance(arg, tuple):
        values = [v for row in arg for v in row]
    else:
        values = [arg]
    days: set[int] = set()
    for v in values:
        if isinstance(v, ErrorValue):
            return v
        if v is BLANK:
            continue
        n = _day_number(v, ctx)
        if isinstance(n, ErrorValue):
            return n
        if n < 0 or n > _limit(ctx):
            return NUM
        days.add(math.trunc(n))
    return frozenset(days)


def _day_of_week(serial: int, ctx: Context) -> int:
    """0 for Monday to 6 for Sunday, Excel's 1900 calendar included."""
    w = weekday(serial, date1904=ctx.date1904, leap=ctx.quirks.leap_year_1900)
    assert w is not None
    return (w + 5) % 7


def _day_number(v: Any, ctx: Context) -> float | ErrorValue:
    """A date or day count: logicals are #VALUE! here, unlike in arithmetic."""
    if isinstance(v, bool):
        return VALUE
    return to_number(scalar(v, ctx), ctx)


def _workday(
    ctx: Context, start: Any, days: Any, weekend: Any, holidays: Any
) -> float | ErrorValue:
    start, days = _day_number(start, ctx), _day_number(days, ctx)
    if isinstance(start, ErrorValue):
        return start
    if isinstance(days, ErrorValue):
        return days
    off = _weekend(weekend, ctx)
    if isinstance(off, ErrorValue):
        return off
    if len(off) == 7:
        return VALUE
    hol = _holidays(holidays, ctx)
    if isinstance(hol, ErrorValue):
        return hol
    s, n, limit = math.trunc(start), math.floor(days), _limit(ctx)
    if s < 0 or s > limit:
        return NUM
    step = 1 if n > 0 else -1
    left = abs(n)
    while left:
        s += step
        if s < 0 or s > limit:
            return NUM
        if _day_of_week(s, ctx) not in off and s not in hol:
            left -= 1
    return float(s)


def _networkdays(
    ctx: Context, start: Any, end: Any, weekend: Any, holidays: Any
) -> float | ErrorValue:
    start, end = _day_number(start, ctx), _day_number(end, ctx)
    if isinstance(start, ErrorValue):
        return start
    if isinstance(end, ErrorValue):
        return end
    off = _weekend(weekend, ctx)
    if isinstance(off, ErrorValue):
        return off
    hol = _holidays(holidays, ctx)
    if isinstance(hol, ErrorValue):
        return hol
    a, b, limit = math.trunc(start), math.trunc(end), _limit(ctx)
    if not (0 <= a <= limit and 0 <= b <= limit):
        return NUM
    sign = 1 if a <= b else -1
    lo, hi = min(a, b), max(a, b)
    weeks, rest = divmod(hi - lo + 1, 7)
    count = weeks * (7 - len(off))
    first = _day_of_week(lo, ctx)
    count += sum(1 for i in range(rest) if (first + i) % 7 not in off)
    count -= sum(1 for h in hol if lo <= h <= hi and _day_of_week(h, ctx) not in off)
    return float(sign * count)


@function("WORKDAY", "value", "value", "raw", min_args=2, category=D)
def WORKDAY(ctx: Context, start: Any, days: Any, holidays: Any = MISSING) -> Any:
    """The date ``days`` working days (Monday to Friday, not holidays) from ``start``."""
    return _workday(ctx, start, days, MISSING, holidays)


@function("WORKDAY.INTL", "value", "value", "raw", "raw", min_args=2, category=D)
@function("_xlfn.WORKDAY.INTL", "value", "value", "raw", "raw", min_args=2, category=D)
def WORKDAY_INTL(
    ctx: Context, start: Any, days: Any, weekend: Any = MISSING, holidays: Any = MISSING
) -> Any:
    """WORKDAY with the weekend days chosen by a code or a seven-character mask.

    Its file form has no prefix; Excel also computes the ``_xlfn.`` spelling (capture
    ``workdays``), so both are registered.
    """
    return _workday(ctx, start, days, _code(weekend, ctx), holidays)


@function("NETWORKDAYS", "value", "value", "raw", min_args=2, category=D)
def NETWORKDAYS(ctx: Context, start: Any, end: Any, holidays: Any = MISSING) -> Any:
    """Working days from ``start`` to ``end``, both counted; negative when ``end`` is earlier."""
    return _networkdays(ctx, start, end, MISSING, holidays)


@function("NETWORKDAYS.INTL", "value", "value", "raw", "raw", min_args=2, category=D)
@function("_xlfn.NETWORKDAYS.INTL", "value", "value", "raw", "raw", min_args=2, category=D)
def NETWORKDAYS_INTL(
    ctx: Context, start: Any, end: Any, weekend: Any = MISSING, holidays: Any = MISSING
) -> Any:
    """NETWORKDAYS with the weekend days chosen by a code or a seven-character mask."""
    return _networkdays(ctx, start, end, _code(weekend, ctx), holidays)


# Time parts -------------------------------------------------------------------------------------


def _seconds(x: float, ctx: Context) -> int | ErrorValue:
    """The time of day of a serial in whole seconds, rounded to the nearest second."""
    if x < 0 or x >= _limit(ctx) + 1:
        return NUM
    return round((x - math.floor(x)) * 86400) % 86400


@function("HOUR", "number", category=D)
def HOUR(ctx: Context, x: float) -> Any:
    """The hour, 0 to 23, of a time."""
    s = _seconds(x, ctx)
    return s if isinstance(s, ErrorValue) else float(s // 3600)


@function("MINUTE", "number", category=D)
def MINUTE(ctx: Context, x: float) -> Any:
    """The minute, 0 to 59, of a time."""
    s = _seconds(x, ctx)
    return s if isinstance(s, ErrorValue) else float(s // 60 % 60)


@function("SECOND", "number", category=D)
def SECOND(ctx: Context, x: float) -> Any:
    """The second, 0 to 59, of a time."""
    s = _seconds(x, ctx)
    return s if isinstance(s, ErrorValue) else float(s % 60)
