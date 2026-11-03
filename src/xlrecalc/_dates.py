"""Excel's calendar: date serials in the 1900 and 1904 systems (ARCHITECTURE.md §4.4).

In the 1900 system with ``leap_year_1900``, February 1900 has 29 days: serial 60 is
February 29, 1900, serial 1 is January 1, 1900, and serial 0 is "January 0, 1900". Without
the quirk, serial ``d`` is ``1899-12-30 + d``. In the 1904 system serial 0 is January 1, 1904.
The last valid day is December 31, 9999 in both. Functions return ``None`` where Excel gives
``#NUM!``.
"""

from __future__ import annotations

import datetime as dt
import math

__all__ = ["MAX_SERIAL_1900", "MAX_SERIAL_1904", "date_serial", "serial_parts", "weekday"]

MAX_SERIAL_1900 = 2958465
MAX_SERIAL_1904 = 2957003
_EPOCH_1900 = dt.date(1899, 12, 31)  # serial 0 (January 0, 1900)
_EPOCH_STRICT = dt.date(1899, 12, 30)
_EPOCH_1904 = dt.date(1904, 1, 1)


def date_serial(
    year: float, month: float, day: float, *, date1904: bool = False, leap: bool = True
) -> int | None:
    """``DATE(year, month, day)``: arguments truncated, month and day may overflow.

    A year from 0 to 1899 has 1900 added (``DATE(1899,12,31)`` is December 31, 3799).
    """
    y, m, d = math.trunc(year), math.trunc(month), math.trunc(day)
    if 0 <= y < 1900:
        y += 1900
    if not 0 <= y < 10000:
        return None
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    if not 1 <= y <= 9999:
        return None
    first = dt.date(y, m, 1)
    if date1904:
        serial = (first - _EPOCH_1904).days + d - 1
        limit = MAX_SERIAL_1904
    elif leap:
        serial = (first - _EPOCH_1900).days + d - 1
        if first > dt.date(1900, 2, 1):
            serial += 1  # the month starts after Excel's February 29, 1900
        limit = MAX_SERIAL_1900
    else:
        serial = (first - _EPOCH_STRICT).days + d - 1
        limit = MAX_SERIAL_1900
    return serial if 0 <= serial <= limit else None


def serial_parts(
    serial: float, *, date1904: bool = False, leap: bool = True
) -> tuple[int, int, int] | None:
    """Year, month and day of a serial (its whole part), as ``YEAR``, ``MONTH``, ``DAY``."""
    s = math.floor(serial)
    if s < 0 or s > (MAX_SERIAL_1904 if date1904 else MAX_SERIAL_1900):
        return None
    if date1904:
        d = _EPOCH_1904 + dt.timedelta(days=s)
    elif not leap:
        d = _EPOCH_STRICT + dt.timedelta(days=s)
    elif s == 0:
        return (1900, 1, 0)
    elif s == 60:
        return (1900, 2, 29)
    else:
        d = _EPOCH_1900 + dt.timedelta(days=s if s < 60 else s - 1)
    return (d.year, d.month, d.day)


def weekday(serial: float, *, date1904: bool = False, leap: bool = True) -> int | None:
    """``WEEKDAY(serial)``: 1 for Sunday to 7 for Saturday."""
    s = math.floor(serial)
    if s < 0 or s > (MAX_SERIAL_1904 if date1904 else MAX_SERIAL_1900):
        return None
    if not date1904 and leap and s <= 60:
        return (s + 6) % 7 + 1  # the calendar before March 1900 counts the false leap day
    parts = serial_parts(s, date1904=date1904, leap=leap)
    assert parts is not None
    return dt.date(*parts).isoweekday() % 7 + 1
