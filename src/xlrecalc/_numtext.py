"""Numbers as text and text as numbers, the way formulas convert them (ARCHITECTURE.md §4.3).

Number to text is Excel's text form of a number, as ``&`` and the text functions use it: up to
15 significant digits (ties rounded down, as xlformat's 15-digit rule), written out in full
when that takes at most 20 characters without the sign, in scientific notation otherwise. It is
not a cell's ``General`` format, which shows fewer digits (``TEXT(x,"General")`` is that).

Text to number is what ``VALUE`` and implicit conversion accept (en-US): numbers with
grouping commas, a currency sign, a percent sign, parentheses for negatives and mixed
fractions; dates and times in the forms captured in ``tests/captures/text_number.tsv``. The
result is ``None`` where Excel gives ``#VALUE!``.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from decimal import ROUND_DOWN, ROUND_HALF_DOWN, ROUND_HALF_UP, Decimal
from typing import NamedTuple

from xlrecalc._dates import date_serial
from xlrecalc._quirks import Quirks

__all__ = ["DBL_MIN", "Parsed", "number_to_text", "text_to_number", "text_to_parts"]

DBL_MIN = 2.2250738585072014e-308  # smallest normal double; Excel has no subnormal numbers
_MAX_TEXT_NUMBER = Decimal("1e308")  # VALUE("1e308") is #VALUE!
_DBL_MAX = Decimal(sys.float_info.max)  # the exact largest double
_FIXED_WIDTH = 20


def number_to_text(x: float, quirks: Quirks = Quirks.EXCEL) -> str:
    """Excel's text form of a number, as ``=x&""`` gives it."""
    if x == 0 or abs(x) < DBL_MIN:
        return "0"
    exact = Decimal(x)
    if quirks.fifteen_digit_precision:
        rounded = _significant(exact, 15, ROUND_HALF_DOWN)
        if abs(rounded) > _DBL_MAX:  # rounding up would leave the doubles: 1.7976931348623E+308
            rounded = _significant(exact, 14, ROUND_DOWN)
    else:
        rounded = Decimal(repr(x))
    sign, digits, exponent = rounded.normalize().as_tuple()
    assert isinstance(exponent, int)
    text = "".join(map(str, digits))
    point = len(text) + exponent  # digits before the decimal point
    if point <= 0:
        fixed = "0." + "0" * -point + text
    elif point >= len(text):
        fixed = text + "0" * (point - len(text))
    else:
        fixed = text[:point] + "." + text[point:]
    if len(fixed) <= _FIXED_WIDTH:
        body = fixed
    else:
        e = point - 1
        exponent_text = f"E{'+' if e >= 0 else '-'}{abs(e):02d}"
        # The scientific form keeps to the same width: with a three-digit exponent the mantissa
        # loses digits, rounded half up from the 15-digit decimal (=1.2345678901234567E-300&""
        # is 1.2345678901235E-300; capture literals).
        room = _FIXED_WIDTH - len(exponent_text) - 1  # digits that fit beside the point
        if len(text) > room:
            reduced = Decimal(text[0] + "." + text[1:]).quantize(
                Decimal(1).scaleb(-(room - 1)), rounding=ROUND_HALF_UP
            )
            if reduced >= 10:  # 9.99...5 rounds up to the next power of ten
                reduced /= 10
                e += 1
                exponent_text = f"E{'+' if e >= 0 else '-'}{abs(e):02d}"
            text = str(reduced.normalize()).replace(".", "")
        mantissa = text[0] + ("." + text[1:] if len(text) > 1 else "")
        body = f"{mantissa}{exponent_text}"
    return ("-" if sign else "") + body


def _significant(value: Decimal, digits: int, rounding: str) -> Decimal:
    quantum = Decimal(1).scaleb(value.adjusted() - digits + 1)
    return value.quantize(quantum, rounding=rounding)


# Text to number ----------------------------------------------------------------------------

# Arabic-Indic digits are accepted (captured); other non-ASCII digits, fullwidth ones included,
# are not, so every pattern below matches ASCII digits only.
_ARABIC_INDIC = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
_NUMBER = re.compile(
    r"(?P<int>\d+(?:,\d{3,})*)?(?:\.(?P<frac>\d*))?(?:[eE](?P<exp>[+-]?\d+))?", re.ASCII
)
_MIXED = re.compile(r"(?P<whole>\d+) (?P<num>\d+)/(?P<den>\d+)", re.ASCII)


def text_to_number(
    text: str,
    quirks: Quirks = Quirks.EXCEL,
    *,
    date1904: bool = False,
    today: dt.date | None = None,
) -> float | None:
    """``VALUE(text)``: the number, or ``None`` for ``#VALUE!``.

    Args:
        text: The text.
        quirks: The quirks in force.
        date1904: The workbook's date system, for dates in the text.
        today: The date that fills in a missing year (``"3/4"``); by default today.
    """
    parsed = text_to_parts(text, quirks, date1904=date1904, today=today)
    return None if parsed is None else parsed.value


def literal_number(text: str, quirks: Quirks = Quirks.EXCEL) -> float:
    """A formula's number literal as Excel computes it.

    Like a text converted to a number, it keeps its first 15 significant digits under
    ``fifteen_digit_precision`` (``0.30000000000000004`` is 0.3; capture ``literals``), and is
    correctly rounded otherwise.
    """
    value = _number(text, quirks)
    return float(text) if value is None else value


class Parsed(NamedTuple):
    """What a text holds: a plain number, or a date serial and/or a time of day."""

    number: float | None = None
    date: int | None = None
    time: float | None = None

    @property
    def value(self) -> float:
        """The number ``VALUE`` gives."""
        if self.number is not None:
            return self.number
        return (self.date or 0) + (self.time or 0.0)


def text_to_parts(
    text: str,
    quirks: Quirks = Quirks.EXCEL,
    *,
    date1904: bool = False,
    today: dt.date | None = None,
) -> Parsed | None:
    """What a text holds, or ``None`` when it is not a number, date or time.

    ``DATEVALUE`` and ``TIMEVALUE`` accept only dates and times, so a plain number is kept
    apart from them.
    """
    s = text.translate(_ARABIC_INDIC).strip(" ")
    if not s:
        return None
    number = _number(s, quirks)
    if number is not None:
        return Parsed(number=number)
    return _date_time(s, quirks, date1904, today or dt.date.today())


def _number(s: str, quirks: Quirks) -> float | None:
    negative = parenthesized = False
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1].strip(" ")
        negative = parenthesized = True
    percent = False
    if s.endswith("%"):
        s, percent = s[:-1].rstrip(" "), True
    elif s.startswith("%"):
        s, percent = s[1:].lstrip(" "), True
    sign_seen = False
    if s[:1] in ("+", "-"):
        if parenthesized:
            return None  # "(-5)"
        negative ^= s[0] == "-"
        s, sign_seen = s[1:], True
    if s[:1] in ("$", "€"):
        if percent:
            return None
        s = s[1:].lstrip(" ")
        if s[:1] in ("+", "-") and not sign_seen:
            if parenthesized:
                return None
            negative ^= s[0] == "-"
            s = s[1:]
    if negative and s.startswith(("-", "+")):
        return None
    if m := _MIXED.fullmatch(s):
        den = int(m["den"])
        if den == 0:
            return None
        value = int(m["whole"]) + int(m["num"]) / den
        return _finish(-value if negative else value, percent)
    m = _NUMBER.fullmatch(s)
    if m is None or (m["int"] is None and not m["frac"]):
        return None
    digits = (m["int"] or "0").replace(",", "") + "." + (m["frac"] or "")
    if m["exp"]:
        digits += "e" + m["exp"]
    amount = Decimal(digits.rstrip(".") if digits.endswith(".") else digits)
    if amount and quirks.fifteen_digit_precision:
        amount = _significant(amount, 15, ROUND_DOWN)
    if percent:
        amount /= 100
    if abs(amount) >= _MAX_TEXT_NUMBER:
        return None
    result = float(amount)
    if abs(result) < DBL_MIN:
        result = 0.0
    return -result if negative and result else result


def _finish(value: float, percent: bool) -> float:
    return value / 100 if percent else value


# Dates and times ---------------------------------------------------------------------------

_MONTHS = {
    name: k
    for k, names in enumerate(
        [
            ("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"),
            ("may",), ("jun", "june"), ("jul", "july"), ("aug", "august"),
            ("sep", "sept", "september"), ("oct", "october"), ("nov", "november"),
            ("dec", "december"),
        ],
        1,
    )
    for name in names
}  # fmt: skip
_TIME = re.compile(
    r"(?P<h>\d+):(?P<m>\d+)(?::(?P<s>\d+(?:\.\d*)?))?(?: (?P<ampm>[AaPp][Mm]))?", re.ASCII
)
_MDY = re.compile(r"(?P<a>\d+)[/-](?P<b>\d+)(?:[/-](?P<c>\d+))?", re.ASCII)
_YMD = re.compile(r"(?P<y>\d{4})[/-](?P<m>\d+)[/-](?P<d>\d+)", re.ASCII)
_D_MON_Y = re.compile(r"(?P<d>\d+)[- ](?P<mon>[A-Za-z]+)(?:[- ](?P<y>\d+))?", re.ASCII)
_MON_D_Y = re.compile(r"(?P<mon>[A-Za-z]+) (?P<d>\d+), (?P<y>\d+)", re.ASCII)
_MON_Y = re.compile(r"(?P<mon>[A-Za-z]+)[- ](?P<y>\d{4})", re.ASCII)


def _date_time(s: str, quirks: Quirks, date1904: bool, today: dt.date) -> Parsed | None:
    time: float | None = None
    date_text = s
    for m in _TIME.finditer(s):
        if m.start() == 0 or s[m.start() - 1] == " ":
            end = m.end()
            if end != len(s) and s[end] != " ":
                continue
            time = _time(m)
            if time is None:
                return None
            date_text = (s[: m.start()] + s[end:]).strip(" ")
            break
    if not date_text:
        return Parsed(time=time) if time is not None else None
    serial = _date(date_text, quirks, date1904, today)
    if serial is None:
        return None
    return Parsed(date=serial, time=time)


def _time(m: re.Match[str]) -> float | None:
    h, minutes = int(m["h"]), int(m["m"])
    seconds = float(m["s"]) if m["s"] else 0.0
    if m["ampm"]:
        if not 1 <= h <= 12:
            return None
        h = h % 12 + (12 if m["ampm"].lower() == "pm" else 0)
    if seconds >= 60:
        return None
    return (h * 3600 + minutes * 60 + seconds) / 86400


def _year(text: str) -> int:
    y = int(text)
    if len(text) <= 2:
        return y + (2000 if y < 30 else 1900)
    return y


def _date(s: str, quirks: Quirks, date1904: bool, today: dt.date) -> int | None:
    y: int | None
    if m := _YMD.fullmatch(s):
        y, mo, d = int(m["y"]), int(m["m"]), int(m["d"])
    elif m := _MDY.fullmatch(s):
        a, b = int(m["a"]), int(m["b"])
        if m["c"] is not None:
            mo, d, y = a, b, _year(m["c"])
        elif 1 <= b <= 31 and _valid(today.year, a, b):
            mo, d, y = a, b, today.year
        else:
            mo, d, y = a, 1, _year(m["b"])  # "1/0": January 2000
    elif m := _MON_D_Y.fullmatch(s):
        mo, d, y = _MONTHS.get(m["mon"].lower(), 0), int(m["d"]), _year(m["y"])
    elif m := _D_MON_Y.fullmatch(s):
        mo, d = _MONTHS.get(m["mon"].lower(), 0), int(m["d"])
        y = _year(m["y"]) if m["y"] else today.year
    elif m := _MON_Y.fullmatch(s):
        mo, d, y = _MONTHS.get(m["mon"].lower(), 0), 1, int(m["y"])
    else:
        return None
    if not (1 <= mo <= 12 and 1900 <= y <= 9999):
        return None
    leap_day = quirks.leap_year_1900 and not date1904 and (y, mo, d) == (1900, 2, 29)
    if not leap_day and not _valid(y, mo, d):
        return None
    return date_serial(y, mo, d, date1904=date1904, leap=quirks.leap_year_1900)


def _valid(y: int, mo: int, d: int) -> bool:
    try:
        dt.date(y, mo, d)
    except ValueError:
        return False
    return True
