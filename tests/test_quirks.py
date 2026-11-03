"""Each quirk switched off gives the stated alternative (QUIRKS.md)."""

from xlrecalc._arith import snap
from xlrecalc._compare import compare
from xlrecalc._dates import date_serial, serial_parts, weekday
from xlrecalc._numtext import number_to_text, text_to_number
from xlrecalc._quirks import Quirks

EXCEL, STRICT = Quirks.EXCEL, Quirks.STRICT


def test_fifteen_digit_precision():
    assert number_to_text(0.1 + 0.2, EXCEL) == "0.3"
    assert number_to_text(0.1 + 0.2, STRICT) == "0.30000000000000004"
    assert text_to_number("0.12345678901234567", EXCEL) == 0.123456789012345
    assert text_to_number("0.12345678901234567", STRICT) == 0.12345678901234567


def test_comparison_rounding():
    assert compare(0.1 + 0.2, 0.3, "=", EXCEL) is True
    assert compare(0.1 + 0.2, 0.3, "=", STRICT) is False
    assert compare(0.1 + 0.2, 0.3, ">", STRICT) is True


def test_cancellation_snapping():
    a = 0.1 + 0.2
    assert snap(a, a - 0.3, EXCEL) == 0.0
    assert snap(a, a - 0.3, STRICT) == a - 0.3


def test_leap_year_1900():
    assert date_serial(1900, 2, 29, leap=True) == 60
    assert serial_parts(60, leap=True) == (1900, 2, 29)
    assert weekday(1, leap=True) == 1  # Sunday, as Excel shows January 1, 1900
    assert serial_parts(60, leap=False) == (1900, 2, 28)
    assert serial_parts(61, leap=False) == (1900, 3, 1)
    assert weekday(1, leap=False) == 1  # 1899-12-31 was a Sunday
    assert date_serial(1900, 3, 1, leap=False) == 61
    assert text_to_number("2/29/1900", STRICT) is None
