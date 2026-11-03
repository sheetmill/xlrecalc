"""Which Excel quirks to reproduce (ARCHITECTURE.md §4.4; evidence in QUIRKS.md)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import ClassVar

import xlformat

__all__ = ["Quirks"]


@dataclass(frozen=True, slots=True)
class Quirks:
    """Which Excel quirks to reproduce. All on by default (the ``EXCEL`` preset).

    ``Quirks.EXCEL`` has every quirk on and is what ``quirks=None`` means; ``Quirks.STRICT``
    has every quirk off. Individual fields can be set: ``Quirks(comparison_rounding=False)``.
    The first two are shared with xlformat and mean the same there.

    Attributes:
        leap_year_1900: Serial 60 is February 29, 1900, a day that did not exist, and weekdays
            before March 1900 are shifted, as in Excel's 1900 date system.
        fifteen_digit_precision: A number becomes text with at most 15 significant digits
            (``0.1+0.2`` gives ``"0.3"``), and text becomes a number from its first 15
            significant digits, the rest truncated.
        comparison_rounding: ``=``, ``<>``, ``<``, ``>`` and the criteria of ``COUNTIF`` and its
            family treat numbers that agree to 15 significant digits as equal
            (``0.1+0.2=0.3`` is TRUE). Exact lookups (``MATCH``, ``VLOOKUP``) never do.
        cancellation_snapping: A formula whose last operation is an addition or subtraction
            gives 0 when the result is less than 8 units in the last place of the left
            operand (``=0.1+0.2-0.3`` is 0).
    """

    leap_year_1900: bool = True
    fifteen_digit_precision: bool = True
    comparison_rounding: bool = True
    cancellation_snapping: bool = True

    EXCEL: ClassVar[Quirks]
    STRICT: ClassVar[Quirks]

    def to_xlformat(self) -> xlformat.Quirks:
        """The settings for xlformat of the quirks the two packages share."""
        return xlformat.Quirks(
            leap_year_1900=self.leap_year_1900,
            fifteen_digit_precision=self.fifteen_digit_precision,
        )


Quirks.EXCEL = Quirks()
Quirks.STRICT = Quirks(
    leap_year_1900=False,
    fifteen_digit_precision=False,
    comparison_rounding=False,
    cancellation_snapping=False,
)
