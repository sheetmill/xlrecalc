"""What a recalculation could not do, and the cells it names (ARCHITECTURE.md §7.6)."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import NamedTuple

__all__ = ["CellRef", "Report", "column_letters"]

# A sheet name needs no quotes when it is a word (letters, digits and underscores, not starting
# with a digit) that cannot be read as a cell reference. Excel quotes some names this leaves
# out (it keeps 'v1.2'!E5 quoted: capture formula_bar), and quoting a name that needs none is
# harmless, so a name with any other character is quoted.
_PLAIN = re.compile(r"[^\W\d]\w*")
_A1_LIKE = re.compile(r"[A-Za-z]{1,3}\d+")
_R1C1_LIKE = re.compile(r"[Rr]\d*([Cc]\d*)?|[Cc]\d*")


def column_letters(col: int) -> str:
    """A 1-based column number as letters: 1 is ``A``, 28 is ``AB``."""
    letters = ""
    while col:
        col, rest = divmod(col - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def sheet_prefix(sheet: str) -> str:
    """A sheet name as a reference writes it, quoted when needed: ``'Sheet 1'!``."""
    if (
        _PLAIN.fullmatch(sheet)
        and not _A1_LIKE.fullmatch(sheet)
        and not _R1C1_LIKE.fullmatch(sheet)
    ):
        return sheet + "!"
    return "'" + sheet.replace("'", "''") + "'!"


class CellRef(NamedTuple):
    """A cell of a workbook: sheet name, row and column, both 1-based.

    ``str()`` gives the address as Excel writes it, the sheet quoted when its name needs it:

    >>> str(CellRef("Sheet 1", 7, 2))
    "'Sheet 1'!B7"
    """

    sheet: str
    row: int
    col: int

    def __str__(self) -> str:
        return f"{sheet_prefix(self.sheet)}{column_letters(self.col)}{self.row}"


@dataclass
class Report:
    """What a recalculation could not do: the cells it left as they were, and why.

    The report accumulates over the recalculations of one workbook. Cells are ``CellRef``
    values.

    Attributes:
        unsupported: Unsupported function or feature, and the number of cells using it.
            ``"iterative calculation"`` counts the cells of ``uncomputed`` that are on a circular
            reference, or depend on one, in a workbook with iterative calculation on.
        exceptions: Cells where an internal error was caught, with the error; such a cell is
            ``#VALUE!`` and uncomputed.
        circular: Cells on a circular reference or depending on one, with iterative calculation
            off. Excel leaves them uncalculated: they keep the value the file holds, and one
            without a value is the empty text (0 for a formula set through the API).
        from_kept: A cell holding the cached value of a function or feature not implemented, and
            every cell computed from one. Such a value is right if the file's value is. After a
            change upstream of that formula, the file's value is out of date: the cells leave this
            set for ``uncomputed``. A change that removes the reliance takes them out altogether.
            The two sets never overlap.
        refused: Cells whose formula Excel refuses in a file, with the reason. They keep their
            cached value and are written as they were; Excel's repair turns each into a
            constant holding that value when it opens the file.
        uncomputed: Cells whose value could not be computed (an unsupported function or
            feature without a cached value, a caught internal error, a circular reference with
            iterative calculation on, or an uncomputed precedent). A written file keeps their
            cached values as they were. An array formula is named by its first cell, and the
            cells of its range keep their values with it.
    """

    unsupported: Counter[str] = field(default_factory=Counter)
    exceptions: dict[CellRef, str] = field(default_factory=dict)
    circular: set[CellRef] = field(default_factory=set)
    from_kept: set[CellRef] = field(default_factory=set)
    refused: dict[CellRef, str] = field(default_factory=dict)
    uncomputed: set[CellRef] = field(default_factory=set)
