"""Position-independent templates: references relative to the formula's own cell (§5.3).

A cell reference without ``$`` becomes an offset from the formula's cell, one with ``$`` stays
absolute. Two formulas whose references land the same way relative to their cells, such as
``A1*2`` in B1 and ``A2*2`` in B2 (xlsx shared formulas), have the same template key.
Coordinates are 0-based.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from xlrecalc._formula._nodes import tokens
from xlrecalc._formula._parser import Formula
from xlrecalc._formula._tokens import Kind, Token, column_number

__all__ = ["Axis", "cell_axes", "line_axes", "template_key"]

_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")
_LINE = re.compile(r"(\$?)([A-Za-z]{1,3}|\d+)")


@dataclass(frozen=True, slots=True)
class Axis:
    """One coordinate of a reference: an offset from the formula's cell, or absolute."""

    value: int
    absolute: bool

    def key(self, letter: str) -> str:
        """R1C1-style text: ``R5`` (absolute row 5, 1-based) or ``R[-1]`` (relative)."""
        if self.absolute:
            return f"{letter}{self.value + 1}"
        return f"{letter}[{self.value}]" if self.value else letter


def cell_axes(text: str, row: int, col: int) -> tuple[Axis, Axis]:
    """Row and column of a cell reference such as ``$B3`` relative to the cell (row, col)."""
    m = _CELL.fullmatch(text)
    if m is None:
        raise ValueError(f"not a cell reference: {text!r}")
    col_abs, letters, row_abs, digits = m.groups()
    r, c = int(digits) - 1, column_number(letters) - 1
    return (
        Axis(r if row_abs else r - row, bool(row_abs)),
        Axis(c if col_abs else c - col, bool(col_abs)),
    )


def line_axes(token: Token, row: int, col: int) -> tuple[Axis, Axis]:
    """The two ends of a column range (``A:C``) or row range (``1:3``)."""
    first, second = token.text.split(":")
    out = []
    for part in (first, second):
        m = _LINE.fullmatch(part)
        if m is None:
            raise ValueError(f"not a range end: {part!r}")
        absolute, value = bool(m[1]), m[2]
        if token.kind is Kind.ROWS:
            n = int(value) - 1
            out.append(Axis(n if absolute else n - row, absolute))
        else:
            n = column_number(value) - 1
            out.append(Axis(n if absolute else n - col, absolute))
    return out[0], out[1]


def template_key(formula: Formula, row: int, col: int) -> str:
    """The deduplication key of a formula in the cell (row, col).

    Whitespace is dropped and names and functions are upper-cased (Excel compares them without
    case); strings, sheet names and everything else keep their text.
    """
    parts: list[str] = []
    for tok in tokens(formula.root):
        kind = tok.kind
        if kind is Kind.CELL:
            r, c = cell_axes(tok.text, row, col)
            parts.append(r.key("R") + c.key("C"))
        elif kind in (Kind.COLUMNS, Kind.ROWS):
            first, second = line_axes(tok, row, col)
            letter = "R" if kind is Kind.ROWS else "C"
            parts.append(f"{first.key(letter)}:{second.key(letter)}")
        elif kind in (Kind.FUNCTION, Kind.NAME, Kind.BOOL):
            parts.append(tok.text.upper())
        else:
            parts.append(tok.text)
        parts.append("\x1f")  # keeps adjacent tokens apart: "1" "2" is not "12"
    return "".join(parts)
