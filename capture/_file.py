"""Cells as Excel's saved file writes them, and numbers and text in its forms, for the probes.

A probe records what Excel wrote, so it reads cells through :class:`Book`, which gives each
value as written: an error is its code as plain text, and an error a file caches as
``#VALUE!`` with a rich value behind it stays ``#VALUE!`` (``xlio`` reads it as the error the
rich value holds; :attr:`Book.value_errors` gives that). ``cm`` and ``vm`` are their text. A
subnormal number reads as 0, as Excel reads it; Excel does not write one.
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterator
from types import TracebackType
from typing import BinaryIO, NamedTuple

import xlio

__all__ = ["Book", "FileCell", "number_text", "text_escape", "unescape"]


class FileCell(NamedTuple):
    """One cell as the file writes it (see the module docstring)."""

    row: int
    column: int
    value: float | str | bool | None
    type: str
    formula: xlio.Formula | None
    cm: str | None = None
    vm: str | None = None


class Book:
    """A workbook opened with ``xlio``, its cells given as :class:`FileCell`."""

    def __init__(self, source: str | os.PathLike[str] | BinaryIO) -> None:
        self._wb = xlio.open(source)
        self.sheets = self._wb.sheets
        self.defined_names = self._wb.defined_names
        self.calc = self._wb.calc
        self.value_errors = {str(k): str(v) for k, v in self._wb.cell_metadata.value_errors.items()}

    def __enter__(self) -> Book:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self._wb.close()

    def cells(self, sheet: str | int | xlio.Sheet) -> Iterator[FileCell]:
        """A worksheet's cells in file order, as written."""
        s = sheet if isinstance(sheet, xlio.Sheet) else self._wb.sheet(sheet)
        rich = self._wb.cell_metadata.value_errors
        for c in s.cells():
            value = c.value
            if c.type == "e" and value is not None:
                value = "#VALUE!" if c.vm is not None and c.vm in rich else str(value)
            yield FileCell(
                c.row,
                c.column,
                value,
                c.type,
                c.formula,
                None if c.cm is None else str(c.cm),
                None if c.vm is None else str(c.vm),
            )


def number_text(value: float) -> str:
    """A number as a file stores it: the shortest exact text, integral values without ``.0``."""
    if value.is_integer() and abs(value) < 2**53:
        return str(int(value))
    return repr(value).replace("e", "E")


_ESCAPED = re.compile(r"_(?=x[0-9A-Fa-f]{4}_)|[\x00-\x08\x0b-\x1f\ud800-\udfff￾￿]")


def _escape_char(m: re.Match[str]) -> str:
    ch = m[0]
    return "_x005F_" if ch == "_" else f"_x{ord(ch):04X}_"


def text_escape(text: str) -> str:
    """Text as an element's content: XML escapes, and OOXML's ``_xHHHH_`` for what XML cannot hold.

    A carriage return is escaped too (XML would read it as a line feed), and a literal
    ``_xHHHH_`` gets its underscore escaped, so it reads back as written.
    """
    text = _ESCAPED.sub(_escape_char, text)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


_UNESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def unescape(text: str) -> str:
    """OOXML's escape of characters XML cannot hold: ``_x000D_`` is a carriage return."""
    if "_x" not in text:
        return text
    return _UNESCAPE.sub(lambda m: chr(int(m[1], 16)), text)
