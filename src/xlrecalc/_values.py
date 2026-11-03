"""Excel's value types that Python has no type for (ARCHITECTURE.md §4.1 and §4.2)."""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import ClassVar, Final

import xlformat

__all__ = ["BLANK", "ErrorValue", "Ref", "Value"]

_CODE: Final = re.compile(r"#[A-Z0-9_/]+[!?]?")
_CODES: Final = (
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#CONNECT!",
    "#BUSY!",
    "#EXTERNAL!",
    "#UNKNOWN!",
    "#PYTHON!",
    "#TIMEOUT!",
)


class ErrorValue:
    """An Excel error value such as ``#N/A``. Not a ``str``.

    ``ErrorValue("#N/A")`` returns the interned ``ErrorValue.NA``; the listed codes are singletons.
    An error equals any ``str`` with its text and hashes as that text, so it compares equal
    to xlformat's ``ExcelError`` of the same code. That also makes ``ErrorValue.NA == "#N/A"``
    true; to tell the error from the text, check the type.
    """

    __slots__ = ("code",)
    code: str
    _interned: ClassVar[dict[str, ErrorValue]] = {}

    NULL: ClassVar[ErrorValue]
    DIV0: ClassVar[ErrorValue]
    VALUE: ClassVar[ErrorValue]
    REF: ClassVar[ErrorValue]
    NAME: ClassVar[ErrorValue]
    NUM: ClassVar[ErrorValue]
    NA: ClassVar[ErrorValue]

    def __new__(cls, code: str) -> ErrorValue:
        """Return the error with this code.

        Raises:
            ValueError: The text is not shaped like an error code.
        """
        code = str(code)
        known = cls._interned.get(code)
        if known is not None:
            return known
        if not _CODE.fullmatch(code):
            raise ValueError(f"not an error code: {code!r}")
        self = object.__new__(cls)
        object.__setattr__(self, "code", code)
        return self

    def __setattr__(self, name: str, value: object) -> None:
        raise AttributeError("ErrorValue is immutable")

    def __reduce__(self) -> tuple[type[ErrorValue], tuple[str]]:
        return (ErrorValue, (self.code,))

    def __eq__(self, other: object) -> bool:
        if isinstance(other, ErrorValue):
            return self.code == other.code
        if isinstance(other, str):
            return self.code == other
        return NotImplemented

    def __hash__(self) -> int:
        return hash(self.code)

    def __repr__(self) -> str:
        return f"ErrorValue({self.code!r})"

    def __str__(self) -> str:
        return self.code

    def to_xlformat(self) -> xlformat.ExcelError:
        """The same error as xlformat's ``ExcelError``."""
        return xlformat.ExcelError(self.code)


class _Blank:
    """An empty cell: 0, ``""`` or FALSE by what it meets. Distinct from ``""`` and 0."""

    __slots__ = ()
    _instance: ClassVar[_Blank | None] = None

    def __new__(cls) -> _Blank:
        if cls._instance is None:
            cls._instance = object.__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "BLANK"

    def __reduce__(self) -> str:
        return "BLANK"


BLANK: Final = _Blank()


for _code in _CODES:
    _error = object.__new__(ErrorValue)
    object.__setattr__(_error, "code", _code)
    ErrorValue._interned[_code] = _error
ErrorValue.NULL = ErrorValue("#NULL!")
ErrorValue.DIV0 = ErrorValue("#DIV/0!")
ErrorValue.VALUE = ErrorValue("#VALUE!")
ErrorValue.REF = ErrorValue("#REF!")
ErrorValue.NAME = ErrorValue("#NAME?")
ErrorValue.NUM = ErrorValue("#NUM!")
ErrorValue.NA = ErrorValue("#N/A")


class Ref:
    """A reference during evaluation: a box of cells on one sheet, 0-based and inclusive.

    It reads its cells through ``source`` (the engine: ``get(sheet, row, col)`` and
    ``box(sheet, r1, c1, r2, c2)``), and exists only while a formula is being evaluated.
    """

    __slots__ = ("c1", "c2", "r1", "r2", "sheet", "source")

    def __init__(self, source: object, sheet: int, r1: int, c1: int, r2: int, c2: int) -> None:
        self.source = source
        self.sheet = sheet
        self.r1, self.c1 = min(r1, r2), min(c1, c2)
        self.r2, self.c2 = max(r1, r2), max(c1, c2)

    @property
    def rows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def columns(self) -> int:
        return self.c2 - self.c1 + 1

    def value(self, row: int, col: int) -> Value:
        """The cell at (row, col), offsets from the box's top-left corner."""
        return self.source.get(self.sheet, self.r1 + row, self.c1 + col)  # type: ignore[attr-defined, no-any-return]

    def cells(self) -> Iterator[tuple[int, int, Value]]:
        """Non-blank cells, row-major, as (row offset, col offset, value)."""
        for r, c, v in self.source.box(self.sheet, self.r1, self.c1, self.r2, self.c2):  # type: ignore[attr-defined]
            yield r - self.r1, c - self.c1, v

    def values(self) -> list[Value]:
        """The non-blank values, row-major, as ``cells`` yields them, in one list."""
        return self.source.values(self.sheet, self.r1, self.c1, self.r2, self.c2)  # type: ignore[attr-defined, no-any-return]

    def is_cell(self) -> bool:
        return self.r1 == self.r2 and self.c1 == self.c2

    def __repr__(self) -> str:
        return f"Ref({self.sheet}, {self.r1}, {self.c1}, {self.r2}, {self.c2})"


Value = float | str | bool | ErrorValue | _Blank
"""A cell value: number, text, logical, error or blank (arrays come with spill)."""
