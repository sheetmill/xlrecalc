"""``load`` and ``Workbook``: read values, change cells, recalculate, save (§11.2)."""

from __future__ import annotations

import enum
import io
import math
import os
import re
from typing import BinaryIO, TypeAlias

from xlrecalc._engine import Engine, node
from xlrecalc._formula import FormulaSyntaxError, TokenizeError, tokenize
from xlrecalc._formula._fastkey import translate
from xlrecalc._formula._fileform import display_form, file_form
from xlrecalc._formula._tokens import Kind, column_number
from xlrecalc._load import load as _load_engine
from xlrecalc._load import refused_in_file
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef, Report
from xlrecalc._save import array_edits, file_value, lookups
from xlrecalc._store import COLS, MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Value
from xlrecalc._xlsx import ReadLimits
from xlrecalc._xlsx._write import Edit, write

__all__ = ["CellValue", "FormulaError", "Workbook", "load"]

CellValue: TypeAlias = float | str | bool | ErrorValue | None
"""A cell's value: a number (always ``float``), text, a logical, an error, or ``None`` (blank)."""

_MAX_TEXT = 32767
_REF = re.compile(
    r"""^(?:(?P<sheet>'(?:[^']|'')+'|[^'!:]+)!)?
        (?P<c1>\$?[A-Za-z]{1,3})(?P<r1>\$?\d+)
        (?::(?P<c2>\$?[A-Za-z]{1,3})(?P<r2>\$?\d+))?$""",
    re.VERBOSE,
)


class FormulaError(ValueError):
    """A formula the formula bar would not take: not well formed, or a form Excel refuses."""


def _axis(letters: str, digits: str) -> tuple[int, int]:
    col = column_number(letters.lstrip("$"))
    row = int(digits.lstrip("$"))
    if not (1 <= col <= MAX_COL + 1 and 1 <= row <= MAX_ROW + 1):
        raise ValueError(f"outside the sheet: {letters}{digits}")
    return row, col


def _public(value: Value) -> CellValue:
    return None if value is BLANK else value  # type: ignore[return-value]


def _constant(value: object) -> Value:
    """A value passed to ``set`` as the engine holds it."""
    if value is None:
        return BLANK
    if isinstance(value, bool):
        return value
    if isinstance(value, ErrorValue):
        return value
    if isinstance(value, int | float):
        number = float(value)
        if not math.isfinite(number):
            raise ValueError(f"not a number Excel holds: {value!r}")
        return number
    if isinstance(value, str):
        if isinstance(value, enum.Enum):  # an error enum such as xlformat.ExcelError
            try:
                return ErrorValue(str(value))
            except ValueError:
                pass
        if len(value) > _MAX_TEXT:
            raise ValueError(f"text longer than Excel's {_MAX_TEXT} characters")
        return str(value)
    raise TypeError(f"not a cell value: {value!r} (numbers, text, logicals, errors or None)")


class Workbook:
    """A workbook loaded from an xlsx file. Use :func:`load` to get one.

    Values are evaluated when asked for: a formula whose cached value the file holds and
    trusts is not evaluated until something it depends on changes. Cells are named as Excel
    names them, ``"Sheet1!B7"`` or ``"'My sheet'!B7"``, or as a :class:`CellRef`; a reference
    without a sheet (``"B7"``) is allowed only in a workbook of one sheet. One workbook is used
    from one thread at a time.

    Unlike openpyxl, where assigning ``"=A1"`` to a cell makes a formula, :meth:`set` always
    stores a constant: ``wb.set("Sheet1!A2", "=A1")`` stores the text ``=A1``. Use
    :meth:`set_formula` for a formula.
    """

    def __init__(
        self, engine: Engine, source: str | os.PathLike[str] | BinaryIO, limits: ReadLimits | None
    ) -> None:
        self._engine = engine
        self._source = source
        self._limits = limits
        self._edits: dict[int, dict[tuple[int, int], bool]] = {}  # sheet -> cell -> formula?
        self._formula_removed = False
        # formulas set: (legacy reading intersects, legacy file form)
        self._marked: dict[tuple[int, int, int], tuple[bool, str]] = {}

    @property
    def report(self) -> Report:
        """What evaluation could not do so far (unsupported functions, uncomputed cells, ...)."""
        return self._engine.report

    @property
    def sheet_names(self) -> list[str]:
        """The names of the workbook's sheets, in order."""
        return [s.name for s in self._engine.sheets]

    # References -------------------------------------------------------------------------------

    def _box(self, ref: str | CellRef) -> tuple[int, int, int, int, int]:
        """A reference as (sheet, r1, c1, r2, c2), 1-based."""
        if isinstance(ref, CellRef):
            sheet, r1, c1 = ref
            return self._sheet(sheet), r1, c1, r1, c1
        m = _REF.match(ref.strip())
        if m is None:
            raise ValueError(f"not a cell or range reference: {ref!r}")
        name = m["sheet"]
        if name is not None and name.startswith("'"):
            name = name[1:-1].replace("''", "'")
        r1, c1 = _axis(m["c1"], m["r1"])
        r2, c2 = _axis(m["c2"], m["r2"]) if m["c2"] else (r1, c1)
        return (
            self._sheet(name),
            min(r1, r2),
            min(c1, c2),
            max(r1, r2),
            max(c1, c2),
        )

    def _sheet(self, name: str | None) -> int:
        if name is None:
            sheets = self.sheet_names
            if len(sheets) != 1:
                raise ValueError(
                    f"a reference without a sheet is ambiguous in a workbook of {len(sheets)} "
                    f"sheets; name one of {sheets}"
                )
            return 0
        index = self._engine.sheet_index(name)
        if index is None:
            raise KeyError(f"no sheet named {name!r}")
        return index

    def _cell(self, ref: str | CellRef) -> tuple[int, int, int]:
        sheet, r1, c1, r2, c2 = self._box(ref)
        if (r1, c1) != (r2, c2):
            raise ValueError(f"a single cell is needed, not a range: {ref!r}")
        return sheet, r1 - 1, c1 - 1

    def _check_editable(self, sheet: int, row: int, col: int) -> None:
        for r1, c1, r2, c2 in self._engine.locked.get(sheet, ()):
            if r1 <= row <= r2 and c1 <= col <= c2:
                raise ValueError("the cell is part of an array formula or a data table")

    # Reading ----------------------------------------------------------------------------------

    def value(self, ref: str | CellRef) -> CellValue:
        """The cell's value, evaluating what it depends on first if needed.

        An error is an :class:`ErrorValue`; it equals its text (``ErrorValue.NA == "#N/A"``),
        so check the type to tell the error from the text ``"#N/A"``.
        """
        sheet, row, col = self._cell(ref)
        return _public(self._engine.get(sheet, row, col))

    def values(self, ref: str) -> list[list[CellValue]]:
        """The values of a range, as a list of rows."""
        sheet, r1, c1, r2, c2 = self._box(ref)
        get = self._engine.get
        return [
            [_public(get(sheet, r - 1, c - 1)) for c in range(c1, c2 + 1)]
            for r in range(r1, r2 + 1)
        ]

    def formula(self, ref: str | CellRef) -> str | None:
        """The cell's formula as the formula bar shows it (``"=B7*2"``), or ``None``."""
        sheet, row, col = self._cell(ref)
        text = self._engine.formula_text(sheet, row, col)
        return None if text is None else "=" + display_form(text)

    # Changing ---------------------------------------------------------------------------------

    def set(self, ref: str | CellRef, value: object) -> None:
        """Make the cell a constant: a number, text, a logical, an error, or ``None`` to clear it.

        A text is always text, even ``"=A1"`` or ``"#N/A"``: this differs from openpyxl, where
        assigning ``"=A1"`` makes a formula. Use :meth:`set_formula` for a formula and
        :class:`ErrorValue` for an error. A formula the cell held is removed.

        Raises:
            ValueError: The cell is part of an array formula or data table, or the value is
                not one Excel holds (a non-finite number, a text over 32,767 characters).
            TypeError: The value is not a number, text, logical, error or ``None``.
        """
        sheet, row, col = self._cell(ref)
        self._check_editable(sheet, row, col)
        constant = _constant(value)
        if row * COLS + col in self._engine.sheets[sheet].templates:
            self._formula_removed = True
        self._engine.set_value(sheet, row, col, constant)
        self._edits.setdefault(sheet, {})[row + 1, col + 1] = False

    def set_formula(self, ref: str | CellRef, formula: str) -> None:
        """Give the cell a formula, as typed in the formula bar: ``"=B7*2"``.

        The formula's dependents and the cell are evaluated when a value is next read, or at
        :meth:`recalculate` or :meth:`save`.

        Raises:
            FormulaError: The text is not a formula the formula bar takes (no leading ``=``,
                not well formed, a tab, ``SUM( )``, the error literals ``#SPILL!`` and
                ``#CALC!``).
            ValueError: The cell is part of an array formula or a data table.
        """
        sheet, row, col = self._cell(ref)
        self._check_editable(sheet, row, col)
        legacy = _file_formula(formula)
        text = _file_formula(formula, dynamic=True)
        engine = self._engine
        try:
            # As Excel 365 takes a typed formula: a dynamic array, which spills (§7.7). Its
            # text read as a legacy formula (an @ written out as _xlfn.SINGLE) tells whether
            # the file needs the dynamic-array marker: an implicit intersection anywhere else.
            intersects = engine.templates[engine.template_for(text, row, col)].intersects
            engine.set_formula(sheet, row, col, text, array=(True, (row, col, row, col)))
        except (FormulaSyntaxError, TokenizeError) as exc:
            raise FormulaError(f"{formula!r}: {exc}") from exc
        self._edits.setdefault(sheet, {})[row + 1, col + 1] = True
        self._marked[sheet, row, col] = (intersects, legacy)

    def _needs_marker(self, sheet: int, row: int, col: int) -> bool:
        """Whether a formula needs the dynamic-array marker in the file.

        A formula set here is a dynamic array; as Excel writes a typed formula (capture
        ``marker_typed``), it is marked when a legacy formula would mean something else: it
        spills, or takes an implicit intersection other than an explicit ``@`` (``=@A1:A3`` is
        written as the plain ``A1:A3``). A dynamic array read from the file stays one.
        """
        anchor = self._engine.anchors.get(node(sheet, row, col))
        if anchor is None:
            return False
        marked = self._marked.get((sheet, row, col))
        if marked is None:
            return True  # from the file
        r1, c1, r2, c2 = anchor.box
        return marked[0] or (r1, c1) != (r2, c2) or anchor.wanted is not None

    # Evaluating and saving --------------------------------------------------------------------

    def recalculate(self, full: bool = False) -> Report:
        """Evaluate every stale formula (every formula with ``full``); returns the report."""
        return self._engine.recalculate(full=full)

    def save(
        self, path: str | os.PathLike[str] | None = None, *, full_calc_on_load: bool = True
    ) -> None:
        """Write the workbook with its changes and current values, surgically and atomically.

        Every part the changes do not touch is copied from the source file byte for byte; see
        :func:`recalculate` for ``full_calc_on_load``.

        Args:
            path: Where to write; by default the file the workbook was loaded from.
            full_calc_on_load: Ask Excel to recalculate when it opens the file.

        Raises:
            ValueError: No ``path`` for a workbook loaded from a file object.
        """
        if path is None:
            if not isinstance(self._source, (str, os.PathLike)):
                raise ValueError("a workbook loaded from a file object needs a path to save to")
            path = self._source
        engine = self._engine
        engine.recalculate()
        uncomputed = {(c.sheet.casefold(), c.row, c.col) for c in engine.report.uncomputed}
        edits: dict[str, dict[tuple[int, int], Edit]] = {}
        for sheet, cells in self._edits.items():
            s = engine.sheets[sheet]
            out = edits.setdefault(s.name, {})
            for (row, col), is_formula in cells.items():
                value = s.values.get((row - 1) * COLS + col - 1, BLANK)
                text = engine.formula_text(sheet, row - 1, col - 1) if is_formula else None
                if text is not None and (s.name.casefold(), row, col) in uncomputed:
                    out[row, col] = Edit(text, None)
                else:
                    out[row, col] = Edit(text, file_value(value))
        filled_edits, fills = array_edits(engine)
        for name, filled in filled_edits.items():
            index = self._sheet(name)
            target = edits.setdefault(name, {})
            for (row, col), edit in filled.items():
                if edit.formula is not None and not self._needs_marker(index, row - 1, col - 1):
                    marked = self._marked.get((index, row - 1, col - 1))
                    text = edit.formula if marked is None else marked[1]
                    edit = Edit(text, edit.value)  # a plain formula means the same
                target[row, col] = edit
        source = self._source
        if not isinstance(source, (str, os.PathLike)):
            source.seek(0)
        write(
            source,
            path,
            lookups(engine),
            edits=edits,
            fills=fills,
            translate=translate,
            drop_calc_chain=self._formula_removed,
            full_calc_on_load=full_calc_on_load,
            limits=self._limits,
        )


def _file_formula(formula: str, dynamic: bool = False) -> str:
    """The file form of a formula as typed, checked as the formula bar checks it."""
    if not isinstance(formula, str) or not formula.startswith("="):
        raise FormulaError(f"a formula starts with '=': {formula!r}")
    try:
        tokens = tokenize(formula[1:])
    except TokenizeError as exc:
        raise FormulaError(f"{formula!r}: {exc}") from exc
    previous = None
    for tok in tokens:
        if "\t" in tok.space:
            raise FormulaError(f"{formula!r}: a tab (the formula bar does not take one)")
        empty_call = previous is not None and previous.kind is Kind.OPEN
        if tok.kind is Kind.CLOSE and empty_call and tok.space:
            raise FormulaError(f"{formula!r}: an argument list of spaces")
        previous = tok
    try:
        text = file_form(formula, dynamic)
    except ValueError as exc:
        raise FormulaError(f"{formula!r}: {exc}") from exc
    reason = refused_in_file(text)
    if reason is not None:
        raise FormulaError(f"{formula!r}: {reason}")
    return text


def load(
    source: str | os.PathLike[str] | BinaryIO,
    *,
    quirks: Quirks | None = None,
    limits: ReadLimits | None = None,
    trust_cached: bool = True,
) -> Workbook:
    """Read an xlsx workbook.

    Nothing is evaluated yet: cached values the file holds are trusted (ARCHITECTURE.md §7.4)
    unless the file asks for a full calculation on load or ``trust_cached`` is false, and a
    value is evaluated when it is read.

    Args:
        source: The file, as a path or a binary file object (read into memory).
        quirks: Which Excel quirks to reproduce; all by default.
        limits: Limits on what reading the file may decompress.
        trust_cached: Treat the file's cached values as current.

    Raises:
        PackageError: The file is not a readable xlsx package.
        LimitError: The file exceeds ``limits``.
    """
    if not isinstance(source, (str, os.PathLike)):
        source = io.BytesIO(source.read())
    engine = _load_engine(source, quirks=quirks, limits=limits, trust_cached=trust_cached)
    return Workbook(engine, source, limits)
