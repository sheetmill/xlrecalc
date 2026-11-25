"""One formula evaluated outside a workbook: ``evaluate`` and ``Formula`` (ARCHITECTURE.md §11).

The formula is compiled as a cell's formula would be and its template run at a position, on an
engine whose cells come from a mapping or a callback (a resolver). It is not placed in a cell:
a conditional format's formula may read the cell it applies to.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass
from typing import Any

import xlformat

from xlrecalc._arrays import ArrayLimitError, NotImplementedFormError, element, is_array, shape
from xlrecalc._engine import Engine, Template, _at, _Cells
from xlrecalc._formula import FormulaSyntaxError, TokenizeError
from xlrecalc._formula._codegen import EXTERNAL
from xlrecalc._functions._excel_names import EXCEL_NAMES
from xlrecalc._limits import StepLimitError
from xlrecalc._load import refused_in_file
from xlrecalc._platform import PlatformName, resolve
from xlrecalc._quirks import Quirks
from xlrecalc._report import column_letters
from xlrecalc._store import COLS, MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Value
from xlrecalc._workbook import _REF, FormulaError, _axis, _constant, _file_formula

__all__ = ["Formula", "Reference", "UnsupportedError", "evaluate"]

_OWN = "*"  # the formula's own sheet when it has no name: no sheet can be called that


class UnsupportedError(FormulaError):
    """The formula uses what Excel computes and this package does not: a function or a form.

    The forms: a union of ranges, a structured, 3-D or external reference, ``CELL``'s
    formatting kinds, ``LINEST`` with several regressors, and the others the workbook's report
    names (``Report.unsupported``).

    A name Excel itself does not know is no error here: it evaluates to ``#NAME?``, as in Excel.
    """


@dataclass(frozen=True, slots=True)
class Reference:
    """A cell or range a formula reads, 1-based; ``sheet`` is ``None`` for the formula's own."""

    sheet: str | None
    first_row: int
    first_col: int
    last_row: int
    last_col: int

    @property
    def is_cell(self) -> bool:
        return self.first_row == self.last_row and self.first_col == self.last_col

    @property
    def a1(self) -> str:
        """The reference as A1 text, sheet-qualified when it names a sheet."""
        first = f"{column_letters(self.first_col)}{self.first_row}"
        text = first
        if not self.is_cell:
            text += f":{column_letters(self.last_col)}{self.last_row}"
        if self.sheet is None:
            return text
        quoted = "'" + self.sheet.replace("'", "''") + "'"
        return f"{quoted}!{text}"


Resolver = Callable[[Reference], Any]
"""Gives a reference's value: one value for a cell, rows of values for a range."""


def evaluate(
    formula: str,
    values: Mapping[str, Any] | None = None,
    *,
    resolver: Resolver | None = None,
    sheet: str | None = None,
    at: str | None = None,
    names: Mapping[str, Any] | None = None,
    legacy: bool = False,
    quirks: Quirks | None = None,
    date1904: bool = False,
    platform: str | xlformat.Platform | None = None,
    regional_settings: xlformat.RegionalSettings | None = None,
) -> Any:
    """Evaluate one formula outside a workbook, as Excel would in a cell.

    Args:
        formula: The formula as typed, starting with ``=`` (``=XLOOKUP(...)``), or in file
            form as a file stores it, without ``=`` (``_xlfn.XLOOKUP(...)``), used as it is.
        values: Cell values by reference: ``"A1"``, ``"Sheet 2!B3"``, a range (``"B2:C3"``,
            rows of values). A cell given no value is blank, as in Excel.
        resolver: Called with a :class:`Reference` for what ``values`` does not hold.
        sheet: The name of the formula's own sheet, if references name it.
        at: The cell the formula is evaluated in (default A1).
        names: Defined names: a value, or a formula text starting with ``=``.
        legacy: Evaluate as a legacy formula (implicit intersection), not as a formula typed
            in Excel today (a dynamic array).
        quirks: Departures from Excel's behavior, as in :func:`load`.
        date1904: The 1904 date system.
        platform: Which Excel computes it: ``"mac"`` (the default) or ``"windows"``, or an
            ``xlformat.Platform``, which counts by its name.
        regional_settings: The viewer's regional settings, an ``xlformat.RegionalSettings``
            (en-US only), which ``TEXT`` follows; by default the platform's en-US preset.

    Returns:
        The value a cell would show: a number, text, logical or :class:`ErrorValue`, or for a
        dynamic array its rows as a tuple of tuples.

    Raises:
        FormulaError: Excel's formula bar would not take the formula.
        UnsupportedError: It uses a function or form this package does not compute.
    """
    parsed = Formula.parse(
        formula,
        sheet=sheet,
        origin=at,
        legacy=legacy,
        quirks=quirks,
        platform=platform,
        regional_settings=regional_settings,
    )
    return parsed.evaluate(values, resolver=resolver, names=names, date1904=date1904)


class Formula:
    """A formula compiled once and evaluated many times, at any cell.

    Relative references move with the cell it is evaluated at, from the cell it was written
    for (``origin``), as a conditional format's formula does over its range.
    """

    def __init__(
        self,
        text: str,
        sheet: str | None,
        origin: tuple[int, int],
        legacy: bool,
        quirks: Quirks,
        platform: PlatformName = "mac",
        regional_settings: xlformat.RegionalSettings | None = None,
    ) -> None:
        self.text = text
        """The formula in file form, without ``=``."""
        self.sheet = sheet
        self.legacy = legacy
        self._origin = origin
        self._quirks = quirks
        self._engine = _AdHoc(
            sheet, quirks=quirks, platform=platform, regional_settings=regional_settings
        )
        row, col = origin
        try:
            tid = self._engine.template_for(text, row, col, array_mode=not legacy)
        except (FormulaSyntaxError, TokenizeError) as exc:
            raise FormulaError(f"{text!r}: {exc}") from exc
        self._template: Template = self._engine.templates[tid]

    @classmethod
    def parse(
        cls,
        formula: str,
        *,
        sheet: str | None = None,
        origin: str | None = None,
        legacy: bool = False,
        quirks: Quirks | None = None,
        platform: str | xlformat.Platform | None = None,
        regional_settings: xlformat.RegionalSettings | None = None,
    ) -> Formula:
        """Compile a formula: as typed (``=...``) or in file form (no ``=``).

        ``origin`` is the cell it was written for (default A1); see :func:`evaluate` for the
        rest. ``platform`` and ``regional_settings`` are taken here, as ``quirks`` is: which
        names Excel knows is decided when the formula is compiled.

        Raises:
            FormulaError: Excel's formula bar would not take it.
        """
        name, regional = resolve(platform, regional_settings)
        if not isinstance(formula, str):
            raise FormulaError(f"a formula is text: {formula!r}")
        if formula.startswith("="):  # as typed: the file form the formula bar would give
            file_text = _file_formula(formula, dynamic=not legacy)
        else:  # the file form, as a file stores it: used as it is
            file_text = formula
            reason = refused_in_file(file_text)
            if reason is not None:
                raise FormulaError(f"{formula!r}: {reason}")
        return cls(
            file_text, sheet, _position(origin), legacy, quirks or Quirks.EXCEL, name, regional
        )

    @property
    def references(self) -> tuple[Reference, ...]:
        """Every cell or range it reads as written, placed at its origin (names excluded)."""
        row, col = self._origin
        out = []
        for spec in self._template.refs:
            if spec.kind != "box":
                continue
            r1, c1 = _at(spec.r1, row), _at(spec.c1, col)
            r2, c2 = _at(spec.r2, row), _at(spec.c2, col)
            r1, r2 = max(0, min(r1, r2)), min(MAX_ROW, max(r1, r2))
            c1, c2 = max(0, min(c1, c2)), min(MAX_COL, max(c1, c2))
            out.append(Reference(spec.sheet, r1 + 1, c1 + 1, r2 + 1, c2 + 1))
        return tuple(out)

    @property
    def names(self) -> tuple[str, ...]:
        """The defined names it reads or calls."""
        return tuple(s.name for s in self._template.refs if s.kind == "name" and s.name)

    @property
    def functions(self) -> frozenset[str]:
        """The functions it calls, in file form (upper case, with any prefix)."""
        return self._template.calls

    @property
    def unsupported(self) -> frozenset[str]:
        """What it uses that Excel computes and this package does not, as the report names it.

        Functions Excel has (``_XLFN.AGGREGATE``) and forms: a union of ranges, a structured, 3-D
        or external reference, ``CELL`` given a formatting kind as text (``CELL("format")``). A
        name Excel does not know is not here: it evaluates to ``#NAME?``, as in Excel. What
        depends on the values (``CELL`` given its kind in a cell, ``LINEST`` with several
        regressors) is found while evaluating, which raises ``UnsupportedError`` then.
        """
        template = self._template
        spans = any(spec.sheet is not None and ":" in spec.sheet for spec in template.refs)
        return (
            frozenset(n for n in template.unsupported if _excel_knows(n) or not _is_name(n))
            | template.forms
            | ({EXTERNAL} if spans else set())  # a 3-D reference: no sheets outside a workbook
        )

    @property
    def volatile(self) -> bool:
        """Whether its result can change with nothing it reads changing (``RAND``, ``NOW``)."""
        return self._template.volatile

    def evaluate(
        self,
        values: Mapping[str, Any] | None = None,
        *,
        resolver: Resolver | None = None,
        at: str | None = None,
        names: Mapping[str, Any] | None = None,
        date1904: bool = False,
    ) -> Any:
        """Its value at ``at`` (default its origin); see :func:`evaluate`.

        Raises:
            UnsupportedError: It uses a function or form this package does not compute.
        """
        if self.unsupported:
            raise UnsupportedError(
                "not computed by this package: " + ", ".join(sorted(self.unsupported))
            )
        row, col = self._origin if at is None else _position(at)
        # The engine that compiled the template runs it: its operators read that engine's
        # position. Its cells and names are this call's alone.
        engine = self._engine
        engine.date1904 = date1904
        engine.prepare(values or {}, resolver, names or {})
        try:
            return engine.run(self._template, row, col, self.legacy)
        finally:
            engine.prepare({}, None, {})


def _position(ref: str | None) -> tuple[int, int]:
    if ref is None:
        return 0, 0
    m = _REF.match(ref.strip())
    if m is None or m["c2"] or m["sheet"]:
        raise ValueError(f"not a cell reference: {ref!r}")
    row, col = _axis(m["c1"], m["r1"])
    return row - 1, col - 1


class _Resolving(_Cells):
    """The cell store, asking the resolver for each box before reading it."""

    __slots__ = ("engine",)

    def __init__(self, engine: _AdHoc) -> None:
        super().__init__(engine.sheets, engine.max_array_cells)
        self.engine = engine

    def get(self, sheet: int, row: int, col: int) -> Value:
        self.engine.resolve(sheet, row, col, row, col)
        return super().get(sheet, row, col)

    def box(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, Value]]:
        self.engine.resolve(sheet, r1, c1, r2, c2)
        return super().box(sheet, r1, c1, r2, c2)

    def walk(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, Value]]:
        self.engine.resolve(sheet, r1, c1, r2, c2)
        return super().walk(sheet, r1, c1, r2, c2)

    def values(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> list[Value]:
        self.engine.resolve(sheet, r1, c1, r2, c2)
        return super().values(sheet, r1, c1, r2, c2)


class _AdHoc(Engine):
    """An engine holding only given values; sheets are made as references name them."""

    def __init__(self, sheet: str | None, **options: Any) -> None:
        super().__init__(**options)
        self.own = sheet
        self.add_sheet(sheet if sheet is not None else _OWN)
        self._resolver: Resolver | None = None
        self._resolved: dict[int, list[tuple[int, int, int, int]]] = {}
        self._given: set[tuple[int, int]] = set()  # (sheet, key) of cells the mapping gave
        self.cells = _Resolving(self)
        # No formula cells to evaluate: what OFFSET and INDIRECT read is resolved the same way.
        self.checked = self.cells  # type: ignore[assignment]

    def sheet_index(self, name: str) -> int | None:
        index = super().sheet_index(name)
        if index is None and self.sheets:  # a sheet a reference names: its cells are blank
            index = self.add_sheet(name)
        return index

    def prepare(
        self, values: Mapping[str, Any], resolver: Resolver | None, names: Mapping[str, Any]
    ) -> None:
        """Start an evaluation: no cell holds a value, then these values, names and resolver."""
        for s in self.sheets:
            s.values.clear()
        self._resolved.clear()
        self._given.clear()
        if self.names:
            self.names.clear()
            self._name_templates.clear()
            self._formulas_changed()
        self._resolver = resolver
        for ref, value in values.items():
            m = _REF.match(ref.strip()) if isinstance(ref, str) else None
            if m is None:
                raise ValueError(f"not a cell or range reference: {ref!r}")
            name = m["sheet"]
            if name is not None and name.startswith("'"):
                name = name[1:-1].replace("''", "'")
            sheet = 0 if name is None or name == self.own else self.sheet_index(name)
            assert sheet is not None
            r1, c1 = _axis(m["c1"], m["r1"])
            r2, c2 = _axis(m["c2"], m["r2"]) if m["c2"] else (r1, c1)
            self._put(sheet, r1 - 1, c1 - 1, r2 - 1, c2 - 1, value, given=True)
        for name, value in names.items():
            self.define_name(name, _name_text(value))

    def _put(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int, value: Any, given: bool = False
    ) -> None:
        r1, r2 = min(r1, r2), max(r1, r2)
        c1, c2 = min(c1, c2), max(c1, c2)
        rows: Sequence[Any]
        if r1 == r2 and c1 == c2 and not _is_grid(value):
            rows = ((value,),)
        else:
            if not _is_grid(value):
                raise ValueError(f"a range takes rows of values, not {value!r}")
            rows = value
        s = self.sheets[sheet]
        for i, row in enumerate(rows):
            if r1 + i > r2:
                break
            for j, v in enumerate(row):
                if c1 + j > c2:
                    break
                key = (r1 + i) * COLS + c1 + j
                if not given and (sheet, key) in self._given:
                    continue
                if given:
                    self._given.add((sheet, key))
                constant = _constant(v)
                if constant is not BLANK:
                    s.set(r1 + i, c1 + j, constant)

    def resolve(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> None:
        """Ask the resolver for a box once, unless an earlier box covered it."""
        if self._resolver is None:
            return
        done = self._resolved.setdefault(sheet, [])
        for a1, b1, a2, b2 in done:
            if a1 <= r1 and r2 <= a2 and b1 <= c1 and c2 <= b2:
                return
        done.append((r1, c1, r2, c2))
        name = self.sheets[sheet].name
        ref = Reference(None if sheet == 0 else name, r1 + 1, c1 + 1, r2 + 1, c2 + 1)
        value = self._resolver(ref)
        if value is not None:
            self._put(sheet, r1, c1, r2, c2, value)

    def cell(self, sheet: str | None, row: int, col: int) -> Value:
        target = self._target(sheet)
        if not isinstance(target, ErrorValue) and 0 <= row <= MAX_ROW and 0 <= col <= MAX_COL:
            self.resolve(target, row, col, row, col)
        return super().cell(sheet, row, col)

    def unsupported(self, label: str, *args: Any) -> Value:
        if _is_name(label) and not _excel_knows(label):  # a name Excel does not know
            return ErrorValue.NAME
        raise UnsupportedError(f"not computed by this package: {label}")

    def run(self, template: Template, row: int, col: int, legacy: bool) -> Any:
        self.sheet, self.row, self.col = 0, row, col
        self.reference_array = None
        try:
            value = template.fn(self, row, col)
        except ArrayLimitError as exc:
            raise UnsupportedError(f"an array of more than {exc.limit:,} cells") from exc
        except StepLimitError as exc:
            raise UnsupportedError(f"function steps past the limit of {exc.limit:,}") from exc
        except NotImplementedFormError as exc:
            raise UnsupportedError(f"not computed by this package: {exc.label}") from exc
        if legacy or not is_array(value):
            return self._result(value)
        rows, cols = shape(value)
        if rows == 1 and cols == 1:
            return self._result(element(value, 1, 1, 0, 0))
        return tuple(
            tuple(self._result(element(value, rows, cols, i, j)) for j in range(cols))
            for i in range(rows)
        )


def _excel_knows(name: str) -> bool:
    """Whether a function name as compiled (upper case, file form) is one of Excel's."""
    for prefix in ("_XLFN._XLWS.", "_XLFN."):
        if name.startswith(prefix):
            name = name[len(prefix) :]
            break
    return name in EXCEL_NAMES


def _is_name(label: str) -> bool:
    """Whether an unsupported label is a function's name, not a form (``CELL("format")``)."""
    return _NAME.fullmatch(label) is not None


_NAME = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")


def _is_grid(value: Any) -> bool:
    return isinstance(value, Sequence) and not isinstance(value, str)


def _name_text(value: Any) -> str:
    """A defined name's definition from a value or a formula text (``=...``)."""
    if isinstance(value, str) and value.startswith("="):
        return _file_formula(value)
    constant = _constant(value)
    if constant is BLANK:
        return '""'
    if isinstance(constant, bool):
        return "TRUE" if constant else "FALSE"
    if isinstance(constant, float):
        if not math.isfinite(constant):
            raise ValueError(f"not a number Excel holds: {value!r}")
        return repr(constant).replace("e", "E")
    if isinstance(constant, ErrorValue):
        return str(constant)
    return '"' + str(constant).replace('"', '""') + '"'
