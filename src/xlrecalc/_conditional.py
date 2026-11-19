"""A conditional format's formula, as Excel evaluates it over the range it applies to.

Excel's rules, from capture ``cf_rules`` (Excel 16.115 for Mac, ARCHITECTURE.md §11.2):

- a rule's relative references are measured from the top-left corner of the box around every
  area of its range (``sqref``), whatever the areas' order: for ``D2:D3 B4:C5`` that is B2;
- the formula is evaluated as a formula is today (a dynamic array), not as a legacy one, and a
  range or array result counts by its top-left element;
- the format applies where ``IF`` would take its true branch on that element: TRUE, a nonzero
  number, the text ``"TRUE"``; an error, FALSE, 0, other text and a blank do not;
- a rule holding an array constant or a structured reference makes Excel offer to repair the
  file: such a formula is refused.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from typing import Any

from xlrecalc._coerce import to_bool
from xlrecalc._compare import Op, compare
from xlrecalc._evaluate import Formula, Resolver
from xlrecalc._formula._tokens import Kind, TokenizeError, column_number, tokenize
from xlrecalc._quirks import Quirks
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._workbook import FormulaError

__all__ = ["ConditionalFormula"]

_AREA = re.compile(r"\$?([A-Z]{1,3})?\$?(\d+)?(?::\$?([A-Z]{1,3})?\$?(\d+)?)?")
_MAX_ROW, _MAX_COL = 1_048_576, 16_384
_OPERATORS: dict[str, Op] = {
    "equal": "=",
    "notEqual": "<>",
    "greaterThan": ">",
    "greaterThanOrEqual": ">=",
    "lessThan": "<",
    "lessThanOrEqual": "<=",
}
_RANGES = ("between", "notBetween")


def _areas(sqref: str) -> list[tuple[int, int, int, int]]:
    """Each area of a ``sqref`` as 1-based (row1, col1, row2, col2); a whole column or row whole."""
    areas = []
    for text in sqref.upper().split():
        m = _AREA.fullmatch(text)
        c1, r1, c2, r2 = m.groups() if m else (None, None, None, None)
        if ":" not in text:
            c2, r2 = c1, r1
            whole = bool(c1 and r1)  # one cell
        else:  # a range, a whole column (C:C) or a whole row (2:2)
            whole = bool((c1 and c2) or (r1 and r2)) and bool(c1) == bool(c2)
            whole = whole and bool(r1) == bool(r2)
        if not whole:
            raise ValueError(f"not a range a conditional format applies to: {sqref!r}")
        row1, row2 = (int(r1), int(r2)) if r1 and r2 else (1, _MAX_ROW)
        col1, col2 = (column_number(c1), column_number(c2)) if c1 and c2 else (1, _MAX_COL)
        areas.append((min(row1, row2), min(col1, col2), max(row1, row2), max(col1, col2)))
    if not areas:
        raise ValueError("a conditional format applies to at least one cell")
    return areas


def _letters(col: int) -> str:
    out = ""
    while col:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out


def _cell(ref: str) -> tuple[int, int]:
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", ref)
    if m is None:
        raise ValueError(f"not a cell: {ref!r}")
    return int(m[2]), column_number(m[1].upper())


def _refused(text: str) -> str | None:
    """Why Excel refuses a file whose rule holds this formula, if it does (capture ``cf_rules``)."""
    try:
        tokens = tokenize(text[1:] if text.startswith("=") else text)
    except TokenizeError:
        return None  # Formula.parse reports it
    for token in tokens:
        if token.kind is Kind.ARRAY_OPEN:
            return "an array constant"
        if token.kind is Kind.STRUCTURED:
            return "a structured reference"
    return None


def _first(value: Any) -> Any:
    """The top-left element of a range or array result; a blank as ``BLANK``."""
    while isinstance(value, tuple):
        value = value[0] if value else BLANK
    return BLANK if value is None else value


class ConditionalFormula:
    """A conditional format's rule over a range: where Excel shows its format.

    Made with :meth:`parse` (an ``expression`` rule's formula) or :meth:`cell_is` (a ``cellIs``
    rule's operator and bounds), from the rule as a file stores it, then asked cell by cell with
    :meth:`applies`. Relative references are measured from :attr:`anchor`, the top-left corner
    of the box around every area of :attr:`sqref` (capture ``cf_rules``).

    One rule is used from one thread at a time.
    """

    def __init__(
        self,
        sqref: str,
        formulas: tuple[Formula, ...],
        operator: str | None,
    ) -> None:
        self.sqref = sqref
        """The range the rule applies to, as given."""
        self._areas = _areas(sqref)
        row = min(a[0] for a in self._areas)
        col = min(a[1] for a in self._areas)
        self.anchor = f"{_letters(col)}{row}"
        """The cell the rule's relative references are measured from."""
        self._formulas = formulas
        self.operator = operator
        """A ``cellIs`` rule's operator, or ``None`` for an expression rule."""
        self._own = (
            Formula.parse(self.anchor, sheet=formulas[0].sheet, origin=self.anchor)
            if operator
            else None
        )

    @classmethod
    def parse(
        cls,
        formula: str,
        *,
        sqref: str,
        sheet: str | None = None,
        quirks: Quirks | None = None,
    ) -> ConditionalFormula:
        """An expression rule: the format applies where the formula is true.

        Args:
            formula: The rule's formula in file form, as a file stores it (``A2>2``), or as
                typed, starting with ``=``.
            sqref: The range the rule applies to, as the file stores it: areas separated by
                spaces (``D2:D3 B4:C5``), a whole column (``C:C``) or row (``2:2``) allowed.
            sheet: The name of the sheet the rule is on, if references name it.
            quirks: Departures from Excel's behavior, as in :func:`load`.

        Raises:
            FormulaError: The formula is not well formed, or holds what Excel refuses in a rule
                (an array constant, a structured reference: Excel offers to repair such a file).
            ValueError: ``sqref`` is not a range.
        """
        return cls(sqref, (_rule_formula(formula, sqref, sheet, quirks),), None)

    @classmethod
    def cell_is(
        cls,
        operator: str,
        formulas: Sequence[str],
        *,
        sqref: str,
        sheet: str | None = None,
        quirks: Quirks | None = None,
    ) -> ConditionalFormula:
        """A ``cellIs`` rule: the format applies where the cell's value compares as stated.

        The comparison is Excel's operators' (numbers before text before logicals, text without
        case; capture ``cf_rules``: every operator, over numbers and over text, numbers as text,
        TRUE and a blank against number and text bounds). ``between`` and ``notBetween`` take
        their two bounds in either order (4 and 2 is 2 and 4).

        Args:
            operator: As the file stores it: ``equal``, ``notEqual``, ``greaterThan``,
                ``greaterThanOrEqual``, ``lessThan``, ``lessThanOrEqual`` (one bound),
                ``between``, ``notBetween`` (two).
            formulas: The bounds' formulas, as :meth:`parse` takes one.
            sqref: The range, as :meth:`parse` takes it.
            sheet: As :meth:`parse` takes it.
            quirks: As :meth:`parse` takes it.

        Raises:
            ValueError: An operator not listed, or the wrong number of bounds for it.
            FormulaError: As :meth:`parse`.
        """
        count = 2 if operator in _RANGES else 1 if operator in _OPERATORS else 0
        if not count:
            raise ValueError(f"not a cellIs operator: {operator!r}")
        if len(formulas) != count:
            raise ValueError(f"{operator} takes {count} formula(s), not {len(formulas)}")
        parsed = tuple(_rule_formula(f, sqref, sheet, quirks) for f in formulas)
        return cls(sqref, parsed, operator)

    def applies(
        self,
        at: str,
        values: Mapping[str, Any] | None = None,
        *,
        resolver: Resolver | None = None,
        names: Mapping[str, Any] | None = None,
        date1904: bool = False,
    ) -> bool:
        """Whether Excel shows the rule's format in a cell of its range.

        The formula is evaluated at the cell, as :meth:`Formula.evaluate` evaluates it: cell
        values from ``values``, then from ``resolver``; ``names`` as there. The format applies
        where ``IF`` would take its true branch on the result's top-left element: TRUE, a
        nonzero number or the text ``"TRUE"``. An error, FALSE, 0, other text or a blank shows
        no format, so this is False.

        Args:
            at: The cell (``"D3"``), one of the range's.
            values: As :func:`evaluate` takes them.
            resolver: As :func:`evaluate` takes it.
            names: As :func:`evaluate` takes them.
            date1904: The 1904 date system.

        Raises:
            ValueError: ``at`` is not a cell of the range.
            UnsupportedError: The formula uses a function or form this package does not compute.
        """
        row, col = _cell(at)
        if not any(r1 <= row <= r2 and c1 <= col <= c2 for r1, c1, r2, c2 in self._areas):
            raise ValueError(f"{at} is not in the range {self.sqref}")

        def result(formula: Formula) -> Any:
            got = formula.evaluate(values, resolver=resolver, at=at, names=names, date1904=date1904)
            return _first(got)

        if self._own is None:
            return to_bool(result(self._formulas[0]), None) is True  # type: ignore[arg-type]
        own = result(self._own)
        bounds = [result(f) for f in self._formulas]
        if isinstance(own, ErrorValue) or any(isinstance(b, ErrorValue) for b in bounds):
            return False
        if self.operator in _RANGES:
            low, high = bounds
            if compare(low, high, ">") is True:
                low, high = high, low
            inside = compare(own, low, ">=") is True and compare(own, high, "<=") is True
            return inside if self.operator == "between" else not inside
        return compare(own, bounds[0], _OPERATORS[self.operator or ""]) is True


def _rule_formula(text: str, sqref: str, sheet: str | None, quirks: Quirks | None) -> Formula:
    reason = _refused(text)
    if reason is not None:
        raise FormulaError(
            f"a conditional format's formula holding {reason}: Excel offers to repair such a "
            f"file (capture cf_rules): {text!r}"
        )
    areas = _areas(sqref)
    anchor = f"{_letters(min(a[1] for a in areas))}{min(a[0] for a in areas)}"
    return Formula.parse(text, sheet=sheet, origin=anchor, quirks=quirks)
