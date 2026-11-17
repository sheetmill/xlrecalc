"""The engine's values in the terms of xlio's writer (ARCHITECTURE.md §9.2).

``xlio``'s writer knows nothing of the engine's types; this gives it, per sheet, a lookup from
a formula cell to its new cached value as a Python value, or ``None`` for a cell whose cached
value must stay as it is: not a formula in the engine (a refused, array or data-table formula,
kept as its cached value), or uncomputed (§7.6).
"""

from __future__ import annotations

from collections.abc import Callable

from xlio import ArrayFill, CachedValues, CellEdit, CellValue, ErrorCode

from xlrecalc._engine import Engine, unpack
from xlrecalc._report import column_letters
from xlrecalc._store import COLS, Sheet, TemplateIds
from xlrecalc._values import BLANK, ErrorValue, Value

__all__ = ["array_edits", "file_value", "lookups"]


_CODES = {e.value: e for e in ErrorCode}


def file_value(value: Value) -> CellValue | None:
    """A cell value as the writer takes a formula's result; ``None`` for a blank.

    An error the file format has no code for (``#FIELD!``, ``#BLOCKED!``, ...) is ``#VALUE!``,
    as Excel caches the newer errors (capture ``arrays``); the writer writes ``#SPILL!`` and
    ``#CALC!`` as ``#VALUE!`` too. ``#GETTING_DATA`` is ``#N/A``, as Excel writes it back.
    """
    if isinstance(value, ErrorValue):
        if value.code == ErrorCode.GETTING_DATA:
            return ErrorCode.NA
        return _CODES.get(value.code, ErrorCode.VALUE)
    if isinstance(value, bool | float | str):
        return value
    return None


def lookups(engine: Engine) -> dict[str, CachedValues]:
    """Per sheet holding formulas, the lookup the writer asks for each formula cell."""
    out: dict[str, CachedValues] = {}
    uncomputed: dict[str, set[tuple[int, int]]] = {}
    for cell in engine.report.uncomputed:
        uncomputed.setdefault(cell.sheet.casefold(), set()).add((cell.row, cell.col))
    for name, row, col in engine.kept_with_values():  # the file holds it already (§7.2)
        uncomputed.setdefault(name.casefold(), set()).add((row, col))
    for name, row, col in engine.kept_values():  # likewise (§7.6)
        uncomputed.setdefault(name.casefold(), set()).add((row, col))
    for sheet in engine.sheets:
        if not sheet.templates:
            continue
        out[sheet.name] = _lookup(
            sheet.templates, sheet.values, uncomputed.get(sheet.name.casefold(), set())
        )
    return out


def _lookup(
    templates: TemplateIds, values: dict[int, Value], skip: set[tuple[int, int]]
) -> CachedValues:
    def lookup(row: int, col: int) -> CellValue | None:
        key = (row - 1) * COLS + col - 1
        if key not in templates or (row, col) in skip:
            return None
        value = values.get(key)
        return None if value is None else file_value(value)

    return lookup


def _a1(row: int, col: int) -> str:
    return f"{column_letters(col + 1)}{row + 1}"


def array_edits(
    engine: Engine,
) -> tuple[dict[str, dict[tuple[int, int], CellEdit]], dict[str, list[ArrayFill]]]:
    """The cells array formulas evaluated here fill, as edits and fills (§7.7, §9.2).

    For each anchor evaluated since loading: the anchor with its range (``ref``) and value as an
    edit, the other cells of its range as a fill (each value read from the store when written),
    and every cell it filled before and fills no more as an edit that clears it.
    """
    out: dict[str, dict[tuple[int, int], CellEdit]] = {}
    fills: dict[str, list[ArrayFill]] = {}
    uncomputed = {(c.sheet.casefold(), c.row, c.col) for c in engine.report.uncomputed}
    for n in engine.evaluated_anchors:
        anchor = engine.anchors.get(n)
        if anchor is None:
            continue
        sheet, row, col = unpack(n)
        s = engine.sheets[sheet]
        cells = out.setdefault(s.name, {})
        r1, c1, r2, c2 = anchor.box
        text = engine.formula_text(sheet, row, col)
        if text is None:
            continue
        ref = _a1(r1, c1) if (r1, c1) == (r2, c2) else f"{_a1(r1, c1)}:{_a1(r2, c2)}"
        if (s.name.casefold(), row + 1, col + 1) in uncomputed:
            cells[row + 1, col + 1] = CellEdit(text, None, ref, anchor.dynamic)
            continue
        cells[row + 1, col + 1] = CellEdit(
            text, file_value(s.values.get(row * COLS + col, BLANK)), ref, anchor.dynamic
        )
        if (r1, c1, r2, c2) != (row, col, row, col):
            fills.setdefault(s.name, []).append(
                ArrayFill(r1 + 1, c1 + 1, r2 + 1, c2 + 1, _reader(s))
            )
    for m in engine.released:
        if engine.member_of(m) is not None:
            continue
        sheet, row, col = unpack(m)
        s = engine.sheets[sheet]
        if row * COLS + col not in s.templates:
            out.setdefault(s.name, {}).setdefault((row + 1, col + 1), CellEdit(None, None))
    return out, fills


def _reader(sheet: Sheet) -> Callable[[int, int], CellValue | None]:
    values = sheet.values

    def value_at(row: int, col: int) -> CellValue | None:
        return file_value(values.get((row - 1) * COLS + col - 1, BLANK))

    return value_at
