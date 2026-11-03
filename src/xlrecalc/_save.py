"""The engine's values in the writer's file terms (ARCHITECTURE.md §9.2).

The writer in ``_xlsx`` knows nothing of the engine's types; this gives it, per sheet, a lookup
from a formula cell to its new cached value as the file stores it, or ``None`` for a cell whose
cached value must stay as it is: not a formula in the engine (a refused, array or data-table
formula, kept as its cached value), or uncomputed (§7.6).
"""

from __future__ import annotations

from xlrecalc._engine import Engine, unpack
from xlrecalc._report import column_letters
from xlrecalc._store import COLS, TemplateIds
from xlrecalc._values import BLANK, ErrorValue, Value
from xlrecalc._xlsx._write import Edit, FileValue, Lookup

__all__ = ["array_edits", "file_value", "lookups"]


def file_value(value: Value) -> FileValue | None:
    """A cell value as a file stores a formula's result; ``None`` for a blank."""
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, float):
        return ("n", value)
    if isinstance(value, str):
        return ("str", value)
    if isinstance(value, ErrorValue):
        return ("e", value.code)
    return None


def lookups(engine: Engine) -> dict[str, Lookup]:
    """Per sheet holding formulas, the lookup the writer asks for each formula cell."""
    out: dict[str, Lookup] = {}
    uncomputed: dict[str, set[tuple[int, int]]] = {}
    for cell in engine.report.uncomputed:
        uncomputed.setdefault(cell.sheet.casefold(), set()).add((cell.row, cell.col))
    for sheet in engine.sheets:
        if not sheet.templates:
            continue
        out[sheet.name] = _lookup(
            sheet.templates, sheet.values, uncomputed.get(sheet.name.casefold(), set())
        )
    return out


def _lookup(templates: TemplateIds, values: dict[int, Value], skip: set[tuple[int, int]]) -> Lookup:
    def lookup(row: int, col: int) -> FileValue | None:
        key = (row - 1) * COLS + col - 1
        if key not in templates or (row, col) in skip:
            return None
        value = values.get(key)
        return None if value is None else file_value(value)

    return lookup


def _a1(row: int, col: int) -> str:
    return f"{column_letters(col + 1)}{row + 1}"


def array_edits(engine: Engine) -> dict[str, dict[tuple[int, int], Edit]]:
    """The cells array formulas evaluated here fill, as edits (§7.7, §9.2).

    For each anchor evaluated since loading: the anchor with its range (``ref``) and value,
    every cell it fills, and every cell it filled before and fills no more (cleared).
    """
    out: dict[str, dict[tuple[int, int], Edit]] = {}
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
            cells[row + 1, col + 1] = Edit(text, None, ref, anchor.dynamic)
            continue
        cells[row + 1, col + 1] = Edit(
            text, file_value(s.values.get(row * COLS + col, BLANK)), ref, anchor.dynamic
        )
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) != (row, col):
                    value = file_value(s.values.get(r * COLS + c, BLANK))
                    cells[r + 1, c + 1] = Edit(None, value)
    for m in engine.released:
        if m in engine.members:
            continue
        sheet, row, col = unpack(m)
        s = engine.sheets[sheet]
        if row * COLS + col not in s.templates:
            out.setdefault(s.name, {}).setdefault((row + 1, col + 1), Edit(None, None))
    return out
