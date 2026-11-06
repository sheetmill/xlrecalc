"""The evaluator: store, templates, scheduling, the evaluation context and the report (§7).

A cell is identified by one integer, ``sheet << 34 | row * 16384 + col`` (0-based). Formula
cells hold a template id; a template is compiled once (``_formula``) and evaluated as
``fn(engine, row, col)`` with the engine as the evaluation context. Recalculation is an
iterative depth-first walk over the stale cells in dependency order (§7.1); dependencies come
from the templates' references and the cell's position (§6). The dependents of a change come
from ``DependencyIndex`` (``_deps.py``), built when first needed after formulas change.
"""

from __future__ import annotations

import bisect
import datetime as dt
import math
import traceback
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import Any

from xlrecalc._arith import flush
from xlrecalc._arrays import (
    ArrayLimitError,
    NotImplementedFormError,
    check_area,
    element,
    is_array,
    shape,
    to_array,
)
from xlrecalc._coerce import intersect
from xlrecalc._deps import DependencyIndex
from xlrecalc._formula import MISSING, FormulaSyntaxError, TokenizeError, compile_formula, parse
from xlrecalc._formula import template_key as _template_key
from xlrecalc._formula._codegen import RefSpec
from xlrecalc._formula._fastkey import raw_key, translate
from xlrecalc._formula._template import Axis
from xlrecalc._functions import REGISTRY, SPECS
from xlrecalc._ops import Ops
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef, Report
from xlrecalc._store import COLS, MAX_COL, MAX_ROW, Sheet
from xlrecalc._values import BLANK, ErrorValue, Ref, Value

__all__ = ["Engine", "Report", "Template"]

SHEET_SHIFT = 34
KEY_MASK = (1 << SHEET_SHIFT) - 1
NAME, REF, VALUE = ErrorValue.NAME, ErrorValue.REF, ErrorValue.VALUE
SPILL = ErrorValue("#SPILL!")
_VOLATILE = frozenset(n for n, s in SPECS.items() if s.volatile)


def node(sheet: int, row: int, col: int) -> int:
    return sheet << SHEET_SHIFT | row * COLS + col


def unpack(n: int) -> tuple[int, int, int]:
    sheet = n >> SHEET_SHIFT
    row, col = divmod(n & KEY_MASK, COLS)
    return sheet, row, col


@dataclass(slots=True)
class Template:
    """A compiled formula shared by every cell whose formula has the same relative shape."""

    key: str
    fn: Any
    refs: tuple[RefSpec, ...]
    calls: frozenset[str]
    volatile: bool
    unsupported: frozenset[str]
    subtotal: bool
    origin: tuple[str, int, int] = ("", 0, 0)
    """The file-form text it was compiled from, and that formula's 0-based cell."""
    intersects: bool = False
    """Compiled as a legacy formula, it takes an implicit intersection somewhere (§7.7)."""
    array_mode: bool = False
    """Compiled as an array formula: ranges are arrays, the result may be an array."""


@dataclass
class Anchor:
    """A formula filling more than its cell: a legacy array formula or a dynamic array (§7.7).

    ``box`` (0-based r1, c1, r2, c2) is the range it fills: the declared range of a legacy array
    formula, the spill range of a dynamic array (its own cell alone when blocked); ``wanted``
    is the range a blocked dynamic array would spill into.
    """

    dynamic: bool
    box: tuple[int, int, int, int]
    wanted: tuple[int, int, int, int] | None = None


class _Cells:
    """What a ``Ref`` reads through: the store, without evaluating anything.

    ``limit`` is the largest array a reference may be made into (``ReadLimits.max_array_cells``).
    """

    __slots__ = ("limit", "sheets")

    def __init__(self, sheets: list[Sheet], limit: int | None = None) -> None:
        self.sheets = sheets
        self.limit = limit

    def get(self, sheet: int, row: int, col: int) -> Value:
        return self.sheets[sheet].values.get(row * COLS + col, BLANK)

    def box(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, Value]]:
        return self.sheets[sheet].box(r1, c1, r2, c2)

    def walk(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, Value]]:
        return self.sheets[sheet].box(r1, c1, r2, c2)

    def values(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> list[Value]:
        return self.sheets[sheet].box_values(r1, c1, r2, c2)


_LARGE_BOX = 256  # cells: smaller boxes are listed again rather than remembered


class _NotEvaluatedError(Exception):
    """A formula read a formula cell not evaluated yet (§6.3): evaluate that one first."""

    def __init__(self, n: int) -> None:
        super().__init__(n)
        self.node = n


class _CheckedCells(_Cells):
    """Reads through the store that stop at a cell not evaluated yet (``Engine._check``)."""

    __slots__ = ("engine",)

    def __init__(self, engine: Engine) -> None:
        super().__init__(engine.sheets, engine.max_array_cells)
        self.engine = engine

    def get(self, sheet: int, row: int, col: int) -> Value:
        self.engine._check(sheet, row, col)
        return super().get(sheet, row, col)

    def box(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, Value]]:
        self.engine._check_box(sheet, r1, c1, r2, c2)
        return super().box(sheet, r1, c1, r2, c2)

    def values(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> list[Value]:
        self.engine._check_box(sheet, r1, c1, r2, c2)
        return super().values(sheet, r1, c1, r2, c2)

    def walk(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, Value]]:
        self.engine._check_anchors(sheet, r1, c1, r2, c2)
        s = self.sheets[sheet]
        for r, c in s.positions(r1, c1, r2, c2):
            self.engine._check(sheet, r, c)
            v = s.values.get(r * COLS + c)
            if v is not None:
                yield r, c, v


class Engine:
    """A workbook's cells and formulas, evaluated as Excel does."""

    def __init__(
        self,
        quirks: Quirks | None = None,
        date1904: bool = False,
        today: dt.date | None = None,
        max_array_cells: int | None = 10_485_760,
    ) -> None:
        self.quirks = quirks or Quirks.EXCEL
        self.max_array_cells = max_array_cells
        self.reference_array: Any = None
        """An array of references made while the current cell is evaluated."""
        """The largest array a formula may make (``ReadLimits.max_array_cells``)."""
        self.date1904 = date1904
        self.today = today or dt.date.today()
        self.sheets: list[Sheet] = []
        self._sheet_index: dict[str, int] = {}
        self.templates: list[Template] = []
        self._template_ids: dict[str, int] = {}
        self._raw_ids: dict[str, int] = {}
        self.names: dict[tuple[int | None, str], str] = {}  # (sheet scope, folded name) -> text
        self._name_templates: dict[tuple[int | None, str], Template | None] = {}
        self.stale: set[int] = set()
        # Stale cells whose dependents may not be stale yet (uncached formulas, volatile ones),
        # repeats allowed: a list, as a file's every formula can be here at once. A change made
        # through set_value adds its dependents at once.
        self._unclosed: list[int] = []
        self.report = Report()
        self.cells = _Cells(self.sheets, max_array_cells)
        self.checked = _CheckedCells(self)
        self._strict = False  # every read checked (a formula on a static cycle, §6.3)
        self._volatile: set[int] = set()  # formula cells of volatile templates
        self._changed = False  # a cell changed since volatile cells were last made stale
        self.active = (0, 0, 0)
        """The active sheet and its active cell, 0-based: what ``CELL`` without a reference
        reads (capture ``cell_noref``)."""
        self.path: str | None = None
        """The workbook's file, for ``CELL("filename")``: an absolute path, or ``None``."""
        self.ops = Ops(self)
        # The cell being evaluated (the evaluation context of §7.3).
        self.sheet = self.row = self.col = 0
        self._index: DependencyIndex | None = None
        self.anchors: dict[int, Anchor] = {}
        """Formula cells that fill a range (§7.7), by node."""
        # Per sheet, rebuilt when first needed after anchors change: the ranges anchors fill or
        # want, (r1, c1, r2, c2, anchor); and per column the row spans they fill, sorted, for
        # member_of. A cell an anchor fills is found from its range, not stored per cell.
        self._anchor_boxes: dict[int, list[tuple[int, int, int, int, int]]] = {}
        self._member_spans: dict[int, dict[int, tuple[list[int], list[tuple[int, int, int]]]]] = {}
        self._anchors_dirty: set[int] = set()
        self._moved: list[tuple[int, int, int, int, int]] = []
        self.evaluated_anchors: set[int] = set()
        """Anchors evaluated since loading: the writer rewrites their ranges."""
        self.released: set[int] = set()
        """Cells an anchor filled and fills no more: the writer clears them."""
        self.locked: dict[int, list[tuple[int, int, int, int]]] = {}
        """Per sheet, the 0-based boxes (r1, c1, r2, c2) of array formulas and data tables,
        whose cells cannot be changed one by one."""

    # Building ---------------------------------------------------------------------------------

    @classmethod
    def from_cells(cls, sheets: dict[str, dict[str, Any]], **options: Any) -> Engine:
        """An engine from ``{"Sheet1": {"A1": 1.0, "B1": "=A1*2"}}``, for tests and examples.

        A text starting with ``=`` is a formula (display form); other values are constants.
        """
        from xlrecalc._formula._template import cell_axes

        engine = cls(**options)
        for name in sheets:
            engine.add_sheet(name)
        for name, cells in sheets.items():
            s = engine.sheet_index(name)
            assert s is not None
            for ref, value in cells.items():
                r_axis, c_axis = cell_axes(ref, 0, 0)
                row, col = r_axis.value, c_axis.value
                if isinstance(value, str) and value.startswith("="):
                    engine.set_formula(s, row, col, value[1:])
                else:
                    if isinstance(value, int) and not isinstance(value, bool):
                        value = float(value)
                    engine.sheets[s].set(row, col, BLANK if value is None else value)
        return engine

    def add_sheet(self, name: str) -> int:
        index = len(self.sheets)
        self.sheets.append(Sheet(name))
        self._sheet_index[name.casefold()] = index
        return index

    def sheet_index(self, name: str) -> int | None:
        return self._sheet_index.get(name.casefold())

    def define_name(self, name: str, text: str, scope: int | None = None) -> None:
        self.names[(scope, name.casefold())] = text
        self._name_templates.clear()
        self._index = None

    def template_for(
        self, text: str, row: int, col: int, raw: str | None = None, array_mode: bool = False
    ) -> int:
        """The id of the template of a formula (file form) in (row, col); raises on bad syntax.

        The raw-text key (§5.3) finds a known template without parsing; a new key is parsed.
        ``array_mode`` compiles an array formula (§7.7), a template of its own.
        """
        if raw is None:
            raw = raw_key(text, row, col)
        if array_mode:
            raw = "\x1dA" + raw
        tid = self._raw_ids.get(raw)
        if tid is None:
            tid = self._parsed_template(text, row, col, array_mode)
            self._raw_ids[raw] = tid
        return tid

    def _parsed_template(self, text: str, row: int, col: int, array_mode: bool = False) -> int:
        formula = parse(text)
        key = _template_key(formula, row, col)
        if array_mode:
            key = "{array}" + key
        tid = self._template_ids.get(key)
        if tid is None:
            tid = self._compile(formula, key, row, col, array_mode)
        return tid

    def _compile(self, formula: Any, key: str, row: int, col: int, array_mode: bool = False) -> int:
        compiled = compile_formula(formula, row, col, REGISTRY, self.ops, self.quirks, array_mode)
        unsupported = {c for c in compiled.calls if c not in REGISTRY} | compiled.features
        template = Template(
            key,
            compiled.fn,
            compiled.refs,
            compiled.calls,
            bool(compiled.calls & _VOLATILE),
            frozenset(unsupported),
            "SUBTOTAL" in compiled.calls,
            (formula.text(), row, col),
            compiled.intersects,
            array_mode,
        )
        tid = len(self.templates)
        self.templates.append(template)
        self._template_ids[key] = tid
        return tid

    def formula_text(self, sheet: int, row: int, col: int) -> str | None:
        """The cell's formula in file form, or ``None`` for a cell without one."""
        tid = self.sheets[sheet].templates.at(row, col)
        if tid is None:
            return None
        text, r0, c0 = self.templates[tid].origin
        return text if (r0, c0) == (row, col) else translate(text, r0, c0, row, col)

    def set_value(self, sheet: int, row: int, col: int, value: Value) -> None:
        """A constant; the cell's formula, if any, is removed and its dependents become stale."""
        s = self.sheets[sheet]
        n = node(sheet, row, col)
        if n in self.anchors:
            self._drop_anchor(n)
        if row * COLS + col in s.templates:
            s.set_template(row, col, None)
            self._index = None
        self._touch_spill(sheet, row, col)
        s.set(row, col, value)
        self._volatile.discard(n)
        self._changed = True
        self.stale.discard(n)
        self.stale |= self.dependents([n])

    def set_formula(
        self,
        sheet: int,
        row: int,
        col: int,
        text: str | None = None,
        *,
        template: int | None = None,
        cached: Value | None = None,
        array: tuple[bool, tuple[int, int, int, int]] | None = None,
    ) -> None:
        """A formula (file form) or an existing template; ``cached`` is its trusted value.

        ``array`` makes it an array formula: ``(dynamic, box)``, a dynamic array with its
        current spill range or a legacy array formula with its declared range (0-based).
        """
        if template is not None:
            tid = template
        else:
            tid = self.template_for(text or "", row, col, array_mode=array is not None)
        s = self.sheets[sheet]
        n = node(sheet, row, col)
        if n in self.anchors:
            self._drop_anchor(n)
        self._touch_spill(sheet, row, col)
        s.set_template(row, col, tid)
        self._index = None
        self._changed = True
        if self.templates[tid].volatile:
            self._volatile.add(n)
        else:
            self._volatile.discard(n)
        if array is not None:
            self.anchors[n] = Anchor(array[0], array[1])
            self._index_anchors(sheet)
        if cached is None:
            s.set(row, col, BLANK)
            self.stale.add(n)
            self._unclosed.append(n)
        else:
            s.set(row, col, cached)
            if self.templates[tid].volatile:
                self.stale.add(n)
                self._unclosed.append(n)

    def set_member(self, sheet: int, row: int, col: int, anchor: int, value: Value | None) -> None:
        """A cell an anchor fills, holding the value the file caches for it."""
        self.sheets[sheet].set(row, col, BLANK if value is None else value)

    # Arrays (§7.7) ----------------------------------------------------------------------------

    def _index_anchors(self, sheet: int) -> None:
        self._anchors_dirty.add(sheet)

    def _anchor_list(self, sheet: int) -> list[tuple[int, int, int, int, int]]:
        """The ranges the sheet's anchors fill or want, as (r1, c1, r2, c2, anchor)."""
        if sheet in self._anchors_dirty:
            self._rebuild_anchors(sheet)
        return self._anchor_boxes.get(sheet, [])

    def _rebuild_anchors(self, sheet: int) -> None:
        self._anchors_dirty.discard(sheet)
        boxes: list[tuple[int, int, int, int, int]] = []
        columns: dict[int, list[tuple[int, int, int]]] = {}
        for n, a in self.anchors.items():
            if n >> SHEET_SHIFT != sheet:
                continue
            boxes.append((*a.box, n))
            if a.wanted is not None:
                boxes.append((*a.wanted, n))
            r1, c1, r2, c2 = a.box
            for c in range(c1, c2 + 1):
                columns.setdefault(c, []).append((r1, r2, n))
        spans = {}
        for c, found in columns.items():
            found.sort()
            spans[c] = ([f[0] for f in found], found)
        self._anchor_boxes[sheet] = boxes
        self._member_spans[sheet] = spans

    def member_of(self, n: int) -> int | None:
        """The anchor whose range holds this cell (the anchor itself excluded), or ``None``."""
        if not self.anchors:
            return None
        sheet, row, col = unpack(n)
        if sheet in self._anchors_dirty:
            self._rebuild_anchors(sheet)
        found = self._member_spans.get(sheet, {}).get(col)
        if not found:
            return None
        starts, spans = found
        i = bisect.bisect_right(starts, row) - 1
        if i < 0:
            return None
        _, last, anchor = spans[i]
        return anchor if row <= last and anchor != n else None

    def _release(
        self, sheet: int, box: tuple[int, int, int, int], keep: tuple[int, ...]
    ) -> list[tuple[int, int, int, int]]:
        """Clear the cells of ``box`` outside ``keep`` (r1, c1, r2, c2) that hold a value.

        Returns the changed runs, (sheet, col, first row, last row), for their dependents.
        """
        s = self.sheets[sheet]
        values = s.values
        for c in range(box[1], box[3] + 1):
            rows = s.columns.get(c)
            if not rows:
                continue
            lo, hi = bisect.bisect_left(rows, box[0]), bisect.bisect_right(rows, box[2])
            for r in list(rows[lo:hi]):
                inside = bool(keep) and keep[0] <= r <= keep[2] and keep[1] <= c <= keep[3]
                if not inside and r * COLS + c in values:
                    s.set(r, c, BLANK)
                    self.released.add(node(sheet, r, c))
        return [(sheet, c, box[0], box[2]) for c in range(box[1], box[3] + 1)]

    def _drop_anchor(self, n: int) -> None:
        anchor = self.anchors.pop(n)
        sheet, row, col = unpack(n)
        changed = self._release(sheet, anchor.box, (row, col, row, col))
        self._index_anchors(sheet)
        self.stale |= self.dependents([], changed)

    def _touch_spill(self, sheet: int, row: int, col: int) -> None:
        """A cell set by hand: a dynamic array spilling there, or wanting to, is evaluated again."""
        for r1, c1, r2, c2, a in self._anchor_list(sheet):
            inside = r1 <= row <= r2 and c1 <= col <= c2 and node(sheet, row, col) != a
            if inside and self.anchors[a].dynamic:
                self.stale.add(a)
                self._unclosed.append(a)

    def _blocked(self, sheet: int, box: tuple[int, int, int, int], anchor: int) -> bool:
        s = self.sheets[sheet]
        r1, c1, r2, c2 = box
        if r2 > MAX_ROW or c2 > MAX_COL:
            return True
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                key = r * COLS + c
                m = node(sheet, r, c)
                if m == anchor:
                    continue
                owner = self.member_of(m)
                if owner == anchor:
                    continue
                if key in s.templates or s.values.get(key, BLANK) is not BLANK or owner is not None:
                    return True
        return False

    def _fill(self, n: int, value: Any) -> None:
        """An anchor's result spread over its range: the anchor and its members' values."""
        anchor = self.anchors[n]
        sheet, row, col = unpack(n)
        s = self.sheets[sheet]
        self.evaluated_anchors.add(n)
        old = anchor.box
        grid = value if is_array(value) else ((value,),)
        rows, cols = shape(grid)
        if not anchor.dynamic:  # a legacy array fills its declared range from the result
            box = old
            check_area(box[2] - box[0] + 1, box[3] - box[1] + 1, self.max_array_cells)
        else:
            box = (row, col, row + rows - 1, col + cols - 1)
            anchor.wanted = None
            if box != (row, col, row, col) and self._blocked(sheet, box, n):
                anchor.wanted = box
                grid, rows, cols, box = ((SPILL,),), 1, 1, (row, col, row, col)
        if box != old:  # cells no longer filled
            self._release(sheet, old, box)
        released = self.released
        for i in range(box[2] - box[0] + 1):  # element by element: no grid of the range
            r = box[0] + i
            for j in range(box[3] - box[1] + 1):
                c = box[1] + j
                v = element(grid, rows, cols, i, j)
                if isinstance(v, Ref) and v.source is self.checked:
                    v = ErrorValue.VALUE  # OFFSET or INDIRECT of an array (capture array_refs)
                s.set(r, c, self._result(v))
                if released:
                    released.discard(node(sheet, r, c))
        if box != old or anchor.wanted is not None:
            anchor.box = box
            self._index_anchors(sheet)
        if box != old:
            self._moved.append(
                (
                    sheet,
                    min(old[0], box[0]),
                    min(old[1], box[1]),
                    max(old[2], box[2]),
                    max(old[3], box[3]),
                )
            )

    def arr(self, v: Any) -> Any:
        """A value for an array context: a range's values as an array (§7.7)."""
        if isinstance(v, Ref):
            return v.value(0, 0) if v.is_cell() else to_array(v)
        return v

    def spill_range(self, ref: Any) -> Any:
        """``A1#`` (``_xlfn.ANCHORARRAY(A1)``): the range the dynamic array at A1 spills into."""
        if not isinstance(ref, Ref) or not ref.is_cell():
            return REF
        anchor = self.anchors.get(node(ref.sheet, ref.r1, ref.c1))
        if anchor is None or not anchor.dynamic:
            return REF
        if anchor.wanted is not None:
            return SPILL
        r1, c1, r2, c2 = anchor.box
        return Ref(self.cells, ref.sheet, r1, c1, r2, c2)

    def _settle(self) -> None:
        """Evaluate again what depends on cells a spill reached or left (bounded)."""
        for _ in range(100):
            if not self._moved:
                return
            moved, self._moved = self._moved, []
            ranges = [
                (sheet, c, r1, r2) for sheet, r1, c1, r2, c2 in moved for c in range(c1, c2 + 1)
            ]
            self.stale |= self.dependents([], ranges)
            self._evaluate(list(self.stale))

    # Values -----------------------------------------------------------------------------------

    def get(self, sheet: int, row: int, col: int) -> Value:
        """The cell's current value, evaluating it first if it is stale."""
        n = node(sheet, row, col)
        if self._unclosed or self._changed:
            self._close()
        anchor = self.member_of(n)
        if anchor is not None and anchor in self.stale:
            self.recalculate(targets=[anchor])
        if n in self.stale:
            self.recalculate(targets=[n])
        return self.sheets[sheet].values.get(row * COLS + col, BLANK)

    def recalculate(self, full: bool = False, targets: Iterable[int] | None = None) -> Report:
        """Evaluate stale cells (every formula with ``full``) in dependency order."""
        if full:
            for s, sheet in enumerate(self.sheets):
                self.stale.update((s << SHEET_SHIFT) | k for k in sheet.templates)
            self._unclosed.clear()
        if targets is None:
            self._changed = True  # a recalculation evaluates volatile formulas again
            self._close()
            if 2 * len(self.stale) >= self.formula_count():
                self._evaluate(self._stale_in_store_order())
            self._evaluate(list(self.stale))
        else:
            self._evaluate(list(targets))
        self._settle()
        self._taint()
        if not self.stale:
            self.stale.clear()  # a set keeps its table when emptied; clearing frees it
        return self.report

    def _stale_in_store_order(self) -> Iterator[int]:
        """Stale formula cells by sheet, column and row, read from the store as evaluated.

        For a recalculation of most of the workbook: no copy of the stale set, which may hold
        millions of cells. What is left stale afterwards (if anything) is evaluated from a copy.
        """
        stale = self.stale
        for s, sheet in enumerate(self.sheets):
            base = s << SHEET_SHIFT
            for col in sorted(sheet.formula_columns):
                for row in sheet.formula_columns[col]:
                    n = base | (row * COLS + col)
                    if n in stale:
                        yield n

    def _close(self) -> None:
        """Make the dependents of new or changed formulas stale (unless every formula is).

        After any change, volatile formulas (``NOW``, ``OFFSET``, ``INDIRECT``, ``CELL``...)
        are stale too, as Excel recalculates them on every recalculation.
        """
        if self._changed and self._volatile:
            self.stale |= self._volatile
            self._unclosed.extend(self._volatile)
        self._changed = False
        if self._unclosed and len(self.stale) < self.formula_count():
            self.stale |= self.dependents(self._unclosed)
        self._unclosed.clear()

    def formula_count(self) -> int:
        return sum(len(s.templates) for s in self.sheets)

    def _evaluate(self, targets: Iterable[int]) -> None:
        """Evaluate stale cells in dependency order.

        The order is a depth-first post-order over the static precedents (§7.1), corrected
        during evaluation (§6.3).
        A formula that reads a formula cell not evaluated yet (through ``OFFSET``,
        ``INDIRECT``, or a static cycle that evaluation may not follow) stops with
        ``_NotEvaluatedError``: that cell is evaluated first, and the formula again. A circular
        reference is a chain of cells each stopped by reading the next that closes on itself; the
        static path alone does not show one, since a formula may not follow a reference.
        """
        pending = self.stale
        on_stack: dict[int, int] = {}  # node -> its position on the current path
        path: list[int] = []
        cyclic: set[int] = set()
        suspects: set[int] = set()  # on a static cycle: evaluated with every read checked
        marked_from = marked_to = 0  # path[marked_from:marked_to] is all in suspects
        clean: set[tuple[int, int, int, int, int]] = set()  # large boxes with nothing stale left
        live: dict[tuple[int, int, int], tuple[RefSpec, ...]] = {}
        waits: dict[int, int] = {}  # a cell stopped by a read, and the cell it read
        for target in targets:  # one at a time: no stack entry per stale cell up front
            if target not in pending:
                continue
            stack: list[tuple[int, bool]] = [(target, False)]
            while stack:
                n, expanded = stack.pop()
                if n not in pending:
                    continue
                if expanded:
                    self._strict = n in suspects
                    try:
                        self._compute(n)
                    except _NotEvaluatedError as wait:
                        t = wait.node
                        waits[n] = t
                        chain = {n}  # who waits for whose value: a real cycle closes this
                        x = t
                        while x != n and x in waits and x not in chain:
                            chain.add(x)
                            x = waits[x]
                        if x != n:  # evaluate what it reads first, then it again
                            stack.append((n, True))
                            if t in on_stack:  # on the path by a reference it may not follow
                                suspects.add(t)
                                stack.append((t, True))
                            else:
                                stack.append((t, False))
                            continue
                        cyclic.update(chain)
                    finally:
                        self._strict = False
                    waits.pop(n, None)
                    del on_stack[n]
                    while path and path[-1] not in on_stack:  # n, and any evaluated ahead of it
                        path.pop()
                    marked_to = min(marked_to, len(path))
                    pending.discard(n)
                    continue
                if n in on_stack:
                    continue
                on_stack[n] = len(path)
                path.append(n)
                stack.append((n, True))
                for p in self._stale_precedents(n, pending, clean, live):
                    if p not in on_stack:
                        stack.append((p, False))
                        continue
                    first, last = on_stack[p], len(path)  # marked once: a path may be long
                    if first < marked_from or first > marked_to:
                        suspects.update(path[first:last])
                        marked_from = first
                    else:
                        suspects.update(path[marked_to:last])
                    marked_to = last
        if cyclic:
            self._circular(cyclic)

    def _check(self, sheet: int, row: int, col: int) -> None:
        """Stop the formula being evaluated if this cell is not evaluated yet (§6.3)."""
        n = node(sheet, row, col)
        stale = self.stale
        if n in stale and self.sheets[sheet].templates.at(row, col) is not None:
            raise _NotEvaluatedError(n)
        anchor = self.member_of(n)
        if anchor is not None and anchor in stale:
            raise _NotEvaluatedError(anchor)

    def _check_box(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> None:
        stale = self.stale
        for r, c in self.sheets[sheet].formulas_in(r1, c1, r2, c2):
            if node(sheet, r, c) in stale:
                raise _NotEvaluatedError(node(sheet, r, c))
        self._check_anchors(sheet, r1, c1, r2, c2)

    def _check_anchors(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> None:
        stale = self.stale
        for a1, b1, a2, b2, anchor in self._anchor_list(sheet):
            if anchor in stale and a1 <= r2 and r1 <= a2 and b1 <= c2 and c1 <= b2:
                raise _NotEvaluatedError(anchor)

    def dynamic_ref(self, sheet: int, r1: int, c1: int, r2: int, c2: int) -> Ref:
        """A reference made during evaluation (``OFFSET``, ``INDIRECT``): its reads are checked."""
        return Ref(self.checked, sheet, r1, c1, r2, c2)

    def _circular(self, cyclic: set[int]) -> None:
        """Cells on a cycle and every cell depending on one: Excel leaves them uncalculated.

        With iteration off, Excel calculates none of them and saves each as the empty text
        (``t="str"`` with an empty ``<v/>``; capture ``circular``), so that is their value.
        """
        for n in cyclic | self.dependents(cyclic):
            sheet, row, col = unpack(n)
            s = self.sheets[sheet]
            if row * COLS + col not in s.templates:
                continue
            s.values[row * COLS + col] = ""
            self.stale.discard(n)  # final: a later read must not compute it again
            self.report.circular.add(CellRef(s.name, row + 1, col + 1))

    def _compute(self, n: int) -> None:
        sheet, row, col = unpack(n)
        s = self.sheets[sheet]
        key = row * COLS + col
        tid = s.templates.at(row, col)
        if tid is None:
            return
        template = self.templates[tid]
        where = CellRef(s.name, row + 1, col + 1)
        if template.unsupported:
            self.report.unsupported.update(template.unsupported)
            if key not in s.values:  # no cached value: a formula cell is never blank otherwise
                s.values[key] = NAME
                self.report.uncomputed.add(where)
            return
        self.sheet, self.row, self.col = sheet, row, col
        self.reference_array = None
        try:
            value = template.fn(self, row, col)
            if self.reference_array is not None and value is not self.reference_array:
                raise NotImplementedFormError("an array of references read by a function")
            if n in self.anchors:
                self._fill(n, value)
                return
        except _NotEvaluatedError:
            raise
        except ArrayLimitError as exc:  # keep the cached value, as for a function not supported
            self.report.unsupported[f"an array of more than {exc.limit:,} cells"] += 1
            self.report.uncomputed.add(where)
            return
        except NotImplementedFormError as exc:  # keep the cached value, likewise
            self.report.unsupported[exc.label] += 1
            self.report.uncomputed.add(where)
            return
        except Exception as exc:  # a bug in a function must not stop the recalculation
            value = VALUE
            self.report.exceptions[where] = "".join(
                traceback.format_exception_only(type(exc), exc)
            ).strip()
            self.report.uncomputed.add(where)
            if n in self.anchors:
                self._fill(n, value)
                return
        value = self._result(value)
        if value is BLANK:
            s.values.pop(key, None)
        else:
            s.values[key] = value

    def _result(self, value: Any) -> Value:
        """A formula's result as a cell holds it: references intersected, blanks as 0."""
        if isinstance(value, tuple):  # an array in one cell shows its first element
            value = value[0][0]
        if isinstance(value, Ref):  # the first of an array of references too
            value = intersect(value, self)
        if value is BLANK or value is MISSING:
            return 0.0
        if isinstance(value, float):
            return flush(value) if math.isfinite(value) else ErrorValue.NUM
        return value  # type: ignore[no-any-return]

    def _taint(self) -> None:
        """Cells depending on an uncomputed cell are uncomputed too."""
        if not self.report.uncomputed:
            return
        seeds = [
            node(self._sheet_index[name.casefold()], r - 1, c - 1)
            for name, r, c in self.report.uncomputed
        ]
        for n in self.dependents(seeds):
            sheet, row, col = unpack(n)
            self.report.uncomputed.add(CellRef(self.sheets[sheet].name, row + 1, col + 1))

    # Dependencies -------------------------------------------------------------------------------

    def _boxes(self, n: int) -> Iterator[tuple[int, int, int, int, int]]:
        """The boxes (sheet, r1, c1, r2, c2) a formula cell's template refers to."""
        sheet, row, col = unpack(n)
        tid = self.sheets[sheet].templates.at(row, col)
        if tid is None:
            return
        for spec in self.templates[tid].refs:
            yield from self._spec_boxes(spec, sheet, row, col)

    def _spec_boxes(
        self, spec: RefSpec, sheet: int, row: int, col: int
    ) -> Iterator[tuple[int, int, int, int, int]]:
        if spec.kind == "name":
            text = self._name_text(spec.name or "", sheet)
            if text is None:
                return
            template = self._name_template(spec.name or "", sheet)
            if template is not None:
                for inner in template.refs:
                    if inner.kind == "box":
                        yield from self._spec_boxes(inner, sheet, 0, 0)
            return
        target = sheet if spec.sheet is None else self.sheet_index(spec.sheet)
        if target is None:
            return
        r1, c1 = _at(spec.r1, row), _at(spec.c1, col)
        r2, c2 = _at(spec.r2, row), _at(spec.c2, col)
        r1, r2 = max(0, min(r1, r2)), min(MAX_ROW, max(r1, r2))
        c1, c2 = max(0, min(c1, c2)), min(MAX_COL, max(c1, c2))
        yield target, r1, c1, r2, c2

    def precedents(self, n: int) -> Iterator[int]:
        """Formula cells a formula cell reads: anchors for the cells they fill (§7.7)."""
        for sheet, r1, c1, r2, c2 in self._boxes(n):
            s = self.sheets[sheet]
            if r1 == r2 and c1 == c2:
                if s.templates.at(r1, c1) is not None:
                    yield node(sheet, r1, c1)
            else:
                for r, c in s.formulas_in(r1, c1, r2, c2):
                    yield node(sheet, r, c)
            for a1, b1, a2, b2, anchor in self._anchor_list(sheet) if self.anchors else ():
                if anchor != n and a1 <= r2 and r1 <= a2 and b1 <= c2 and c1 <= b2:
                    yield anchor

    def _live_specs(self, tid: int, sheet: int, col: int) -> tuple[RefSpec, ...]:
        """The references of a template in a column that can reach a formula cell.

        A reference to columns holding no formula has no precedent, whatever the row.
        """
        out = []
        for spec in self.templates[tid].refs:
            target = sheet if spec.sheet is None else self.sheet_index(spec.sheet)
            if spec.kind == "box" and not self.anchors:
                if target is None:
                    continue
                c1, c2 = _at(spec.c1, col), _at(spec.c2, col)
                c1, c2 = max(0, min(c1, c2)), min(MAX_COL, max(c1, c2))
                columns = self.sheets[target].formula_columns
                candidates = range(c1, c2 + 1) if c2 - c1 < len(columns) else columns
                if not any(c1 <= c <= c2 and columns.get(c) for c in candidates):
                    continue
            out.append(spec)
        return tuple(out)

    def _stale_precedents(
        self,
        n: int,
        stale: set[int],
        clean: set[tuple[int, int, int, int, int]],
        live: dict[tuple[int, int, int], tuple[RefSpec, ...]],
    ) -> Iterator[int]:
        """``precedents`` that are stale, for one pass of the scheduler.

        Nothing becomes stale during a pass, and no formula appears, so a large box found with
        nothing stale (``clean``) is not listed again by the many formulas that may refer to it,
        and a template's references to columns of no formula are dropped once for all its cells
        (``live``).
        """
        at_sheet, row, col = unpack(n)
        tid = self.sheets[at_sheet].templates.at(row, col)
        if tid is None:
            return
        key = (tid, at_sheet, col)
        specs = live.get(key)
        if specs is None:
            specs = live[key] = self._live_specs(tid, at_sheet, col)
        boxes = (box for spec in specs for box in self._spec_boxes(spec, at_sheet, row, col))
        for sheet, r1, c1, r2, c2 in boxes:
            s = self.sheets[sheet]
            if r1 == r2 and c1 == c2:
                m = node(sheet, r1, c1)
                if m in stale and s.templates.at(r1, c1) is not None:
                    yield m
                found = False
            else:
                box = (sheet, r1, c1, r2, c2)
                if box in clean:
                    continue
                found = False
                for r, c in s.formulas_in(r1, c1, r2, c2):
                    m = node(sheet, r, c)
                    if m in stale:
                        found = True
                        yield m
            for a1, b1, a2, b2, anchor in self._anchor_list(sheet) if self.anchors else ():
                if anchor in stale and a1 <= r2 and r1 <= a2 and b1 <= c2 and c1 <= b2:
                    if anchor != n:
                        found = True
                        yield anchor
                    else:
                        found = True
            if not found and (r2 - r1 + 1) * (c2 - c1 + 1) >= _LARGE_BOX:
                clean.add((sheet, r1, c1, r2, c2))

    def index(self) -> DependencyIndex:
        """The reverse-dependency index, rebuilt after formulas change."""
        if self._index is None:
            self._index = DependencyIndex(self)
        return self._index

    def dirty_runs(
        self, cells: Iterable[int], ranges: Iterable[tuple[int, int, int, int]] = ()
    ) -> dict[tuple[int, int], list[tuple[int, int]]]:
        """Every formula cell depending on the cells, as row runs per (sheet, col)."""
        return self.index().dirty(cells, ranges)

    def dependents(
        self, cells: Iterable[int], ranges: Iterable[tuple[int, int, int, int]] = ()
    ) -> set[int]:
        """Every formula cell that depends, directly or not, on the cells and the ranges.

        ``ranges`` are runs of changed cells, as (sheet, col, first row, last row). An anchor
        that depends on them changes the cells it fills, so their dependents count too (§7.7).
        """
        seeds = list(cells)
        found = self._dependents(seeds, ranges)
        if not self.anchors:
            return found
        expanded: set[int] = set()
        frontier = set(seeds) | found
        while True:
            extra: list[tuple[int, int, int, int]] = []
            for a in frontier & self.anchors.keys():
                if a in expanded:
                    continue
                expanded.add(a)
                sheet, _, _ = unpack(a)
                anchor = self.anchors[a]
                for box in (anchor.box, anchor.wanted):
                    if box is not None:
                        extra += [(sheet, c, box[0], box[2]) for c in range(box[1], box[3] + 1)]
            if not extra:
                return found
            more = self._dependents([], extra) - found
            found |= more
            frontier = more

    def _dependents(
        self, cells: Iterable[int], ranges: Iterable[tuple[int, int, int, int]] = ()
    ) -> set[int]:
        found: set[int] = set()
        for (sheet, col), runs in self.dirty_runs(cells, ranges).items():
            base = sheet << SHEET_SHIFT
            for first, last in runs:
                found.update(base | (r * COLS + col) for r in range(first, last + 1))
        return found

    # The evaluation context the generated code calls (§7.3) ------------------------------------

    def _target(self, sheet: str | None) -> int | ErrorValue:
        if sheet is None:
            return self.sheet
        if "[" in sheet or ":" in sheet:
            return self.unsupported("external or 3-D reference")  # type: ignore[return-value]
        index = self.sheet_index(sheet)
        return REF if index is None else index

    def cell(self, sheet: str | None, row: int, col: int) -> Value:
        target = self._target(sheet)
        if isinstance(target, ErrorValue):
            return target
        if not (0 <= row <= MAX_ROW and 0 <= col <= MAX_COL):
            return REF
        if self._strict:
            self._check(target, row, col)
        return self.sheets[target].values.get(row * COLS + col, BLANK)

    def cref(self, sheet: str | None, row: int, col: int) -> Ref | ErrorValue:
        return self.box(sheet, row, col, row, col)

    def box(self, sheet: str | None, r1: int, c1: int, r2: int, c2: int) -> Ref | ErrorValue:
        target = self._target(sheet)
        if isinstance(target, ErrorValue):
            return target
        if min(r1, r2, c1, c2) < 0 or max(r1, r2) > MAX_ROW or max(c1, c2) > MAX_COL:
            return REF
        return Ref(self.checked if self._strict else self.cells, target, r1, c1, r2, c2)

    def columns(self, sheet: str | None, c1: int, c2: int) -> Ref | ErrorValue:
        return self.box(sheet, 0, c1, MAX_ROW, c2)

    def rows(self, sheet: str | None, r1: int, r2: int) -> Ref | ErrorValue:
        return self.box(sheet, r1, 0, r2, MAX_COL)

    def _name_text(self, name: str, sheet: int) -> str | None:
        folded = name.casefold()
        return self.names.get((sheet, folded), self.names.get((None, folded)))

    def _name_template(self, name: str, sheet: int) -> Template | None:
        key = (sheet, name.casefold())
        if key not in self._name_templates:
            text = self._name_text(name, sheet)
            template = None
            if text is not None:
                try:
                    template = self.templates[self.template_for(text, 0, 0)]
                except (FormulaSyntaxError, TokenizeError):
                    template = None
            self._name_templates[key] = template
        return self._name_templates[key]

    def name_text(self, name: str) -> str | None:
        """A defined name's definition as the current sheet sees it, or ``None``."""
        return self._name_text(name, self.sheet)

    def name(self, sheet: str | None, name: str) -> Any:
        scope = self.sheet if sheet is None else self.sheet_index(sheet)
        if scope is None:
            return REF
        template = self._name_template(name, scope)
        if template is None:
            return NAME
        if template.unsupported:
            return self.unsupported(sorted(template.unsupported)[0])
        saved = self.sheet, self.row, self.col
        self.sheet = scope
        try:
            return template.fn(self, 0, 0)
        finally:
            self.sheet, self.row, self.col = saved

    def structured(self, text: str) -> Value:
        return self.unsupported("structured reference")

    def spill(self, value: Any) -> Value:
        return self.unsupported("spill reference")

    def unsupported(self, label: str, *args: Any) -> Value:
        self.report.unsupported[label] += 1
        s = self.sheets[self.sheet]
        self.report.uncomputed.add(CellRef(s.name, self.row + 1, self.col + 1))
        return NAME

    def is_subtotal(self, sheet: int, row: int, col: int) -> bool:
        tid = self.sheets[sheet].templates.at(row, col)
        return tid is not None and self.templates[tid].subtotal


def _at(axis: Axis | None, base: int) -> int:
    assert axis is not None
    return axis.value if axis.absolute else base + axis.value
