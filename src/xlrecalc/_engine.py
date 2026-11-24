"""The evaluator: store, templates, scheduling, evaluation context and report (ARCHITECTURE.md §7).

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
from collections import defaultdict
from collections.abc import Callable, Iterable, Iterator, Sequence
from dataclasses import dataclass, replace
from functools import cmp_to_key
from typing import Any

from xlrecalc._arith import flush
from xlrecalc._arrays import (
    ArrayLimitError,
    NotImplementedFormError,
    Padded,
    check_area,
    element,
    is_array,
    shape,
    to_array,
)
from xlrecalc._coerce import intersect
from xlrecalc._compare import compare
from xlrecalc._deps import DependencyIndex, Runs
from xlrecalc._formula import MISSING, FormulaSyntaxError, TokenizeError, compile_formula, parse
from xlrecalc._formula import template_key as _template_key
from xlrecalc._formula._codegen import EXTERNAL, RefSpec
from xlrecalc._formula._fastkey import raw_key, translate
from xlrecalc._formula._prefixes import PREFIXES
from xlrecalc._formula._template import Axis
from xlrecalc._functions import REGISTRY
from xlrecalc._limits import StepLimitError
from xlrecalc._opening import Volatility
from xlrecalc._ops import Ops
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef, CellStatus, Report
from xlrecalc._store import COLS, MAX_COL, MAX_ROW, Sheet
from xlrecalc._values import BLANK, ErrorValue, Lambda, Ref, Value

__all__ = ["Engine", "Report", "Template"]

SHEET_SHIFT = 34
KEY_MASK = (1 << SHEET_SHIFT) - 1
NAME, REF, VALUE = ErrorValue.NAME, ErrorValue.REF, ErrorValue.VALUE
SPILL, CALC = ErrorValue("#SPILL!"), ErrorValue("#CALC!")
ITERATION_LIMIT = "iterative calculation past ReadLimits.max_iteration_evaluations"
STEP_LIMIT = "function steps past ReadLimits.max_function_steps"
ITERATION_UNREAD = "iterative calculation reading a cell its pass did not evaluate"


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
    forms: frozenset[str] = frozenset()
    """Forms its text uses that are not computed where evaluation reaches them (``Compiled``)."""


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


_PAD_ROWS = 4096
"""A range taller than this whose data fills less than half of it is padded, not made dense."""
_LARGE_BOX = 256  # cells: smaller boxes are listed again rather than remembered
_COMPUTED = CellStatus("computed")  # frozen: one instance serves every such cell
_CONSTANT = CellStatus("constant")


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
        iterate: bool = False,
        iterate_count: int = 100,
        iterate_delta: float = 0.001,
        max_iteration_evaluations: int | None = 100_000_000,
        max_function_steps: int | None = 1_000_000_000,
    ) -> None:
        self.quirks = quirks or Quirks.EXCEL
        self.iterate = iterate
        """The workbook has iterative calculation on (``<calcPr iterate>``): §7.2."""
        self.iterate_count = max(1, iterate_count)
        """Iterations a run makes at most (``iterateCount``)."""
        self.iterate_delta = iterate_delta
        """A run stops once no iterating cell changes by this much (``iterateDelta``)."""
        self.max_iteration_evaluations = max_iteration_evaluations
        """Cell evaluations one run may make (``ReadLimits.max_iteration_evaluations``)."""
        self._iterating = False  # within a run: kept cells are computed
        self._iteration_due = False
        # finishing a pass before its run: the passes it starts leave the run to it
        self._finishing = False
        self._iterated: set[int] = set()  # kept cells a run computed: the writer writes them
        self._last_stop: str | None = None  # why the last run ended: "delta", "count" or a label
        self._before_pass: dict[int, Value] = {}  # with iteration on: values before this pass
        self._partial = False  # a pass with formula cells not stale: a loop may be met late
        self.calc_always: set[int] = set()
        """With iteration on, formulas the file marks calculate-always (``<f ca="1">``), as Excel
        marks a loop's: an edit makes them stale, so it runs a loop a trusting load never met."""
        self.max_array_cells = max_array_cells
        """The largest array a formula may make (``ReadLimits.max_array_cells``)."""
        self.max_function_steps = max_function_steps
        """Steps functions may take in one recalculation (``ReadLimits.max_function_steps``)."""
        self._steps_left = max_function_steps
        self._report_due = False  # reads deferred the report's passes (``settled_report``)
        self._spill_due = False  # a dynamic array may be stale: evaluated before a read
        self.reference_array: Any = None
        """An array of references made while the current cell is evaluated."""
        self.date1904 = date1904
        self.today = today or dt.date.today()
        self.sheets: list[Sheet] = []
        self._sheet_index: dict[str, int] = {}
        self.templates: list[Template] = []
        self._template_ids: dict[str, int] = {}
        self._raw_ids: dict[str, int] = {}
        # Volatile cells whose cached value a trusting load kept, not computed since: a change
        # does not wake them, nor reach them through a volatile name's references, as Excel does
        # with an unmarked one (capture volatile_edit), until computed.
        self._dormant: set[int] = set()
        self._unloaded: set[int] = set()  # sheets a lazy Loader has not loaded (need)
        self.load_sheet: Callable[[int], None] | None = None
        self._uncompiled: set[int] = set()  # sheets loaded whose formulas wait (need_formulas)
        self.compile_sheet: Callable[[int], None] | None = None
        self._wanted: list[int] = []
        self._loading = False
        self._loads = 0
        self._spellings: dict[int, str] = {}  # template id -> the raw-text key it shows
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
        self._volatility: Volatility | None = None  # from the names, rebuilt when one changes
        self._changed = False  # a cell changed since volatile cells were last made stale
        # Cells on a cycle or depending on one (§7.2), each with whether it holds a cached value,
        # and the runs of rows they make, so that many cycles walk each dependent once.
        self._kept: dict[int, bool] = {}
        self._kept_runs: Runs = defaultdict(list)
        self.entered: set[int] = set()
        """Formula cells set through the API, not read from a file (§7.2)."""
        # Formula cells the engine cannot evaluate, which keep the value the file holds (§7.6):
        # those an edit reached, whose kept value is out of date; what depends on those, which is
        # not evaluated; the ones an evaluation reached; every cell evaluated, and the boxes it
        # read through OFFSET or INDIRECT, which no static edge shows. All empty, and nothing
        # recorded, in a workbook with no such formula.
        self._absent: set[int] = set()
        self._outdated: set[int] = set()
        self._held_back: set[int] = set()
        self._held_refs: set[CellRef] = set()
        self._reached: set[int] = set()
        # The labels of cells whose own formula was stopped or kept (limits, features not
        # implemented), for :meth:`status`.
        self._stopped: dict[int, set[str]] = {}
        # The (cell, label) pairs ``Report.unsupported`` has counted, each once; and the cells a
        # data table or an array formula with no text keeps with the file's values (§7.6).
        self._labelled: set[tuple[int, str]] = set()
        self._kept_constants: set[int] = set()
        self._evaluated: set[int] = set()
        self._dynamic_reads: dict[int, list[tuple[int, int, int, int, int]]] = {}
        self.active = (0, 0, 0)
        """The active sheet and its active cell, 0-based: what ``CELL`` without a reference
        reads (capture ``cell_noref``)."""
        self.path: str | None = None
        """The workbook's file, for ``CELL("filename")``: an absolute path, or ``None``."""
        self.ops = Ops(self)
        # The cell being evaluated (the evaluation context of §7.3).
        self.sheet = self.row = self.col = 0
        self._index: DependencyIndex | None = None
        self._live: dict[tuple[int, int, int], tuple[RefSpec, ...]] = {}
        """A template's references that can reach a formula cell, per (template, sheet, column):
        kept, as the index is, until a formula or a name changes (:meth:`_formulas_changed`)."""
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
        index = self._sheet_index.get(name.casefold())
        if self._unloaded and index in self._unloaded:
            self.need(index)
        return index

    def need(self, sheet: int) -> None:
        """Load a sheet not loaded yet (a lazy ``Loader``'s, ``_load.py``) before it is read.

        Every reference that names a sheet resolves it through :meth:`sheet_index`, which calls
        this; a load asked for during another (a compile resolving a sheet) waits for it to end.
        """
        if sheet not in self._unloaded or self.load_sheet is None:
            return
        self._wanted.append(sheet)
        if self._loading:
            return
        self._loading = True
        try:
            while self._wanted:
                wanted = self._wanted.pop()
                if wanted in self._unloaded:
                    self._unloaded.discard(wanted)
                    self._loads += 1
                    self.load_sheet(wanted)
        finally:
            self._loading = False
        self._keep_late()

    def _keep_late(self) -> None:
        """Formulas a lazy load adds that depend on a cycle kept before keep their values too.

        A cycle keeps every cell depending on it (:meth:`_keep_cycle`) through the dependency
        index of the sheets read when it closed; a sheet read later holds dependents that index
        did not have. With iteration on, every sheet is read before anything is evaluated.
        """
        if self._kept and not self.iterate:
            self._keep_cycle(set(self._kept))

    def need_formulas(self, sheet: int) -> None:
        """Load a sheet, and compile its formulas if its load left them.

        A lazy ``Loader`` leaves them where every cached value stands (``opening`` is
        ``"none"``).
        """
        self.need(sheet)
        if sheet in self._uncompiled and self.compile_sheet is not None:
            self._uncompiled.discard(sheet)
            self._loads += 1
            self.compile_sheet(sheet)
            self._keep_late()

    def read_all(self) -> None:
        """Load every sheet not loaded yet, in the sheets' order."""
        while self._unloaded:
            self.need(min(self._unloaded))

    def load_all(self) -> None:
        """Load and compile every sheet not yet loaded, in order: for whole-workbook work."""
        self.read_all()
        while self._uncompiled:
            self.need_formulas(min(self._uncompiled))

    def define_name(self, name: str, text: str, scope: int | None = None) -> None:
        self.names[(scope, name.casefold())] = text
        self._volatility = None
        self._name_templates.clear()
        self._formulas_changed()

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
            if self._spellings.setdefault(tid, raw) != raw:
                # The same formula spelled otherwise (its spaces): computed alike, and shown as
                # written (capture formula_bar) whatever cell was compiled first.
                spelled = replace(self.templates[tid], origin=(text, row, col))
                tid = len(self.templates)
                self.templates.append(spelled)
                self._spellings[tid] = raw
            self._raw_ids[raw] = tid
        return tid

    @property
    def volatility(self) -> Volatility:
        """Which formulas are volatile, given the names defined so far (``_opening.Volatility``).

        A template's flag is taken when it is compiled: names are defined before formulas.
        """
        if self._volatility is None:
            self._volatility = Volatility((name, text) for (_, name), text in self.names.items())
        return self._volatility

    def _parsed_template(self, text: str, row: int, col: int, array_mode: bool = False) -> int:
        formula = parse(text)
        key = _template_key(formula, row, col)
        if array_mode:
            key = "{array}" + key
        tid = self._template_ids.get(key)
        if tid is None:
            tid = self._compile(formula, key, row, col, array_mode)
        return tid

    def _compile(
        self,
        formula: Any,
        key: str,
        row: int,
        col: int,
        array_mode: bool = False,
        as_reference: bool = False,
    ) -> int:
        compiled = compile_formula(
            formula, row, col, REGISTRY, self.ops, self.quirks, array_mode, as_reference
        )
        unsupported = {
            bare_label(c) for c in compiled.calls if c not in REGISTRY and not self._defined(c)
        }
        unsupported |= compiled.features
        template = Template(
            key,
            compiled.fn,
            compiled.refs,
            compiled.calls,
            self.volatility(formula.text()),
            frozenset(unsupported),
            "SUBTOTAL" in compiled.calls,
            (formula.text(), row, col),
            compiled.intersects,
            array_mode,
            compiled.forms,
        )
        tid = len(self.templates)
        self.templates.append(template)
        self._template_ids[key] = tid
        return tid

    def _defined(self, call: str) -> bool:
        """A call of a defined name (a ``LAMBDA`` it holds, when evaluated), in any scope."""
        folded = call.casefold()
        return "." not in call and any(name == folded for _, name in self.names)

    def formula_text(self, sheet: int, row: int, col: int) -> str | None:
        """The cell's formula in file form, or ``None`` for a cell without one."""
        if self._uncompiled:
            self.need_formulas(sheet)
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
            self._formulas_changed()
        if self._absent:
            self._absent.discard(n)
            self._outdated.discard(n)
        self._touch_spill(sheet, row, col)
        s.set(row, col, value)
        self._volatile.discard(n)
        self._dormant.discard(n)
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
        stale: bool = False,
        array: tuple[bool, tuple[int, int, int, int]] | None = None,
    ) -> None:
        """A formula (file form) or an existing template; ``cached`` is its trusted value.

        With ``stale`` the cell is evaluated all the same, and ``cached`` is only what it holds
        until then: it stays if the cell turns out to be on a cycle (§7.2).

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
        self._formulas_changed()
        self._changed = True
        compiled = self.templates[tid]
        self._dormant.discard(n)
        if compiled.volatile:
            self._volatile.add(n)
        else:
            self._volatile.discard(n)
        if compiled.unsupported:
            self._absent.add(n)
            self._outdated.discard(n)
        elif self._absent:
            self._absent.discard(n)
            self._outdated.discard(n)
        if array is not None:
            self.anchors[n] = Anchor(array[0], array[1])
            self._spill_due = True
            self._index_anchors(sheet)
        if cached is None or stale:
            s.set(row, col, BLANK if cached is None else cached)
            self.stale.add(n)
            self._unclosed.append(n)
        else:
            s.set(row, col, cached)
            if self.templates[tid].volatile:
                self.stale.add(n)
                self._unclosed.append(n)
            elif array is not None and array[0] and cached == SPILL:
                # A blocked dynamic array is evaluated (to the same #SPILL!) to learn the range
                # it wants, so that clearing a cell there lets it spill; its dependents stay
                # trusted.
                self.stale.add(n)

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

    def arr_padded(self, v: Any) -> Any:
        """As :meth:`arr`, but a tall range whose data stops early is a :class:`Padded` (§7.7).

        The compiler calls this only where every reader of the array is an elementwise operator
        or a function taught to read a padded array.
        """
        if not isinstance(v, Ref) or v.is_cell() or v.rows <= _PAD_ROWS:
            return self.arr(v)
        cells = {(r, c): x for r, c, x in v.cells()}
        height = max((r for r, _ in cells), default=-1) + 1
        if 2 * height > v.rows:
            return to_array(v)
        check_area(height, v.columns, self.max_array_cells)
        head = tuple(
            tuple(cells.get((r, c), BLANK) for c in range(v.columns)) for r in range(height)
        )
        return Padded(head, v.rows, v.columns, BLANK)

    def spill_range(self, ref: Any) -> Any:
        """``A1#`` (``_xlfn.ANCHORARRAY(A1)``): the range the dynamic array at A1 spills into."""
        if not isinstance(ref, Ref) or not ref.is_cell():
            return REF
        anchor = self.anchors.get(node(ref.sheet, ref.r1, ref.c1))
        if anchor is None or not anchor.dynamic:
            return REF
        if anchor.wanted is not None:  # blocked: the anchor alone, holding #SPILL! (dyn_sequence)
            return Ref(self.cells, ref.sheet, ref.r1, ref.c1, ref.r1, ref.c1)
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
            self._evaluate_stale()

    # Values -----------------------------------------------------------------------------------

    def get(self, sheet: int, row: int, col: int) -> Value:
        """The cell's current value, evaluating it first if it is stale."""
        n = node(sheet, row, col)
        if self._unclosed or self._changed:
            self._close()
        if self._spill_due:
            self._settle_spills()
        anchor = self.member_of(n)
        if anchor is not None and anchor in self.stale:
            self.recalculate(targets=[anchor], defer_report=True)
        if n in self.stale:
            self.recalculate(targets=[n], defer_report=True)
        return self.sheets[sheet].values.get(row * COLS + col, BLANK)

    def _settle_spills(self) -> None:
        """Evaluate the stale dynamic arrays, as a read does before anything else.

        A dynamic array an edit makes stale may spill into cells outside the range it fills now:
        a blank cell there, or a formula reading one, has no edge to it until it is evaluated
        and its range grows (§7.7). Excel recalculates every dynamic array an edit reaches before
        anything is read, so a read does too.
        """
        self._spill_due = False
        stale = [a for a, x in self.anchors.items() if x.dynamic and a in self.stale]
        if stale:
            self.recalculate(targets=stale, defer_report=True)

    def get_many(self, cells: Sequence[tuple[int, int, int]]) -> list[Value]:
        """The cells' values, as reading each in turn with :meth:`get` gives them.

        One evaluation takes every read's stale cells, in the order and with the allowance of
        function steps each read would give them (an array's anchor first, then the cell), and
        each value is taken when its read would return it. Where a spill moved, the evaluation
        stops, settles it as the read would before returning, and goes on. Where something is
        left to close before a read (the cells of a sheet a lazy load read during the pass), the
        evaluation stops and closes it, as that read would before evaluating: closed only at the
        end, it would make cells already read stale, and each status read after would evaluate
        them and run the report's passes again. With iteration on, each cell is read
        with :meth:`get`.
        """
        if self.iterate:
            return [self.get(*cell) for cell in cells]
        values: list[Value] = [BLANK] * len(cells)
        at, phase, due = 0, 0, False  # phase: 0 not begun, 1 the anchor's turn done, 2 the cell's

        def feed() -> Iterator[int]:
            nonlocal at, phase, due
            fed = False
            while at < len(cells):
                if fed and self._moved:
                    return  # settle before the next read, as a read does before returning
                if phase == 0 and (self._unclosed or self._changed or self._spill_due):
                    return  # close before the next read, as a read does before evaluating
                sheet, row, col = cells[at]
                n = node(sheet, row, col)
                if phase == 0:
                    phase = 1
                    anchor = self.member_of(n)
                    if anchor is not None and anchor in self.stale:
                        self._steps_left, due, fed = self.max_function_steps, True, True
                        yield anchor
                        continue
                if phase == 1:
                    phase = 2
                    if n in self.stale:
                        self._steps_left, due, fed = self.max_function_steps, True, True
                        yield n
                        continue
                values[at] = self.sheets[sheet].values.get(row * COLS + col, BLANK)
                at, phase = at + 1, 0

        while at < len(cells):
            # feed() stops at a read with something to close or a dynamic array due, and reads
            # nothing until both are done: this close and this settling are what let the loop
            # advance (without either the loop never ends).
            if self._unclosed or self._changed:
                self._close()
            if self._spill_due:
                self._settle_spills()
            self._evaluate(feed())
            self._settle()
        if due:
            self._report_due = True
            if not self.stale:
                self.stale.clear()  # a set keeps its table when emptied; clearing frees it
        return values

    def loop_live(self) -> bool:
        """With iteration on, a loop has run: every later recalculation runs it again (§7.2)."""
        return self.iterate and bool(self._kept)

    def save_run(self) -> None:
        """Saving runs a live loop once more, as Excel's saving does (§7.2)."""
        if self.loop_live():
            self.recalculate()

    @property
    def dormant(self) -> set[int]:
        """Volatile cells a trusting load kept, not computed since (:meth:`trust_volatile`).

        The dependency index leaves out a volatile name's references for them (§6.3).
        """
        return self._dormant

    def has_volatile(self) -> bool:
        """Whether a formula cell's template is volatile (``RAND``, ``NOW``, ``OFFSET``...)."""
        return bool(self._volatile)

    def trust_volatile(self) -> None:
        """Volatile formulas holding a cached value keep it.

        As Excel keeps them on opening a workbook whose cells it does not recalculate in full:
        in manual calculation (capture ``open_calc``, ``manual_volatile``), and otherwise unless
        their formula is marked ``ca="1"`` (capture ``volatile_names``), which the load then
        makes stale. ``load`` calls it under ``recalc_on_load="excel"`` when it does not evaluate
        every formula.
        """
        trusted = {
            n for n in self._volatile if (n & KEY_MASK) in self.sheets[n >> SHEET_SHIFT].values
        }
        self.stale -= trusted
        self._unclosed = [n for n in self._unclosed if n not in trusted]
        self._dormant |= trusted

    def loaded(self) -> None:
        """The load is done: its own formulas are no change.

        So the first edit, like every later one, makes the volatile formulas stale
        (:meth:`_close`).
        """
        self._changed = False

    def make_all_stale(self) -> None:
        """Every formula stale: evaluated when next read or recalculated."""
        for s, sheet in enumerate(self.sheets):
            self.stale.update((s << SHEET_SHIFT) | k for k in sheet.templates)
        self._unclosed.clear()
        self._spill_due = True

    def recalculate(
        self,
        full: bool = False,
        targets: Iterable[int] | None = None,
        *,
        defer_report: bool = False,
    ) -> Report:
        """Evaluate stale cells (every formula with ``full``) in dependency order.

        ``defer_report``: leave the report's whole-workbook passes (what is uncomputed and what
        rests on a kept value) for :meth:`settled_report`, so that reading cells one by one, each
        read evaluating only what that cell needs, does not repeat them for every cell.
        """
        self._steps_left = self.max_function_steps
        if full:
            self.make_all_stale()
        if targets is None:
            self._changed = True  # a recalculation evaluates volatile formulas again
            self._close()
            self._evaluate_stale()
        else:
            self._evaluate(list(targets))
        self._settle()
        if targets is None:
            self._before_pass.clear()  # nothing is left stale: no later pass can need them
        if defer_report:
            self._report_due = True
        else:
            self._finish_report()
        if not self.stale:
            self.stale.clear()  # a set keeps its table when emptied; clearing frees it
        return self.report

    def _finish_report(self) -> None:
        """The report's whole-workbook passes, each a function of the engine's state."""
        self._report_passes()
        self._order_report()

    def _report_passes(self) -> None:
        """What is uncomputed and what rests on a kept value, over the sheets read so far."""
        self._report_held_back()
        self._taint()
        self._report_from_kept()
        self._report_due = False

    def settled_report(self) -> Report:
        """The report, with the passes deferred by the reads since the last one run."""
        if self._report_due:
            self._finish_report()
        else:
            self._order_report()
        return self.report

    def _count_labels(self, n: int, labels: Iterable[str]) -> None:
        """Count a formula cell's own unsupported functions in the report, each once a cell."""
        for label in labels:
            if (n, label) not in self._labelled:
                self._labelled.add((n, label))
                self.report.unsupported[label] += 1

    def keep_constants(self, sheet: int, box: tuple[int, int, int, int], label: str) -> None:
        """Cells a data table or an array formula with no text keeps with the file's values."""
        r1, c1, r2, c2 = box
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                n = node(sheet, r, c)
                self._kept_constants.add(n)
                self.stop(n, label)
        self._report_due = True  # they are in ``from_kept`` from here on

    def stop(self, n: int, label: str) -> None:
        """Note a cell whose own formula was stopped or kept, under ``label`` (:meth:`status`)."""
        self._stopped.setdefault(n, set()).add(label)

    def status(self, sheet: int, row: int, col: int) -> CellStatus:
        """What the report says of one cell, from the sheets read so far.

        The cell's sheet is read and compiled, and the cell evaluated as :meth:`get` evaluates
        it, which reads every sheet the references of the formulas it evaluates name, through
        defined names and in branches not taken too: the sheets a static precedent can lie on,
        which the report's passes follow. (A "some" opening and iterative calculation read every
        sheet first; under "none" nothing is evaluated, and an edit reads every sheet.) The
        passes then run, if a read since the last run left them due, so that reading a range's
        values first and then each cell's status runs them once. A cell an array formula fills
        answers with that formula's status.
        """
        self.need_formulas(sheet)
        n = node(sheet, row, col)
        anchor = self.member_of(n)
        target = n if anchor is None else anchor
        self.get(sheet, row, col)
        if self._report_due:
            self._report_passes()
        return self._classify(target, anchor is not None)

    def _classify(self, n: int, filled: bool) -> CellStatus:
        """A cell's status from the report as it stands, in :class:`CellStatus`'s order."""
        sheet, row, col = unpack(n)
        s = self.sheets[sheet]
        tid = s.templates.get(row * COLS + col)
        report = self.report
        if not filled and not (
            report.refused or report.uncomputed or report.circular or report.from_kept
        ):  # a report naming no cell: the usual answer, without building the cell's reference
            return _CONSTANT if tid is None else _COMPUTED
        where = CellRef(s.name, row + 1, col + 1)
        anchor = where if filled else None
        refusal = report.refused.get(where)
        if refusal is not None:
            return CellStatus("refused", refusal=refusal, anchor=anchor)
        if where in report.uncomputed:
            own = self._labels(n, tid)
            error = report.exceptions.get(where)
            inherited = not own and error is None
            return CellStatus("uncomputed", own, error, inherited=inherited, anchor=anchor)
        if tid is None:
            if where in report.from_kept:  # a data table's, or an array formula's with no text
                return CellStatus("kept", self._labels(n, tid), anchor=anchor)
            return _CONSTANT
        if where in report.circular:
            return CellStatus("circular", anchor=anchor)
        if where in report.from_kept:
            kept = n in self._absent
            own = self._labels(n, tid) if kept else ()
            return CellStatus("kept", own, inherited=not kept, anchor=anchor)
        return CellStatus("computed", anchor=anchor) if filled else _COMPUTED

    def _labels(self, n: int, tid: int | None) -> tuple[str, ...]:
        """A cell's own labels: the limits its evaluation passed and its formula's features."""
        labels = set(self._stopped.get(n, ()))
        if tid is not None:
            labels.update(self.templates[tid].unsupported)
        return tuple(sorted(labels))

    def _order_report(self) -> None:
        """The report's cells in sheet, row and column order, and its labels in text order.

        Cells enter it in the order they are compiled and evaluated, which depends on the order
        sheets were read and values asked; ordered, the report is the same however it was made.
        """
        report = self.report
        index = {s.name: i for i, s in enumerate(self.sheets)}
        last = len(self.sheets)

        def position(item: tuple[CellRef, str]) -> tuple[int, str, int, int]:
            ref = item[0]
            return (index.get(ref.sheet, last), ref.sheet, ref.row, ref.col)

        for cells in (report.refused, report.exceptions):
            if len(cells) > 1:
                ordered = sorted(cells.items(), key=position)
                cells.clear()
                cells.update(ordered)
        if len(report.unsupported) > 1:
            labels = sorted(report.unsupported.items())
            report.unsupported.clear()
            report.unsupported.update(dict(labels))

    def _evaluate_stale(self) -> None:
        """Evaluate every stale cell, in an order that depends only on which cells are stale.

        Which cells pass a step limit, and the order of ``RAND``'s draws, follow the order cells
        are evaluated in; a set's own order follows the order its cells were added. A large set
        goes in store order (sheet, column, row: the store's own sorted columns, read as
        evaluated, cheaper than sorting past an eighth of the formulas), the rest sorted (sheet,
        row, column).
        """
        if 8 * len(self.stale) >= self.formula_count():
            self._evaluate(self._stale_in_store_order())
        if self.stale:
            self._evaluate(sorted(self.stale))

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
            woken = self._volatile - self._dormant if self._dormant else self._volatile
            self.stale |= woken
            self._unclosed.extend(woken)
        if self._changed:
            self._before_pass.clear()  # a recalculation starts from what the cells hold now
        if self._changed and self._kept:
            self._reconsider_kept()
        if self._changed and (self._absent or self._held_back):
            self._held_back = self._out_of_date()
        self._changed = False
        if self._unclosed and len(self.stale) < self.formula_count():
            reached = self.dependents(self._unclosed)
            if self._kept:  # kept by a cycle closed since they were compiled: final until a change
                reached -= self._kept.keys()
            self.stale |= reached
        self._unclosed.clear()
        self._spill_due = True

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
        # Only where some formula is not stale can a loop's cell be computed before its cycle
        # shows itself (it reads a loop cell that is not stale): keep the values a run starts
        # from then, and not in a full pass, where they would cost an entry per formula.
        self._partial = self.iterate and len(pending) < self.formula_count()
        on_stack: dict[int, int] = {}  # node -> its position on the current path
        path: list[int] = []
        suspects: set[int] = set()  # on a static cycle: evaluated with every read checked
        marked_from = marked_to = 0  # path[marked_from:marked_to] is all in suspects
        clean: set[tuple[int, int, int, int, int]] = set()  # large boxes with nothing stale left
        live = self._live
        waits: dict[int, int] = {}  # a cell stopped by a read, and the cell it read
        for target in targets:  # one at a time: no stack entry per stale cell up front
            if target not in pending:
                continue
            stack: list[tuple[int, bool]] = [(target, False)]
            while stack:
                n, expanded = stack.pop()
                if n not in pending:
                    if expanded:  # kept while it was on the path, by a cycle closing
                        on_stack.pop(n, None)
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
                        self._keep_cycle(chain)
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

        if self._iteration_due and not self._iterating and not self._finishing:
            if self.stale:  # one run after every cell Excel computes with it (iterate_calc)
                self._finishing = True
                try:
                    self._evaluate_stale()
                finally:
                    self._finishing = False
            self._iterate()

    def spend_steps(self, steps: int) -> None:
        """Charge a function's loop to the recalculation's allowance before it runs.

        Raises:
            StepLimitError: The steps would pass ``ReadLimits.max_function_steps``.
        """
        if self._steps_left is None:
            return
        if steps > self._steps_left:
            raise StepLimitError(self.max_function_steps or 0)
        self._steps_left -= steps

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
        if self._absent or self._kept_constants:
            where = node(self.sheet, self.row, self.col)
            self._dynamic_reads.setdefault(where, []).append((sheet, r1, c1, r2, c2))
        return Ref(self.checked, sheet, r1, c1, r2, c2)

    def _keep_cycle(self, chain: set[int]) -> None:
        """A cycle closed: its cells and every cell depending on them keep what they hold (§7.2).

        They leave the stale set at once, so evaluation replaces nothing and never reaches a
        cell depending on them. With iteration off that is what Excel does (capture
        ``circular_cached``), and the cells are ``circular``; with it on, they are iterated once the
        pass ends (``_iterate``).
        """
        index = self.index()
        cells: Iterable[int] = chain
        seeds: Iterable[int] = chain
        ranges: list[tuple[int, int, int, int]] = []
        while True:
            for n in cells:
                anchor = self.anchors.get(n)
                if anchor is not None:  # its range changes with it: what reads that follows
                    sheet = n >> SHEET_SHIFT
                    for box in (anchor.box, anchor.wanted):
                        if box is not None:
                            ranges += [
                                (sheet, c, box[0], box[2]) for c in range(box[1], box[3] + 1)
                            ]
                self._keep(n)
            found = index.spread(seeds, ranges, self._kept_runs)
            if not found:
                return
            seeds, ranges = (), []
            cells = [
                (sheet << SHEET_SHIFT) | (row * COLS + col)
                for sheet, col, first, last in found
                for row in range(first, last + 1)
            ]

    def _keep(self, n: int) -> None:
        sheet, row, col = unpack(n)
        s = self.sheets[sheet]
        key = row * COLS + col
        if key not in s.templates or n in self._kept:
            return
        held = s.values.get(key) is not None
        self._kept[n] = held
        where = CellRef(s.name, row + 1, col + 1)
        self.stale.discard(n)  # final: a later read must not compute it again
        if self.iterate:  # iterated once the pass ends; with no value, the empty text (§7.2)
            if not held:
                s.values[key] = ""
            self._iteration_due = True
        else:
            if not held:  # from a file, the empty text; entered, 0 (capture circular_cached)
                s.values[key] = 0.0 if n in self.entered else ""
            self.report.circular.add(where)

    def _iteration_order(self, cells: Iterable[int]) -> list[int]:
        """Excel's order: sheets by name in its text order, then row, then column (§7.2)."""
        names = [s.name for s in self.sheets]

        def before(a: int, b: int) -> int:
            if compare(names[a], names[b], "<") is True:
                return -1
            return 1 if compare(names[a], names[b], ">") is True else 0

        by_name = sorted(range(len(names)), key=cmp_to_key(before))
        rank = {sheet: i for i, sheet in enumerate(by_name)}
        return sorted(cells, key=lambda n: (rank[n >> SHEET_SHIFT], n & KEY_MASK))

    def _iterate(self) -> None:
        """One run of iterative calculation over the cells kept on and after cycles (§7.2).

        Each iteration computes every one in Excel's order, in place, reading what the others
        hold; the run stops after ``iterate_count`` iterations, or after the first in which no
        number changed by ``iterate_delta`` or more (an old value that is not a number counts as 0,
        and a new one that is not a number is no change).
        Past ``max_iteration_evaluations`` the cells keep the values they held before the run.
        """
        self._iteration_due = False
        members = [
            n
            for n in self._iteration_order(self._kept)
            if self.sheets[n >> SHEET_SHIFT].templates.at(*divmod(n & KEY_MASK, COLS)) is not None
        ]
        stores = [(self.sheets[n >> SHEET_SHIFT].values, (n & KEY_MASK)) for n in members]
        for n, (values, key) in zip(members, stores, strict=True):
            if n in self._before_pass:  # computed before the cycle showed itself
                old = self._before_pass.pop(n)
                values[key] = "" if old is BLANK else old  # no value: the empty text, as Excel
        before = [values.get(key, BLANK) for values, key in stores]
        limit = self.max_iteration_evaluations
        delta = self.iterate_delta
        stopped: str | None = None  # why the run was undone: the report's label
        evaluations = 0
        self._iterating = True
        try:
            for _ in range(self.iterate_count):
                evaluations += len(members)
                if limit is not None and evaluations > limit:
                    stopped = ITERATION_LIMIT
                    break
                largest = 0.0
                for n, (values, key) in zip(members, stores, strict=True):
                    old = values.get(key, BLANK)
                    self._compute(n)
                    new = values.get(key, BLANK)
                    if type(new) is float:  # an old value not a number counts as 0
                        change = abs(new - old) if type(old) is float else abs(new)
                        if change > largest:
                            largest = change
                if largest < delta:
                    self._last_stop = "delta"
                    break
            else:
                self._last_stop = "count"
        except _NotEvaluatedError:  # a read the pass did not see: keep what was held
            stopped = ITERATION_UNREAD
        finally:
            self._iterating = False
        if stopped:
            self._last_stop = stopped
        for n, (values, key), value in zip(members, stores, before, strict=True):
            if stopped:
                if value is BLANK:
                    values.pop(key, None)
                else:
                    values[key] = value
                sheet, row, col = unpack(n)
                self.report.unsupported[stopped] += 1
                self.report.uncomputed.add(CellRef(self.sheets[sheet].name, row + 1, col + 1))
                self.stop(n, stopped)
            else:
                self._iterated.add(n)
            self._kept[n] = True  # what it holds now is its value until a change

    def _reconsider_kept(self) -> None:
        """After a change, kept cells are evaluated again: an edit may have broken the loop."""
        report = self.report
        for n, held in self._kept.items():
            sheet, row, col = unpack(n)
            s = self.sheets[sheet]
            key = row * COLS + col
            where = CellRef(s.name, row + 1, col + 1)
            report.uncomputed.discard(where)
            report.circular.discard(where)
            if key in s.templates:
                self.stale.add(n)
                if not held:
                    s.values.pop(key, None)
        del report.unsupported[ITERATION_LIMIT]
        del report.unsupported[ITERATION_UNREAD]
        self._iterated.clear()
        self._kept.clear()
        self._kept_runs.clear()

    def kept_with_values(self) -> Iterator[tuple[str, int, int]]:
        """Cells on a cycle or after one holding the file's value, as (sheet, row, col), 1-based."""
        for n, held in self._kept.items():
            if held and n not in self._iterated:
                sheet, row, col = unpack(n)
                yield self.sheets[sheet].name, row + 1, col + 1

    def _compute(self, n: int) -> None:
        if self._dormant:
            self._dormant.discard(n)  # computed: from now on recomputed at every change
        sheet, row, col = unpack(n)
        s = self.sheets[sheet]
        key = row * COLS + col
        tid = s.templates.at(row, col)
        if tid is None:
            return
        if self._kept and n in self._kept and not self._iterating:  # until a change
            return
        template = self.templates[tid]
        where = CellRef(s.name, row + 1, col + 1)
        if template.unsupported:
            self._count_labels(n, template.unsupported)
        if self._held_back and n in self._held_back:  # rests on a value out of date (§7.6)
            if key not in s.values:
                s.values[key] = NAME
            return
        if template.unsupported:
            if key not in s.values:  # no cached value: a formula cell is never blank otherwise
                s.values[key] = NAME
                self.report.uncomputed.add(where)
            else:
                self._reached.add(n)
            return
        if self._absent or self._kept_constants:
            self._evaluated.add(n)
            self._dynamic_reads.pop(n, None)
        if self._partial and not self._iterating and n not in self._before_pass:
            self._before_pass[n] = s.values.get(key, BLANK)  # a run starts from it (§7.2)
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
            label = f"an array of more than {exc.limit:,} cells"
            self.report.unsupported[label] += 1
            self.report.uncomputed.add(where)
            self.stop(n, label)
            return
        except StepLimitError:  # likewise
            self.report.unsupported[STEP_LIMIT] += 1
            self.report.uncomputed.add(where)
            self.stop(n, STEP_LIMIT)
            return
        except NotImplementedFormError as exc:  # keep the cached value, likewise
            self.report.unsupported[exc.label] += 1
            self.report.uncomputed.add(where)
            self.stop(n, exc.label)
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
        if type(value) is Lambda:  # a function value is no cell's value (capture lambda_calls)
            return CALC
        return value  # type: ignore[no-any-return]

    def edited(self, n: int) -> None:
        """A change through the API: what it reaches among the formulas keeping a value (§7.6).

        Their kept values are out of date from here on. A formula using ``OFFSET`` or ``INDIRECT``
        is never evaluated, so what it reads is not known: any change may reach it.
        """
        self._report_due = True  # what the report's passes follow has changed
        if self._absent:
            self._outdated |= self._absent & self.dependents([n])
            self._outdated |= self._absent & self._volatile
        if self.calc_always and not self._kept:  # any edit runs Excel's loops (iterate_calc)
            self.stale |= self.calc_always
            self._unclosed.extend(self.calc_always)

    def _closure(self, seeds: set[int]) -> set[int]:
        """Formula cells depending on the seeds, through static edges and the reads recorded."""
        found = self.dependents(seeds)
        reads = self._dynamic_reads
        while reads:
            more = [
                reader
                for reader, boxes in reads.items()
                if reader not in found and self._reads_any(boxes, seeds, found)
            ]
            if not more:
                break
            found.update(more)
            found |= self.dependents(more)
        return found

    def _reads_any(
        self, boxes: list[tuple[int, int, int, int, int]], seeds: set[int], found: set[int]
    ) -> bool:
        for sheet, r1, c1, r2, c2 in boxes:
            for r, c in self.sheets[sheet].formulas_in(r1, c1, r2, c2):
                cell = node(sheet, r, c)
                if cell in seeds or cell in found:
                    return True
        return False

    def _out_of_date(self) -> set[int]:
        return self._outdated | self._closure(self._outdated) if self._outdated else set()

    def _report_held_back(self) -> None:
        """Cells resting on a kept value an edit put out of date are uncomputed, and only they."""
        report = self.report
        report.uncomputed -= self._held_refs
        refs: set[CellRef] = set()
        if self._outdated:
            for n in self._out_of_date():
                sheet, row, col = unpack(n)
                s = self.sheets[sheet]
                tid = s.templates.get(row * COLS + col)
                if tid is not None:
                    refs.add(CellRef(s.name, row + 1, col + 1))
                    if self.templates[tid].unsupported:
                        self._count_labels(n, self.templates[tid].unsupported)
        self._held_refs = refs
        report.uncomputed |= refs

    def _report_from_kept(self) -> None:
        """Name what rests on a kept value: it, and every cell evaluated from it (§7.6)."""
        names = self.report.from_kept
        names.clear()
        if not self._absent and not self._kept_constants:
            return
        kept = (self._absent - self._outdated) | self._kept_constants
        rest = self._closure(kept) & self._evaluated if kept else set()
        behind = (kept & self._reached) | self._kept_constants
        if rest and not kept <= behind:
            behind |= self._kept_behind(rest, kept)
        for n in behind | rest:
            sheet, row, col = unpack(n)
            s = self.sheets[sheet]
            tid = s.templates.get(row * COLS + col)
            if tid is not None or n in self._kept_constants:
                names.add(CellRef(s.name, row + 1, col + 1))
            if tid is not None and n in self._absent:  # named for its own function: counted
                self._count_labels(n, self.templates[tid].unsupported)
        names -= self.report.uncomputed
        names -= self.report.circular

    def _kept_behind(self, rest: set[int], kept: set[int]) -> set[int]:
        """The formulas keeping a value that the evaluated cells read, walking back from them."""
        seen: set[int] = set(rest)
        queue = list(rest)
        behind: set[int] = set()
        while queue:
            n = queue.pop()
            before = list(self.precedents(n))
            for sheet, r1, c1, r2, c2 in self._dynamic_reads.get(n, ()):
                before += [
                    node(sheet, r, c) for r, c in self.sheets[sheet].formulas_in(r1, c1, r2, c2)
                ]
            for p in before:
                if p in seen:
                    continue
                seen.add(p)
                if p in kept:
                    behind.add(p)
                queue.append(p)
        return behind

    def kept_values(self) -> Iterator[tuple[str, int, int]]:
        """Formulas keeping the file's value, for the writer to leave: (sheet, row, col)."""
        for n in self._absent:
            sheet, row, col = unpack(n)
            s = self.sheets[sheet]
            if s.values.get(row * COLS + col) is not NAME:
                yield s.name, row + 1, col + 1

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
            if n in self._absent:  # its own function, never evaluated: counted all the same
                tid = self.sheets[sheet].templates.get(row * COLS + col)
                if tid is not None:
                    self._count_labels(n, self.templates[tid].unsupported)

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
            for inner in self.name_boxes(spec.name or "", sheet):
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

        Nothing becomes stale during a pass, so a large box found with nothing stale (``clean``)
        is not listed again by the many formulas that may refer to it. A template's references
        to columns of no formula are dropped once for all its cells (``live``), and stay dropped
        from one pass to the next until a formula or a name changes: reading cells one by one
        runs a pass for each.
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

    def _formulas_changed(self) -> None:
        """A formula, a name or an array's anchor changed: what is derived from them goes."""
        self._index = None
        self._live.clear()

    def index(self) -> DependencyIndex:
        """The reverse-dependency index, rebuilt after formulas change.

        Building it resolves the sheets formulas name, which may load one (:meth:`need`): it is
        built again until no sheet loaded while it was.
        """
        while self._index is None:
            loads = self._loads
            index = DependencyIndex(self)
            if loads == self._loads:
                self._index = index
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
            return self.unsupported(EXTERNAL)  # type: ignore[return-value]
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

    def _name_text(self, name: str, sheet: int | None) -> str | None:
        folded = name.casefold()
        return self.names.get((sheet, folded), self.names.get((None, folded)))

    def _name_template(self, name: str, sheet: int | None) -> Template | None:
        key = (sheet, name.casefold())
        if key not in self._name_templates:
            text = self._name_text(name, sheet)
            template = None
            if text is not None:
                try:
                    formula = parse(text)
                    name_key = "{name}" + _template_key(formula, 0, 0)
                    tid = self._template_ids.get(name_key)
                    if tid is None:
                        tid = self._compile(formula, name_key, 0, 0, as_reference=True)
                    template = self.templates[tid]
                except (FormulaSyntaxError, TokenizeError):
                    template = None
            self._name_templates[key] = template
        return self._name_templates[key]

    def name_boxes(self, name: str, sheet: int) -> Iterator[RefSpec]:
        """The boxes a defined name refers to, and those of the names it refers to or calls.

        Followed iteratively, each name once, so a name that calls itself (a recursive
        ``LAMBDA``) or a chain of names ends.
        """
        todo, seen = [name], set()
        while todo:
            name = todo.pop()
            if name.casefold() in seen:
                continue
            seen.add(name.casefold())
            template = self._name_template(name, sheet)
            for inner in template.refs if template is not None else ():
                if inner.kind == "box":
                    yield inner
                elif inner.sheet is None:
                    todo.append(inner.name or "")

    def name_text(self, name: str) -> str | None:
        """A defined name's definition as the current sheet sees it, or ``None``."""
        return self._name_text(name, self.sheet)

    def name(self, sheet: str | None, name: str) -> Any:
        workbook = sheet == "[0]"  # [0]!Rate: this workbook's own name (capture workbook_name)
        if sheet is not None and "[" in sheet and not workbook:  # [1]!Rate: another workbook's
            return self.unsupported(EXTERNAL)
        scope = self.sheet if sheet is None or workbook else self.sheet_index(sheet)
        if scope is None:
            return REF
        template = self._name_template(name, None if workbook else scope)
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
        self.stop(node(self.sheet, self.row, self.col), label)
        return NAME

    def row_hidden(self, sheet: int, row: int) -> bool:
        return row in self.sheets[sheet].hidden_rows

    def filtering(self, sheet: int) -> bool:
        """A filter is hiding rows on this sheet."""
        return self.sheets[sheet].filter_mode

    def is_subtotal(self, sheet: int, row: int, col: int) -> bool:
        tid = self.sheets[sheet].templates.at(row, col)
        return tid is not None and self.templates[tid].subtotal


def bare_label(name: str) -> str:
    """The report's label for a function the registry does not have.

    A name Excel writes with a prefix in a file (``_xlfn.XLOOKUP``), found there without it, is
    a name Excel does not know: the label says so and gives the prefix. Otherwise the name.
    """
    prefix = PREFIXES.get(name)
    if not prefix:
        return name
    if name in _BARE_REFUSED_NAMES:
        return f"{name} written without Excel's file prefix ({prefix}): Excel refuses the file"
    return f"{name} written without Excel's file prefix ({prefix}): Excel shows #NAME?"


_BARE_REFUSED_NAMES = frozenset({"SORT", "FILTER"})  # capture dyn_*: the file does not open


def _at(axis: Axis | None, base: int) -> int:
    assert axis is not None
    return axis.value if axis.absolute else base + axis.value
