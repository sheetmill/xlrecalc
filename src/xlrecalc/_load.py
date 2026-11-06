"""Build an engine from an xlsx file read by ``_xlsx`` (ARCHITECTURE.md §7.4, §9.1, §9.5).

Cached values are kept and trusted per §7.4: a formula is stale when the file holds no cached
value for it, when the workbook asks for a full calculation on load, or when its template is
volatile; cells depending on a stale cell become stale when the engine recalculates. A formula
Excel refuses in a file is dropped and its cell keeps its cached value, as Excel's repair does
(§9.5), and the report lists it.
"""

from __future__ import annotations

import datetime as dt
import os
from typing import Any, BinaryIO

from xlrecalc._dates import date_serial
from xlrecalc._engine import Engine, node
from xlrecalc._formula import FormulaSyntaxError, Kind, TokenizeError, tokenize
from xlrecalc._formula._fastkey import raw_key
from xlrecalc._numtext import DBL_MIN
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef
from xlrecalc._values import BLANK, ErrorValue, Value
from xlrecalc._xlsx import Book, Cell, ReadLimits
from xlrecalc._xlsx._cells import _column as _cell

__all__ = ["load", "refused_in_file"]

_UNSEEN = object()

# Error literals a file may hold in a formula (capture formula_forms: #SPILL! and #CALC! are not).
_FILE_ERRORS = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}
)


def refused_in_file(text: str) -> str | None:
    """Why Excel refuses this formula in a file (``formula_forms`` capture), or ``None``."""
    try:
        toks = tokenize(text)
    except TokenizeError as exc:
        return str(exc)
    previous = before = None
    for tok in toks:
        if "\t" in tok.space:
            return "a tab"
        if tok.kind is Kind.ERROR and tok.text not in _FILE_ERRORS:
            return f"the error literal {tok.text}"
        if tok.kind is Kind.OPERATOR and tok.text == "#":
            return "the spill operator"
        if (  # OFFSET(5,0,0): capture dynrefs, the file does not open
            tok.kind in (Kind.NUMBER, Kind.STRING, Kind.BOOL)
            and previous is not None
            and previous.kind is Kind.OPEN
            and before is not None
            and before.kind is Kind.FUNCTION
            and before.text.upper() == "OFFSET"
        ):
            return "a constant as OFFSET's reference"
        if (
            tok.kind is Kind.CLOSE
            and previous is not None
            and previous.kind is Kind.OPEN
            and tok.space
        ):
            return "an argument list of spaces"
        before, previous = previous, tok
    return None


def _filled_by(boxes: list[tuple[int, int, int, int, int]], row: int, col: int) -> int | None:
    """The array formula whose range holds the cell, if any."""
    for r1, c1, r2, c2, anchor in boxes:
        if r1 <= row <= r2 and c1 <= col <= c2:
            return anchor
    return None


def _value(
    cell: Cell, date1904: bool, quirks: Quirks, errors: dict[str, str] | None = None
) -> Value | None:
    """A cell's cached value as the engine holds it; ``None`` when there is none.

    An error a file caches as ``#VALUE!`` with a rich value behind it (``vm``) is that error
    (a blocked spill is ``#SPILL!``: capture ``arrays``).
    """
    t, v = cell.type, cell.value
    if t == "none" or v is None:
        return None
    if t == "n":
        assert isinstance(v, float)
        return 0.0 if -DBL_MIN < v < DBL_MIN else v  # Excel reads a subnormal as 0
    if t == "b":
        return bool(v)
    if t == "e":
        if errors and cell.vm is not None and cell.vm in errors:
            return ErrorValue(errors[cell.vm])
        try:
            return ErrorValue(str(v))
        except ValueError:
            return str(v)
    if t == "d":
        return _iso_serial(str(v), date1904, quirks)
    return str(v)


def _iso_serial(text: str, date1904: bool, quirks: Quirks) -> Value:
    try:
        when = dt.datetime.fromisoformat(text)
    except ValueError:
        try:
            t = dt.time.fromisoformat(text)
        except ValueError:
            return ErrorValue.VALUE
        return (t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6) / 86400
    days = date_serial(
        when.year, when.month, when.day, date1904=date1904, leap=quirks.leap_year_1900
    )
    if days is None:
        return ErrorValue.NUM
    seconds = when.hour * 3600 + when.minute * 60 + when.second + when.microsecond / 1e6
    return days + seconds / 86400


def load(
    source: str | os.PathLike[str] | BinaryIO,
    *,
    quirks: Quirks | None = None,
    limits: ReadLimits | None = None,
    trust_cached: bool = True,
    today: dt.date | None = None,
) -> Engine:
    """Read a workbook into an engine; formulas whose cached values are trusted stay unevaluated."""
    with Book(source, limits) as book:
        engine = Engine(
            quirks,
            book.date1904,
            today,
            (limits or ReadLimits()).max_array_cells,
            book.calc.iterate,
        )
        if isinstance(source, (str, os.PathLike)):
            engine.path = os.path.abspath(os.fspath(source))
        for sheet in book.sheets:
            engine.add_sheet(sheet.name)
        for name in book.defined_names:
            engine.define_name(name.name, name.text, name.sheet)
        active = book.active_tab if 0 <= book.active_tab < len(book.sheets) else 0
        row, col = book.active_cell(book.sheets[active])
        engine.active = (active, row - 1, col - 1)
        full = book.calc.full_calc_on_load or not trust_cached
        for index, sheet in enumerate(book.sheets):
            if sheet.kind == "worksheet":
                _load_sheet(engine, book, index, full)
    return engine


def _load_sheet(engine: Engine, book: Book, index: int, full: bool) -> None:
    s = engine.sheets[index]
    masters: dict[int, int] = {}  # shared-formula index -> template id
    refusals: dict[str, object] = {}
    dependents: list[tuple[int, int, int, Value | None]] = []
    report = engine.report

    def held(tid: int, value: Value | None) -> Value | None:
        """What a formula holds until evaluated: its cached value, which a cycle keeps (§7.2)."""
        if full and not engine.iterate and engine.templates[tid].unsupported:
            return None
        return value

    boxes: list[tuple[int, int, int, int, int]] = []  # ranges array formulas fill (§7.7)
    for cell in book.cells(book.sheets[index]):
        row, col = cell.row - 1, cell.column - 1
        cached = _value(cell, engine.date1904, engine.quirks, book.value_errors)
        f = cell.formula
        if f is None:
            anchor = _filled_by(boxes, row, col)
            if anchor is not None:
                engine.set_member(index, row, col, anchor, cached)
            elif cached is not None:
                s.set(row, col, cached)
            continue
        where = CellRef(s.name, cell.row, cell.column)
        if f.kind == "shared" and f.text is None:
            dependents.append((row, col, f.si or 0, cached))
            continue
        box = _box(f.ref) if f.kind in ("array", "dataTable") and f.ref else None
        dynamic = f.kind == "array" and cell.cm is not None and cell.cm in book.dynamic_cm
        if box is not None and not dynamic:  # cells Excel does not let one change one by one
            engine.locked.setdefault(index, []).append(box)
        if f.kind == "dataTable" or f.text is None:
            label = "array formula" if f.kind == "array" else "data table"
            _keep(engine, index, row, col, cached, label, where)
            continue
        if f.kind == "array" and box is not None:
            if refused_in_file(f.text) is not None:
                report.refused[where] = str(refused_in_file(f.text))
                s.set(row, col, BLANK if cached is None else cached)
                continue
            try:
                tid = engine.template_for(f.text, row, col, array_mode=True)
            except (FormulaSyntaxError, TokenizeError) as exc:
                report.refused[where] = str(exc)
                s.set(row, col, BLANK if cached is None else cached)
                continue
            engine.set_formula(
                index,
                row,
                col,
                template=tid,
                cached=held(tid, cached),
                stale=full,
                array=(dynamic, box),
            )
            boxes.append((*box, node(index, row, col)))
            continue
        # Whether Excel refuses a formula depends only on its text outside the references, so
        # it is decided once per raw-text key (§5.3), like the template.
        raw = raw_key(f.text, row, col)
        reason = refusals.get(raw, _UNSEEN)
        if reason is _UNSEEN:
            reason = refusals[raw] = refused_in_file(f.text)
        if reason is not None:
            report.refused[where] = str(reason)
            s.set(row, col, BLANK if cached is None else cached)
            continue
        try:
            tid = engine.template_for(f.text, row, col, raw)
        except (FormulaSyntaxError, TokenizeError) as exc:
            report.refused[where] = str(exc)
            s.set(row, col, BLANK if cached is None else cached)
            continue
        if f.kind == "shared" and f.si is not None:
            masters[f.si] = tid
        engine.set_formula(index, row, col, template=tid, cached=held(tid, cached), stale=full)
    for row, col, si, cached in dependents:
        master = masters.get(si)
        if master is None:  # no master in the file: keep the value, as there is no formula
            if cached is not None:
                s.set(row, col, cached)
            continue
        engine.set_formula(
            index, row, col, template=master, cached=held(master, cached), stale=full
        )


def _box(ref: str) -> tuple[int, int, int, int] | None:
    """A range such as ``B2:C5`` as 0-based (r1, c1, r2, c2)."""
    first, _, last = ref.partition(":")
    try:
        r1, c1 = _cell(first)
        r2, c2 = _cell(last or first)
    except ValueError:
        return None
    return (r1 - 1, c1 - 1, r2 - 1, c2 - 1)


def _keep(
    engine: Engine, sheet: int, row: int, col: int, cached: Any, label: str, where: CellRef
) -> None:
    """A formula the engine cannot evaluate yet: its cached value stays, reported."""
    engine.sheets[sheet].set(row, col, BLANK if cached is None else cached)
    engine.report.unsupported[label] += 1
    if cached is None:
        engine.report.uncomputed.add(where)
