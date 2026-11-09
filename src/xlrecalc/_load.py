"""Build an engine from an xlsx file read by ``_xlsx`` (ARCHITECTURE.md §7.4, §9.1, §9.5).

Cached values are kept and trusted per §7.4: a formula is stale when the file holds no cached
value for it, when the workbook asks for a full calculation on load, or when its template is
volatile; cells depending on a stale cell become stale when the engine recalculates. A formula
Excel refuses in a file is dropped and its cell keeps its cached value, as Excel's repair does
(§9.5), and the report lists it.
"""

from __future__ import annotations

import datetime as dt
import decimal
import os
from collections.abc import Iterable
from typing import Any, BinaryIO

from xlrecalc._dates import date_serial
from xlrecalc._engine import Engine, node
from xlrecalc._formula import FormulaSyntaxError, Kind, Token, TokenizeError, tokenize
from xlrecalc._formula._fastkey import raw_key
from xlrecalc._formula._prefixes import PREFIXES
from xlrecalc._numtext import DBL_MIN
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef
from xlrecalc._values import BLANK, ErrorValue, Value
from xlrecalc._xlsx import Book, Cell, ReadLimits, SheetRows
from xlrecalc._xlsx._cells import _column as _cell

__all__ = ["load", "refused_in_file"]

_UNSEEN = object()

# Error literals a file may hold in a formula (capture formula_forms: #SPILL! and #CALC! are not).
_FILE_ERRORS = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}
)


# The first argument must be a reference or #REF!: a constant, an array or another error makes
# Excel refuse the file (capture info_family, tests/captures/database_refused.tsv).
_REFERENCE_ONLY = frozenset({
    "AREAS", "ISFORMULA", "FORMULATEXT", "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN",
    "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP",
})  # fmt: skip


# An empty argument at these positions makes Excel refuse a file, and F() is position 0 (capture
# empty_args, refused rows in tests/captures/empty_args_refused.tsv).
_EMPTY_REFUSED: dict[str, tuple[int, ...]] = {
    "ABS": (0,),
    "ACOS": (0,),
    "ACOSH": (0,),
    "ACOT": (0,),
    "ACOTH": (0,),
    "ARABIC": (0,),
    "AREAS": (0,),
    "ASIN": (0,),
    "ASINH": (0,),
    "ATAN": (0,),
    "ATANH": (0,),
    "AVERAGEIF": (0, 2),
    "AVERAGEIFS": (0, 1),
    "BIN2DEC": (0,),
    "CELL": (1,),
    "CHAR": (0,),
    "CLEAN": (0,),
    "CODE": (0,),
    "COLUMNS": (0,),
    "COS": (0,),
    "COSH": (0,),
    "COT": (0,),
    "COTH": (0,),
    "COUNTBLANK": (0,),
    "COUNTIF": (0,),
    "COUNTIFS": (0,),
    "CSC": (0,),
    "CSCH": (0,),
    "DATEVALUE": (0,),
    "DAVERAGE": (0,),
    "DAY": (0,),
    "DCOUNT": (0,),
    "DCOUNTA": (0,),
    "DEGREES": (0,),
    "DGET": (0,),
    "DMAX": (0,),
    "DMIN": (0,),
    "DPRODUCT": (0,),
    "DSTDEV": (0,),
    "DSTDEVP": (0,),
    "DSUM": (0,),
    "DVAR": (0,),
    "DVARP": (0,),
    "ERROR.TYPE": (0,),
    "EVEN": (0,),
    "EXP": (0,),
    "FACT": (0,),
    "FACTDOUBLE": (0,),
    "FISHER": (0,),
    "FISHERINV": (0,),
    "FORMULATEXT": (0,),
    "GAUSS": (0,),
    "HEX2DEC": (0,),
    "HOUR": (0,),
    "INT": (0,),
    "ISBLANK": (0,),
    "ISERR": (0,),
    "ISERROR": (0,),
    "ISEVEN": (0,),
    "ISFORMULA": (0,),
    "ISLOGICAL": (0,),
    "ISNA": (0,),
    "ISNONTEXT": (0,),
    "ISNUMBER": (0,),
    "ISODD": (0,),
    "ISOWEEKNUM": (0,),
    "ISREF": (0,),
    "ISTEXT": (0,),
    "LEN": (0,),
    "LENB": (0,),
    "LN": (0,),
    "LOG10": (0,),
    "LOWER": (0,),
    "MAXIFS": (0, 1),
    "MINIFS": (0, 1),
    "MINUTE": (0,),
    "MONTH": (0,),
    "N": (0,),
    "NORM.S.INV": (0,),
    "NORMSDIST": (0,),
    "NORMSINV": (0,),
    "NOT": (0,),
    "OCT2DEC": (0,),
    "ODD": (0,),
    "OFFSET": (0,),
    "PHI": (0,),
    "PROPER": (0,),
    "RADIANS": (0,),
    "RANK": (1,),
    "RANK.AVG": (1,),
    "RANK.EQ": (1,),
    "ROWS": (0,),
    "SEC": (0,),
    "SECH": (0,),
    "SECOND": (0,),
    "SIGN": (0,),
    "SIN": (0,),
    "SINH": (0,),
    "SQRT": (0,),
    "SQRTPI": (0,),
    "SUBTOTAL": (1, 2),
    "SUMIF": (0, 2),
    "SUMIFS": (0, 1),
    "T": (0,),
    "TAN": (0,),
    "TANH": (0,),
    "TIMEVALUE": (0,),
    "TRANSPOSE": (0,),
    "TRIM": (0,),
    "TYPE": (0,),
    "UNICHAR": (0,),
    "UNICODE": (0,),
    "UPPER": (0,),
    "VALUE": (0,),
    "YEAR": (0,),
}


# A repeating tail: arguments before it, and the size of a group. A position in a later group
# counts as the same position of the first (capture empty_args).
_EMPTY_REPEATING = {
    "COUNTIFS": (0, 2), "SUMIFS": (1, 2), "AVERAGEIFS": (1, 2), "MAXIFS": (1, 2), "MINIFS": (1, 2),
    "IFS": (0, 2), "SWITCH": (1, 2),
}  # fmt: skip


def _empty_is_refused(name: str, position: int) -> bool:
    head, size = _EMPTY_REPEATING.get(name, (position + 1, 1))
    at = position if position < head else head + (position - head) % size
    return at in _EMPTY_REFUSED.get(name, ())


def _refused_empty(toks: Iterable[Token]) -> str | None:
    """The function whose empty argument Excel refuses, as a reason, or ``None``."""
    stack: list[list[Any]] = []  # [function or None, position, nothing seen in this argument]
    previous = None
    for tok in toks:
        if stack and tok.kind is Kind.COMMA and stack[-1][0] is not None:
            if stack[-1][2] and _empty_is_refused(stack[-1][0], stack[-1][1]):
                return f"an empty argument of {stack[-1][0]}"
            stack[-1][1] += 1
            stack[-1][2] = True
            previous = tok
            continue
        if tok.kind is Kind.CLOSE and stack:
            closed, position, empty = stack.pop()
            if closed is not None and empty and _empty_is_refused(closed, position):
                return f"an empty argument of {closed}"
        if stack and tok.kind is not Kind.COMMA:
            stack[-1][2] = False
        if tok.kind is Kind.OPEN:
            name = None
            if previous is not None and previous.kind is Kind.FUNCTION:
                name = previous.text.upper().removeprefix("_XLFN.")
            stack.append([name, 0, True])
        previous = tok
    return None


# Argument counts of the functions that make arrays, outside which Excel refuses the file
# (capture dyn_*); SORTBY takes an array and pairs of a key and its order, the last order optional.
_ARRAY_COUNTS: dict[str, tuple[int, int | None]] = {
    "_XLFN.SEQUENCE": (1, 4), "_XLFN.UNIQUE": (1, 3), "_XLFN._XLWS.SORT": (1, 4),
    "_XLFN._XLWS.FILTER": (2, 3), "MMULT": (2, 2),
    "_XLFN.CHOOSECOLS": (2, None),
    "_XLFN.CHOOSEROWS": (2, None),
    "_XLFN.TAKE": (2, 3),
    "_XLFN.DROP": (2, 3),
    "_XLFN.EXPAND": (2, 4),
    "_XLFN.VSTACK": (1, None),
    "_XLFN.HSTACK": (1, None),
    "_XLFN.TOCOL": (1, 3),
    "_XLFN.TOROW": (1, 3),
    "_XLFN.WRAPROWS": (2, 3),
    "_XLFN.WRAPCOLS": (2, 3),
}  # fmt: skip
# Without its prefix in a file, SORT or FILTER makes Excel refuse it; UNIQUE, SORTBY and
# SEQUENCE are names it does not know (capture dyn_*).
_BARE_REFUSED = frozenset({"SORT", "FILTER"})


def _count_refused(name: str, count: int) -> bool:
    if name == "_XLFN.SORTBY":
        return count < 2 or (count > 3 and count % 2 == 0)
    bounds = _ARRAY_COUNTS.get(name)
    if bounds is None:
        return False
    lo, hi = bounds
    return count < lo or (hi is not None and count > hi)


def _refused_count(toks: Iterable[Token]) -> str | None:
    """The function called with a number of arguments Excel refuses, as a reason, or ``None``."""
    stack: list[list[Any]] = []  # [function or None, commas, anything seen]
    previous = None
    braces = 0  # inside an array constant, a comma separates elements
    for tok in toks:
        if tok.kind is Kind.ARRAY_OPEN:
            braces += 1
        elif tok.kind is Kind.ARRAY_CLOSE:
            braces -= 1
        if braces and tok.kind is Kind.COMMA:
            continue
        if tok.kind is Kind.FUNCTION and tok.text.upper() in _BARE_REFUSED:
            bare = tok.text.upper()
            prefix = PREFIXES[bare]
            return f"{bare} written without Excel's file prefix ({prefix}): Excel refuses the file"
        if tok.kind is Kind.OPEN:
            name = None
            if previous is not None and previous.kind is Kind.FUNCTION:
                name = previous.text.upper()
            stack.append([name, 0, False])
        elif tok.kind is Kind.CLOSE and stack:
            name, commas, seen = stack.pop()
            count = commas + 1 if (commas or seen) else 0
            if name is not None and _count_refused(name, count):
                return f"{name} with {count} arguments"
        elif stack:
            if tok.kind is Kind.COMMA:
                stack[-1][1] += 1
            else:
                stack[-1][2] = True
        previous = tok
    return None


def _out_of_range(text: str) -> bool:
    """A number literal Excel refuses in a file (capture ``number_literals``).

    Its exponent in scientific form must lie between -309 and 307: Excel keeps 15 significant
    digits, cut, not rounded (``9.9999999999999999E+307`` is accepted), which leaves the exponent
    as written, so ``1E+308`` and ``1E-310`` are refused and ``1E-309`` is accepted (as 0).
    """
    try:
        number = decimal.Decimal(text)
    except decimal.InvalidOperation:
        return False
    return bool(number) and not -309 <= number.adjusted() <= 307


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
        if tok.kind is Kind.NUMBER and _out_of_range(tok.text):
            return f"the number {tok.text}, outside 1E-309 to 9.99999999999999E+307"
        if (  # OFFSET(5,0,0): capture dynrefs, the file does not open
            tok.kind in (Kind.NUMBER, Kind.STRING, Kind.BOOL)
            and previous is not None
            and previous.kind is Kind.OPEN
            and before is not None
            and before.kind is Kind.FUNCTION
            and before.text.upper() == "OFFSET"
        ):
            return "a constant as OFFSET's reference"
        if (  # AREAS(1), ISFORMULA({1,2}), DSUM(1,"a",A1:A2): the file does not open
            (
                tok.kind in (Kind.NUMBER, Kind.STRING, Kind.BOOL, Kind.ARRAY_OPEN)
                or (tok.kind is Kind.ERROR and tok.text != "#REF!")
            )
            and previous is not None
            and previous.kind is Kind.OPEN
            and before is not None
            and before.kind is Kind.FUNCTION
            and before.text.upper().removeprefix("_XLFN.") in _REFERENCE_ONLY
        ):
            return f"an argument of {before.text} that is no reference"
        if (
            tok.kind is Kind.CLOSE
            and previous is not None
            and previous.kind is Kind.OPEN
            and tok.space
        ):
            return "an argument list of spaces"
        before, previous = previous, tok
    return _refused_empty(toks) or _refused_count(toks)


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


def _iterate_count(count: int | None) -> int:
    """The file's ``iterateCount`` in Excel's range, 1 to 32,767; 100 when absent or unreadable."""
    return 100 if count is None or count < 1 else min(count, 32_767)


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
            _iterate_count(book.calc.iterate_count),
            book.calc.iterate_delta if book.calc.iterate_delta is not None else 0.001,
            (limits or ReadLimits()).max_iteration_evaluations,
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
    boxes: list[tuple[int, int, int, int, int]] = []  # ranges array formulas fill (§7.7)
    rows = SheetRows()
    for cell in book.cells(book.sheets[index], rows):
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
                cached=cached,
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
        engine.set_formula(index, row, col, template=tid, cached=cached, stale=full)
        if engine.iterate and f.attributes.get("ca") in ("1", "true"):
            engine.calc_always.add(node(index, row, col))
    s.hidden_rows = frozenset(r - 1 for r in rows.hidden)
    s.filter_mode = rows.filter_mode
    for row, col, si, cached in dependents:
        master = masters.get(si)
        if master is None:  # no master in the file: keep the value, as there is no formula
            if cached is not None:
                s.set(row, col, cached)
            continue
        engine.set_formula(index, row, col, template=master, cached=cached, stale=full)


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
