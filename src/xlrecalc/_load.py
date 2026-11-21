"""Build an engine from an xlsx file read by ``xlio`` (ARCHITECTURE.md §7.4, §9.1, §9.5).

Cached values are kept and trusted per §7.4: a formula is stale when the file holds no cached
value for it, or when Excel recalculates it on opening (``recalc_on_load``, ``_opening.Rule``);
cells depending on a stale cell become stale when the engine recalculates. A formula Excel
refuses in a file is dropped and its cell keeps its cached value, as Excel's repair does
(§9.5), and the report lists it.
"""

from __future__ import annotations

import datetime as dt
import decimal
import itertools
import os
import warnings
from array import array
from collections.abc import Generator, Iterable, Iterator
from typing import Any, BinaryIO, Literal

import xlio

from xlrecalc._dates import date_serial
from xlrecalc._engine import Engine, node
from xlrecalc._formula import FormulaSyntaxError, Kind, Token, TokenizeError, parse, tokenize
from xlrecalc._formula._fastkey import raw_key
from xlrecalc._formula._nodes import Call, Reference
from xlrecalc._formula._prefixes import PREFIXES
from xlrecalc._limits import ReadLimits, io_limits
from xlrecalc._opening import Opening, Rule, Volatility
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef
from xlrecalc._store import COLS
from xlrecalc._values import BLANK, ErrorValue, Value

__all__ = ["defined_names", "file_app", "load", "refused_in_file"]

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
    "LAMBDA": (0,),  # any empty parameter or body (capture empty_args, with _EMPTY_REPEATING)
    "ISOMITTED": (0,),
    "ERF.PRECISE": (0,),
    "ERFC": (0,),
    "ERFC.PRECISE": (0,),
    "GAMMA": (0,),
    "GAMMALN": (0,),
    "GAMMALN.PRECISE": (0,),
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
    "IMABS": (0,),
    "IMAGINARY": (0,),
    "IMARGUMENT": (0,),
    "IMCONJUGATE": (0,),
    "IMCOS": (0,),
    "IMCOSH": (0,),
    "IMCOT": (0,),
    "IMCSC": (0,),
    "IMCSCH": (0,),
    "IMEXP": (0,),
    "IMLN": (0,),
    "IMLOG10": (0,),
    "IMLOG2": (0,),
    "IMREAL": (0,),
    "IMSEC": (0,),
    "IMSECH": (0,),
    "IMSIN": (0,),
    "IMSINH": (0,),
    "IMSQRT": (0,),
    "IMTAN": (0,),
}


# A repeating tail: arguments before it, and the size of a group. A position in a later group
# counts as the same position of the first (capture empty_args).
_EMPTY_REPEATING = {
    "COUNTIFS": (0, 2), "SUMIFS": (1, 2), "AVERAGEIFS": (1, 2), "MAXIFS": (1, 2), "MINIFS": (1, 2),
    "IFS": (0, 2), "SWITCH": (1, 2), "LAMBDA": (0, 1),
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
    "_XLFN.MUNIT": (1, 1),
    "MDETERM": (1, 1),
    "MINVERSE": (1, 1),
    "FREQUENCY": (2, 2),
    "_XLFN.MODE.MULT": (1, None),
    "_XLFN.TEXTSPLIT": (2, 6),
    "_XLFN.RANDARRAY": (0, 5),
    "_XLFN.REDUCE": (3, None),  # two refused (refusal_probe, Excel 16.115); more not captured
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
    return _refused_empty(toks) or _refused_count(toks) or _refused_let(text)


def _refused_let(text: str) -> str | None:
    """A ``LET`` Excel refuses in a file (capture ``dyn_let``, each form alone in a workbook).

    A name must be written with ``_xlpm.`` (``_xlfn.LET(x,1,x+1)`` is refused), be a name (not
    ``1``) and be bound once in its call. A call with fewer than two arguments is refused too,
    uncaptured, so the cell keeps the value it holds.
    """
    if "_XLFN.LET(" not in text.upper():
        return None
    try:
        root = parse(text).root
    except (FormulaSyntaxError, TokenizeError):
        return None
    stack: list[Any] = [root]
    while stack:
        item = stack.pop()
        if isinstance(item, list):
            stack.extend(item)
            continue
        if not hasattr(item, "__dataclass_fields__") or isinstance(item, Token):
            continue
        stack.extend(getattr(item, f) for f in item.__dataclass_fields__)
        if not (isinstance(item, Call) and item.prefix is None):
            continue
        if item.name.text.upper() != "_XLFN.LET":
            continue
        if len(item.args) < 2:
            return f"LET with {len(item.args)} argument{'s' * (len(item.args) != 1)}"
        pairs = item.args if len(item.args) % 2 == 0 else item.args[:-1]
        seen: set[str] = set()
        for name in pairs[::2]:
            if not (
                isinstance(name, Reference)
                and name.prefix is None
                and name.token.kind is Kind.NAME
                and name.token.text[:6].casefold() == "_xlpm."
            ):
                return "a LET name that is no name written with _xlpm."
            folded = name.token.text.casefold()
            if folded in seen:
                return f"a LET name bound twice ({name.token.text})"
            seen.add(folded)
    return None


def _filled_by(boxes: list[tuple[int, int, int, int, int]], row: int, col: int) -> int | None:
    """The array formula whose range holds the cell, if any."""
    for r1, c1, r2, c2, anchor in boxes:
        if r1 <= row <= r2 and c1 <= col <= c2:
            return anchor
    return None


def _value(cell: xlio.Cell, date1904: bool, quirks: Quirks) -> Value | None:
    """A cell's cached value as the engine holds it; ``None`` when there is none.

    ``xlio`` reads values as Excel does: a subnormal number is 0, and an error a file caches as
    ``#VALUE!`` with a rich value behind it (``vm``) is that error (a blocked spill is
    ``#SPILL!``: capture ``arrays``).
    """
    t, v = cell.type, cell.value
    if t == "none" or v is None:
        return None
    if t == "n":
        assert isinstance(v, float)
        return v
    if t == "b":
        return bool(v)
    if t == "e":
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


RecalcOnLoad = Literal["excel", "always", "never"]
RECALC_ON_LOAD = ("excel", "always", "never")


def recalc_mode(
    recalc_on_load: RecalcOnLoad, alias: str, value: bool | None, *, stacklevel: int
) -> RecalcOnLoad:
    """``recalc_on_load``, or what a deprecated alias given instead maps to, with its warning.

    ``trust_cached=False`` and ``full=True`` map to ``"always"``; ``trust_cached=True`` and
    ``full=False`` to ``"excel"``, the default. Both aliases will be removed in 1.0.
    """
    if value is None:
        return recalc_on_load
    mode: RecalcOnLoad = "always" if (value if alias == "full" else not value) else "excel"
    warnings.warn(
        f'{alias} is deprecated and will be removed in 1.0: use recalc_on_load="{mode}"',
        DeprecationWarning,
        stacklevel=stacklevel + 1,
    )
    return mode


class _Formulas:
    """What a load's formulas showed, for ``Rule.decide``: ``RecalcCheck.add``'s four facts.

    Taken from each formula cell's text as it is read, before anything is compiled; a formula
    Excel refuses in a file is left out, since Excel refuses the file. ``marked`` lists the cells
    marked ``ca="1"`` that the compile made formulas, to be made stale.
    """

    __slots__ = (
        "_names",
        "_own",
        "_volatility",
        "ca",
        "marked",
        "marked_stale",
        "own_unmarked",
        "uncached",
        "volatile",
    )

    def __init__(self, marked_stale: bool, volatility: Volatility) -> None:
        self.marked_stale = marked_stale  # a cell marked ca="1" is computed (``Rule.marked``)
        self.marked: list[int] = []  # those cells, made stale once the formulas are compiled
        self.uncached = False  # a formula cell holds no cached value
        self.volatile = False  # a formula is volatile, through a defined name too
        self.ca = False  # a formula cell is marked ca="1"
        self.own_unmarked = False  # a volatile function in a formula's own text, unmarked
        self._volatility = volatility
        self._own: dict[object, bool] = {}  # key -> calls a volatile function in its own text
        self._names: dict[object, bool] = {}  # key -> volatile, through a defined name too

    def note(self, key: object, text: str, marked: bool) -> None:
        """A formula cell holding its own text, as ``RecalcCheck.add`` notes it.

        Decided from the text alone, once per ``key`` (the raw-text key, or an array formula's
        text in a tuple), so that a formula the parser cannot read counts too.
        """
        if not marked and not self.own_unmarked:
            own = self._own.get(key)
            if own is None:
                own = self._own[key] = self._volatility.own(text)
            if own:
                self.own_unmarked = self.volatile = True
                return
        if not self.volatile:
            volatile = self._names.get(key)
            if volatile is None:
                volatile = self._names[key] = self._volatility(text)
            self.volatile = volatile

    def decide(self, rule: Rule) -> Opening:
        """The rule's answer from what the formulas showed so far."""
        return rule.decide(self.uncached, self.volatile, self.ca, self.own_unmarked)


def defined_names(book: xlio.Workbook) -> list[tuple[str, str, int | None]]:
    """The workbook's defined names as (name, text, scope), scope the sheet's index or ``None``.

    ``localSheetId`` counts every ``<sheet>``, one without a part too: a name local to that is
    dropped.
    """
    index_of = {sheet.position: i for i, sheet in enumerate(book.sheets)}
    names = []
    for name in book.defined_names:
        scope = None if name.sheet is None else index_of.get(name.sheet)
        if name.sheet is None or scope is not None:
            names.append((name.name, name.text, scope))
    return names


def file_app(book: xlio.Workbook) -> str | None:
    """``<fileVersion appName>``: ``"xl"`` when Excel saved the file."""
    return next(
        (u.element.get("appName") for u in book.unparsed if u.path == "workbook/fileVersion"),
        None,
    )


def _iterate_count(count: int | None) -> int:
    """The file's ``iterateCount`` in Excel's range, 1 to 32,767; 100 when absent or unreadable."""
    return 100 if count is None or count < 1 else min(count, 32_767)


def load(
    source: str | os.PathLike[str] | BinaryIO,
    *,
    quirks: Quirks | None = None,
    limits: ReadLimits | None = None,
    recalc_on_load: RecalcOnLoad = "excel",
    today: dt.date | None = None,
    saving: bool = False,
) -> Engine:
    """Read a workbook into an engine; formulas whose cached values are trusted stay unevaluated.

    ``recalc_on_load``: ``"excel"`` evaluates what Excel recalculates on opening (``Rule``: every
    formula, the cells marked ``ca="1"``, or in manual calculation none unless ``saving``) and
    trusts the cached values elsewhere; ``"always"`` evaluates every formula; ``"never"`` trusts
    every cached value the file holds but the volatile formulas'. A formula without one is
    evaluated in every mode. ``saving``: the load is followed by a save, as in
    ``recalculate(path)``, whose result is the copy Excel saves after opening.
    """
    if recalc_on_load not in RECALC_ON_LOAD:  # checked before the file is opened
        raise ValueError(f"recalc_on_load must be one of {RECALC_ON_LOAD}: {recalc_on_load!r}")
    with xlio.open(source, limits=io_limits(limits)) as book:
        path = (
            os.path.abspath(os.fspath(source)) if isinstance(source, (str, os.PathLike)) else None
        )
        loader = _Loader(
            book,
            quirks=quirks,
            limits=limits,
            recalc_on_load=recalc_on_load,
            today=today,
            saving=saving,
            path=path,
        )
        for sheet in book.sheets:
            for _ in loader.cells(sheet):
                pass
        engine = loader.finish()
        engine.load_all()  # load's engine is compiled, as every caller of it expects
        return engine


class _Loader:
    """A workbook read into an engine from an open ``xlio`` workbook, its cells fed by the caller.

    The workbook's own parts (names, calculation properties, views) are read at once. Each sheet's
    cells pass through :meth:`cells`, in any order, while they are recorded (``_load_sheet``);
    :meth:`settle` reads every sheet not fed or fed part-way, compiles every sheet's formulas
    (``_compile_sheet``) in the sheets' order, and applies the opening rule. A sheet fed after
    that is passed through unrecorded. ``load`` is this with every sheet fed in order.
    """

    def __init__(
        self,
        book: xlio.Workbook,
        *,
        quirks: Quirks | None = None,
        limits: ReadLimits | None = None,
        recalc_on_load: RecalcOnLoad = "excel",
        today: dt.date | None = None,
        saving: bool = False,
        path: str | None = None,
    ) -> None:
        if recalc_on_load not in RECALC_ON_LOAD:
            raise ValueError(f"recalc_on_load must be one of {RECALC_ON_LOAD}: {recalc_on_load!r}")
        engine = Engine(
            quirks,
            book.date1904,
            today,
            (limits or ReadLimits()).max_array_cells,
            book.calc.iterate,
            _iterate_count(book.calc.iterate_count),
            book.calc.iterate_delta if book.calc.iterate_delta is not None else 0.001,
            (limits or ReadLimits()).max_iteration_evaluations,
            (limits or ReadLimits()).max_function_steps,
        )
        if path is not None:
            engine.path = path
        for sheet in book.sheets:
            engine.add_sheet(sheet.name)
        for name, text, scope in defined_names(book):
            engine.define_name(name, text, scope)
        active = book.views[0].active_tab if book.views else 0
        if not 0 <= active < len(book.sheets):
            active = 0
        row, col = _active_cell(book.sheets[active])
        engine.active = (active, row - 1, col - 1)
        self.engine = engine
        self._book = book
        self._rule = Rule.of(book.calc, file_app(book), recalc_on_load, saving)
        self._seen = _Formulas(self._rule.marked, engine.volatility)
        self._dynamic = book.cell_metadata.dynamic_arrays
        self._sheets: dict[int, _SheetLoad] = {}
        self._fed: set[int] = set()  # the sheets the caller has fed
        self._settled = False
        self._scanned: set[int] = set()  # sheets whose formulas were scanned (_decided)
        self._late = "everything"  # how a sheet loads after settle: _load
        self._full = False  # every formula stale, a late sheet's too (settle)
        self._trust = False  # volatile cells keep their cached values (settle)
        self._finished = False

    def _index(self, sheet: xlio.Sheet) -> int:
        for index, candidate in enumerate(self._book.sheets):
            if candidate is sheet:
                return index
        raise ValueError(f"sheet {sheet.name!r} is not a sheet of this workbook")

    def _state(self, index: int) -> _SheetLoad:
        state = self._sheets[index] = _SheetLoad(
            self.engine,
            self._book.sheets[index],
            self._dynamic,
            index,
            self._rule.everything or self._full,
            self._seen,
        )
        return state

    def cells(self, sheet: xlio.Sheet, *, rich: bool = False) -> Iterator[xlio.Cell]:
        """Each cell of the sheet, recorded as it passes unless the loader has read it already.

        A sheet is fed once, in any order and at any time. ``rich`` is passed to ``sheet.cells``:
        it adds a string's runs and phonetic guides to the cells handed on, and changes nothing
        recorded (a cell's value is its text either way).
        """
        index = self._index(sheet)
        if index in self._fed:
            raise ValueError(f"sheet {sheet.name!r} has already been fed")
        self._fed.add(index)
        if sheet.kind != "worksheet":
            self._sheets.setdefault(index, _SheetLoad.done())
            return iter(())
        if index in self._sheets:  # read by the loader for itself: passed through
            return sheet.cells(rich=rich)
        state = self._state(index)
        cells = iter(sheet.cells(rich=rich))
        state.running = _load_sheet(state, cells)
        return _passing(state.running, cells)

    @property
    def opening(self) -> Opening | None:
        """``Rule.decide``'s answer, from the cells recorded and a formulas-only scan of the rest.

        The sheets not recorded are scanned (``_scan``) while the book is open; ``None`` once it
        is closed with the answer still depending on a sheet not read.
        """
        return self._decided()

    def _decided(self, *, some_is_enough: bool = False) -> Opening | None:
        """The rule's answer, scanning sheets not recorded or scanned yet while it is open.

        ``"all"`` is final as soon as it shows; with ``some_is_enough``, ``"some"`` stops the
        scan (:meth:`settle` then reads every sheet).
        """
        answer = self._seen.decide(self._rule)
        if answer == "all" or self._settled:
            return answer
        for index, sheet in enumerate(self._book.sheets):
            if sheet.kind != "worksheet" or index in self._scanned:
                continue
            state = self._sheets.get(index)
            if state is not None and state.recorded:
                continue
            if self._book.closed:
                return None
            _scan(self._seen, sheet)
            self._scanned.add(index)
            if some_is_enough and self._seen.decide(self._rule) == "some":
                return "some"
        return self._seen.decide(self._rule)

    def _read(self, index: int) -> None:
        """Record the rest of a sheet not fed, or fed part-way, reading it again from its start.

        A loop of the caller's still running over it records no more: it hands on the rest of
        its cells unrecorded.
        """
        state = self._sheets.get(index) or self._state(index)
        if state.recorded:
            return
        if state.running is not None:
            state.running.close()  # sets state.fed to the cells it recorded
            state.running = None
        sheet = self._book.sheets[index]
        if self._book.closed:
            raise ValueError(
                f"sheet {sheet.name!r} is needed but not read yet, and the workbook is closed"
            )
        for _ in _load_sheet(state, itertools.islice(sheet.cells(), state.fed, None)):
            pass

    def settle(self) -> Engine:
        """The engine, the opening rule applied; made once.

        A sheet not recorded yet is left to load when an evaluation names it (``Engine.need``)
        where no cell of it can change how the others load, which the rule's answer tells,
        without iterative calculation (a run covers every loop):

        - ``"all"``, from the file alone (``Rule.everything``) or from a volatile function a
          formula calls in its own text without ``ca="1"``: every formula is stale, a late
          sheet's too;
        - ``"none"``: no formula is stale, a late sheet's neither; a sheet's formulas, fed or
          not, are compiled only when they are needed (``Engine.need_formulas``), since every
          value is the cached one.

        For ``"some"`` every sheet is read and compiled first, as the stale cells' dependents may
        be anywhere. Sheets compiled now are compiled in the sheets' order.
        """
        if self._settled:
            return self.engine
        engine, rule, seen = self.engine, self._rule, self._seen
        answer = None if engine.iterate else self._decided(some_is_enough=True)
        lazy = answer in ("all", "none")
        sheets = self._book.sheets
        later: set[int] = set()
        for index, sheet in enumerate(sheets):
            if sheet.kind == "worksheet":
                state = self._sheets.get(index)
                if lazy and (state is None or not state.recorded):
                    later.add(index)
                else:
                    self._read(index)
        full = rule.everything or answer == "all"
        for state in self._sheets.values():
            if not state.compiled:
                state.full = full  # a late sheet's formulas are stale where every one is
        self._full = full
        uncompiled: set[int] = set()
        for index, sheet in enumerate(sheets):
            if sheet.kind == "worksheet" and index not in later:
                if answer == "none":
                    uncompiled.add(index)
                else:
                    _compile_sheet(self._sheets[index])
        self._settled = True
        self._late = "values" if answer == "none" else "everything"
        if later or uncompiled:
            engine._unloaded, engine._uncompiled = later, uncompiled
            engine.load_sheet, engine.compile_sheet = self._load, self._compile
        self._trust = not rule.everything and not rule.volatile
        if self._trust:
            engine.trust_volatile()  # kept on opening unless marked ca="1" (or manual)
        engine.stale.update(seen.marked)  # what Excel recalculates on opening
        if not rule.everything and seen.decide(rule) == "all":
            engine.make_all_stale()  # a volatile function without ca="1": every formula
        engine.loaded()
        return engine

    def finish(self) -> Engine:
        """:meth:`settle`, once, with every sheet read.

        Every sheet is compiled too, unless no formula is computed (``opening`` is ``"none"``).
        """
        if self._finished:
            raise ValueError("the loader has already finished")
        self._finished = True
        engine = self.settle()
        engine.read_all()
        return engine

    def _load(self, index: int) -> None:
        """A sheet an evaluation needs, after :meth:`settle`: read (the rest), then compiled.

        Where the cached values stand, compiling waits for :meth:`_compile`.
        """
        self._read(index)
        if self._late == "values":
            self.engine._uncompiled.add(index)
        else:
            self._compile(index)

    def _compile(self, index: int) -> None:
        """A loaded sheet's formulas, compiled as at :meth:`settle`."""
        state = self._sheets[index]
        if state.compiled:
            return
        changed = self.engine._changed
        _compile_sheet(state)
        self.engine._changed = changed  # a sheet loaded late is the load's, not a change
        if self._trust and not self._full:
            self.engine.trust_volatile()  # as at settle, for the cells compiled now


def _passing(
    recording: Generator[xlio.Cell, None, None], cells: Iterator[xlio.Cell]
) -> Generator[xlio.Cell, None, None]:
    """The caller's loop over a sheet: recorded while ``recording`` runs, then passed through.

    The loader closes ``recording`` when it reads the sheet itself; the caller still receives every
    cell, the rest straight from ``cells``. A sheet recorded after :meth:`_Loader.settle` is
    compiled when an evaluation first needs it (``_Loader._load``).
    """
    yield from recording
    yield from cells


def _active_cell(sheet: xlio.Sheet) -> tuple[int, int]:
    """A worksheet's active cell in its first view, 1-based (row, col); A1 without a view.

    As Excel reads it (``xlio``'s capture ``views``): ``activeCell`` when it lies in the
    selection, else the selection's first cell.
    """
    if sheet.kind != "worksheet" or not sheet.views:
        return 1, 1
    cell = sheet.views[0].active_cell
    return cell.row, cell.column


_Queued = tuple[xlio.Formula, str | None, bool, object]
"""A formula queued with what compiling it needs: (formula, raw-text key, dynamic, refusal)."""


class _SheetLoad:
    """What one worksheet's loading keeps between its cells, so that it can stop and go on.

    Loading has two stages. Recording (``_load_sheet``) stores each cell's value, a formula's
    cached value included, and queues each formula cell; compiling (``_compile_sheet``) then makes
    the queued formulas the engine's, in the order read. The queue is kept small, since a sheet
    may hold millions of formulas: ``cells`` holds each queued cell's store key, and ``entries``
    what compiling it needs. Most entries are a formula repeating the raw-text key (§5.3) of one
    queued before it on the sheet that parses: the key itself (a string shared by them all), or
    the key in a 1-tuple when the cell is marked ``ca="1"``. Any other formula is ``(formula,
    raw-text key or None, dynamic, refusal)``; ``None`` is a cell without a value in an array
    formula's range. A shared formula's dependents wait apart (``dependents``, ``sis``,
    ``marked_dependents``) until every master is known.

    ``fed`` counts the cells recorded. A reading the caller left part-way is continued by reading
    the sheet again from its start and recording from cell ``fed`` on; ``running`` is the caller's
    recording loop while it may still go on.
    """

    def __init__(
        self,
        engine: Engine,
        sheet: xlio.Sheet,
        dynamic_cm: frozenset[int],
        index: int,
        full: bool,
        seen: _Formulas,
    ) -> None:
        self.engine = engine
        self.sheet = sheet
        self.dynamic_cm = dynamic_cm
        self.index = index
        self.full = full
        self.seen = seen
        self.refusals: dict[str, object] = {}  # raw-text key -> refused_in_file's reason
        self.keys: dict[str, str | None] = {}  # raw-text key -> itself if it parses, else None
        self.cells = array("q")
        self.entries: list[str | tuple[str] | _Queued | None] = []
        self.dependents = array("q")
        self.sis = array("q")
        self.marked_dependents: set[int] = set()
        self.ranges: list[tuple[int, int, int, int, int]] = []  # array formulas' ranges, as read
        self.fed = 0
        self.running: Generator[xlio.Cell, None, None] | None = None
        self.recorded = False
        self.compiled = False

    @classmethod
    def done(cls) -> _SheetLoad:
        """The state of a sheet with no cells to record (a chart sheet)."""
        state = cls.__new__(cls)
        state.fed, state.running, state.recorded, state.compiled = 0, None, True, True
        return state


def _load_sheet(state: _SheetLoad, cells: Iterable[xlio.Cell]) -> Generator[xlio.Cell, None, None]:
    """Record a worksheet's cells, yielding each one first; after the last, the sheet is recorded.

    A value is stored at once, a formula's cached value too; a formula cell, and a cell without a
    value in an array formula's range, is queued for ``_compile_sheet`` (``_SheetLoad``), with
    what the opening rule needs taken from its text (``_Formulas.note``). A cell is recorded when
    the caller asks for the next one, so a caller that stops after receiving a cell leaves that
    cell unrecorded, and ``state.fed`` counts only what was recorded.
    """
    engine, seen, dynamic_cm = state.engine, state.seen, state.dynamic_cm
    s = engine.sheets[state.index]
    queued, entries, refusals, keys, ranges = (
        state.cells,
        state.entries,
        state.refusals,
        state.keys,
        state.ranges,
    )
    fed = state.fed
    try:
        for cell in cells:
            yield cell
            fed += 1
            row, col = cell.row - 1, cell.column - 1
            cached = _value(cell, engine.date1904, engine.quirks)
            f = cell.formula
            if cached is not None:
                s.set(row, col, cached)
            if f is None:
                if cached is None and ranges and _filled_by(ranges, row, col) is not None:
                    queued.append(row * COLS + col)
                    entries.append(None)
                continue
            marked = f.attributes.get("ca") in ("1", "true")
            raw, reason, box = _formula_facts(seen, refusals, f, row, col, cached is None, marked)
            if f.kind == "shared" and f.text is None:
                state.dependents.append(row * COLS + col)
                state.sis.append(f.si or 0)
                if marked:
                    state.marked_dependents.add(row * COLS + col)
                continue
            if box is not None:
                ranges.append((*box, node(state.index, row, col)))
            elif raw is not None and reason is None and f.kind != "shared":
                if raw not in keys:  # the first: queued with its text, to compile
                    keys[raw] = raw if _parses(f.text) else None  # type: ignore[arg-type]
                elif (known := keys[raw]) is not None:  # a repeat of a key that compiles
                    queued.append(row * COLS + col)
                    entries.append((known,) if marked else known)
                    continue
            queued.append(row * COLS + col)
            entries.append((f, raw, f.kind == "array" and cell.cm in dynamic_cm, reason))
    finally:
        state.fed = fed
    s.hidden_rows = frozenset(r - 1 for r, f in state.sheet.rows.items() if f.hidden)
    # filterMode only decides whether SUBTOTAL leaves out hidden rows: with none hidden it changes
    # nothing, and reading it costs xlio a second parse of the part's head
    s.filter_mode = bool(s.hidden_rows) and state.sheet.properties.filter_mode
    state.recorded = True


def _scan(seen: _Formulas, sheet: xlio.Sheet) -> None:
    """Note the opening rule's facts from a sheet's formulas alone (``xlio``'s ``Sheet.formulas``).

    As recording notes them (``_formula_facts``).
    """
    refusals: dict[str, object] = {}
    for cell in sheet.formulas():
        f = cell.formula
        marked = f.attributes.get("ca") in ("1", "true")
        _formula_facts(seen, refusals, f, cell.row - 1, cell.column - 1, not cell.cached, marked)


def _formula_facts(
    seen: _Formulas,
    refusals: dict[str, object],
    f: xlio.Formula,
    row: int,
    col: int,
    uncached: bool,
    marked: bool,
) -> tuple[str | None, object, tuple[int, int, int, int] | None]:
    """Note what the opening rule needs from one formula cell.

    Recording and the formulas-only scan both call it, so the two note the same facts.

    Returns its raw-text key (a formula compiled by key), the reason Excel refuses it in a file
    (``refusals`` holds one per key; such a formula is left out of the rule, since Excel refuses
    the file), and an array formula's range (0-based).
    """
    raw: str | None = None
    reason: object = None
    box: tuple[int, int, int, int] | None = None
    if f.kind == "dataTable" or f.text is None:  # a shared formula's dependent too
        pass
    elif f.kind == "array" and f.ref and (box := _box(f.ref)) is not None:
        reason = refused_in_file(f.text)
        if reason is None:
            seen.note((f.text,), f.text, marked)
    else:
        # Whether Excel refuses a formula depends only on its text outside the references, so
        # it is decided once per raw-text key (§5.3), like the template.
        raw = raw_key(f.text, row, col)
        reason = refusals.get(raw, _UNSEEN)
        if reason is _UNSEEN:
            reason = refusals[raw] = refused_in_file(f.text)
        if reason is None:
            seen.note(raw, f.text, marked)
    if reason is None:
        seen.uncached |= uncached
        seen.ca |= marked
    return raw, reason, box


def _parses(text: str) -> bool:
    """Whether the parser reads a formula: its key's repeats then need no text of their own."""
    try:
        parse(text)
    except (FormulaSyntaxError, TokenizeError):
        return False
    return True


def _compile_sheet(state: _SheetLoad) -> None:
    """Make a recorded sheet's queued formulas the engine's, in the order they were read.

    The same steps, in the same order, as recording them as they were read: a formula the parser
    cannot read, or that Excel refuses in a file, keeps its cached value and is reported; a cell
    without a value fills the range of an array formula compiled before it; a shared formula's
    dependents follow, once every master is known.
    """
    engine, index, full, seen = state.engine, state.index, state.full, state.seen
    s = engine.sheets[index]
    values = s.values
    report = engine.report
    set_formula, template_for = engine.set_formula, engine.template_for
    masters: dict[int, int] = {}  # shared-formula index -> template id
    boxes: list[tuple[int, int, int, int, int]] = []  # ranges array formulas fill (§7.7)
    for key, entry in zip(state.cells, state.entries, strict=True):
        row, col = divmod(key, COLS)
        cached = values.get(key)
        if isinstance(entry, str) or (entry is not None and len(entry) == 1):
            marked = not isinstance(entry, str)
            common = entry if isinstance(entry, str) else entry[0]
            # a key queued before with its text, and compiled: found without the text
            tid = template_for("", row, col, common)
            set_formula(index, row, col, template=tid, cached=cached, stale=full)
            if marked:
                if seen.marked_stale:
                    seen.marked.append(node(index, row, col))
                if engine.iterate:
                    engine.calc_always.add(node(index, row, col))
            continue
        if entry is None:
            anchor = _filled_by(boxes, row, col)
            if anchor is not None:
                engine.set_member(index, row, col, anchor, cached)
            continue
        f, raw, dynamic, reason = entry
        where = CellRef(s.name, row + 1, col + 1)
        box = _box(f.ref) if f.kind in ("array", "dataTable") and f.ref else None
        if box is not None and not dynamic:  # cells Excel does not let one change one by one
            engine.locked.setdefault(index, []).append(box)
        if f.kind == "dataTable" or f.text is None:
            label = "array formula" if f.kind == "array" else "data table"
            _keep(engine, index, row, col, cached, label, where)
            continue
        if reason is not None:
            report.refused[where] = str(reason)
            s.set(row, col, BLANK if cached is None else cached)
            continue
        marked = f.attributes.get("ca") in ("1", "true")
        try:
            if f.kind == "array" and box is not None:
                tid = template_for(f.text, row, col, array_mode=True)
            else:
                tid = template_for(f.text, row, col, raw)
        except (FormulaSyntaxError, TokenizeError) as exc:
            report.refused[where] = str(exc)
            s.set(row, col, BLANK if cached is None else cached)
            continue
        if f.kind == "array" and box is not None:
            set_formula(
                index, row, col, template=tid, cached=cached, stale=full, array=(dynamic, box)
            )
            if marked and seen.marked_stale:
                seen.marked.append(node(index, row, col))
            boxes.append((*box, node(index, row, col)))
            continue
        if f.kind == "shared" and f.si is not None:
            masters[f.si] = tid
        set_formula(index, row, col, template=tid, cached=cached, stale=full)
        if marked and seen.marked_stale:
            seen.marked.append(node(index, row, col))
        if engine.iterate and marked:
            engine.calc_always.add(node(index, row, col))
    for key, si in zip(state.dependents, state.sis, strict=True):
        row, col = divmod(key, COLS)
        master = masters.get(si)
        if master is None:  # no master in the file: the cached value stays, with no formula
            continue
        set_formula(index, row, col, template=master, cached=values.get(key), stale=full)
        if key in state.marked_dependents and seen.marked_stale:
            seen.marked.append(node(index, row, col))
    state.cells, state.entries, state.ranges = array("q"), [], []
    state.dependents, state.sis, state.marked_dependents = array("q"), array("q"), set()
    state.keys = {}
    state.compiled = True


def _cell(ref: str) -> tuple[int, int]:
    """A cell reference such as ``B2`` as 1-based (row, col)."""
    col = 0
    for i, ch in enumerate(ref):
        o = ord(ch)
        if 65 <= o <= 90:
            col = col * 26 + o - 64
        elif 97 <= o <= 122:
            col = col * 26 + o - 96
        else:
            return int(ref[i:]), col
    raise ValueError(f"not a cell reference: {ref!r}")


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
