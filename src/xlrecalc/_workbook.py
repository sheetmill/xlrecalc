"""``load`` and ``Workbook``: read and change cells, recalculate, save (ARCHITECTURE.md §11.2)."""

from __future__ import annotations

import contextlib
import enum
import io
import math
import os
import re
from collections.abc import Callable, Iterator
from typing import BinaryIO, TypeAlias

import xlio
from xlio import CellEdit

from xlrecalc._engine import Engine, node
from xlrecalc._formula import FormulaSyntaxError, TokenizeError, parse, tokenize
from xlrecalc._formula._codegen import kind_at
from xlrecalc._formula._fastkey import translate
from xlrecalc._formula._fileform import display_form, file_form
from xlrecalc._formula._nodes import (
    ArrayConstant,
    Binary,
    Call,
    Group,
    Invoke,
    Leaf,
    Missing,
    Node,
    Postfix,
    Prefix,
    Reference,
)
from xlrecalc._formula._tokens import Kind, column_number
from xlrecalc._functions import REGISTRY, SPECS
from xlrecalc._functions._registry import _LIFTED
from xlrecalc._limits import ReadLimits, io_limits
from xlrecalc._load import RecalcOnLoad, _Loader, recalc_mode, refused_in_file
from xlrecalc._load import load as _load_engine
from xlrecalc._opening import Opening
from xlrecalc._quirks import Quirks
from xlrecalc._report import CellRef, Report
from xlrecalc._save import array_edits, file_value, lookups
from xlrecalc._store import COLS, MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Value

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
        self,
        engine: Engine,
        source: str | os.PathLike[str] | BinaryIO | None,
        limits: ReadLimits | None,
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
        return self._engine.settled_report()

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
        self._engine.edited(node(sheet, row, col))
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
            template = engine.templates[engine.template_for(text, row, col)]
            intersects = template.intersects or _may_be_array(
                text, lambda name: engine._name_text(name, sheet)
            )
            engine.set_formula(sheet, row, col, text, array=(True, (row, col, row, col)))
        except (FormulaSyntaxError, TokenizeError) as exc:
            raise FormulaError(f"{formula!r}: {exc}") from exc
        engine.entered.add(node(sheet, row, col))
        engine.edited(node(sheet, row, col))
        self._edits.setdefault(sheet, {})[row + 1, col + 1] = True
        self._marked[sheet, row, col] = (intersects, legacy)

    def _needs_marker(self, sheet: int, row: int, col: int) -> bool:
        """Whether a formula needs the dynamic-array marker in the file.

        A formula set here is a dynamic array; as Excel writes a typed formula (capture
        ``marker_typed``), it is marked when a legacy formula would mean something else: it
        spills, or takes an implicit intersection other than an explicit ``@`` (``=@A1:A3`` is
        written as the plain ``A1:A3``), or its text may give an array, whatever the size of the
        result (``SEQUENCE(1)``, ``-SEQUENCE(1)``, not ``SUM(SEQUENCE(3))``; ``_may_be_array``). A
        dynamic array read from the file stays one.
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

        With iterative calculation on, saving runs a loop once more, as Excel's saving does, once
        the loop has run (an evaluating load, an edit): the values in memory change with the
        file's. A trusting load with no edit runs nothing (ARCHITECTURE.md §7.2).

        Args:
            path: Where to write; by default the file the workbook was loaded from.
            full_calc_on_load: Ask Excel to recalculate when it opens the file; with
                ``False`` a request the file held is removed, as Excel saves a file.

        Raises:
            ValueError: No ``path`` for a workbook loaded from a file object, or a workbook read
                through a :class:`Loader`, which has no file to copy.
        """
        if self._source is None:
            raise ValueError(
                "a workbook read through a Loader has no file to copy: load or recalculate the "
                "file to save it"
            )
        if path is None:
            if not isinstance(self._source, (str, os.PathLike)):
                raise ValueError("a workbook loaded from a file object needs a path to save to")
            path = self._source
        engine = self._engine
        live = engine.loop_live()
        engine.recalculate()  # with a live loop, this is the save's run
        if not live:
            engine.save_run()  # the recalculation was the opening's: the save's run follows
        uncomputed = {
            (c.sheet.casefold(), c.row, c.col) for c in engine.settled_report().uncomputed
        }
        edits: dict[str, dict[tuple[int, int], CellEdit]] = {}
        for sheet, cells in self._edits.items():
            s = engine.sheets[sheet]
            out = edits.setdefault(s.name, {})
            for (row, col), is_formula in cells.items():
                value = s.values.get((row - 1) * COLS + col - 1, BLANK)
                text = engine.formula_text(sheet, row - 1, col - 1) if is_formula else None
                if text is not None and (s.name.casefold(), row, col) in uncomputed:
                    out[row, col] = CellEdit(text, None)
                else:
                    out[row, col] = CellEdit(text, file_value(value))
        filled_edits, fills = array_edits(engine)
        for name, filled in filled_edits.items():
            index = self._sheet(name)
            target = edits.setdefault(name, {})
            for (row, col), edit in filled.items():
                if edit.formula is not None and not self._needs_marker(index, row - 1, col - 1):
                    marked = self._marked.get((index, row - 1, col - 1))
                    text = edit.formula if marked is None else marked[1]
                    edit = CellEdit(text, edit.value)  # a plain formula means the same
                target[row, col] = edit
        source = self._source
        if not isinstance(source, (str, os.PathLike)):
            source.seek(0)
        xlio.patch(
            source,
            path,
            cached=lookups(engine),
            cells=edits,
            fills=fills,
            translate=translate,
            drop_calc_chain=self._formula_removed,
            full_calc_on_load=full_calc_on_load,
            limits=io_limits(self._limits),
        )


# Arguments that carry an array to the result although the registry does not lift them: N's
# value, and INDEX's array, whose shape the text does not give (capture marker_typed).
_CARRIED = {"N": (0,), "INDEX": (0,)}
# ROW and COLUMN give an array for a range: a call's result there may be one (capture
# lambda_calls: ROW(LAMBDA(x,x)(Z2)) is marked, ISREF(LAMBDA(x,x)(Z1)) is not).
_CARRIED |= {"ROW": (0,), "COLUMN": (0,)}


def _calls_a_function(node: Node) -> bool:
    """A call of a function value, or of a function that calls one (``MAP``, ``MAKEARRAY``...)."""
    while isinstance(node, Group):
        node = node.inner
    if isinstance(node, Invoke):
        return True
    if not isinstance(node, Call):
        return False
    spec = SPECS.get(
        ((node.prefix.text if node.prefix is not None else "") + node.name.text).upper()
    )
    return spec is not None and "function" in spec.kinds


def _may_be_array(text: str, name_text: Callable[[str], str | None] | None = None) -> bool:
    """Whether a formula (file form) may give an array, read from its text as Excel reads it.

    Capture ``marker_typed``: an array constant, or a call of a function that makes an array
    (``makes_array``), carried through operators, the arguments a function takes element by
    element (``IF``'s branches, ``ABS``, ``IFERROR``, ``OFFSET``'s offsets) and those of
    ``_CARRIED``, and through a defined name (``name_text`` gives a name's formula), but not
    through a function that takes an array (``SUM``, ``ROWS``) nor ``@``; ``INDEX`` with an
    index that is not a number other than 0, and ``OFFSET`` with a height or width that is not a
    number, since their size is then not in the text. ``LET`` carries its calculation, a name in
    it standing for what it is bound to (capture ``let_typed``: ``LET(x,SEQUENCE(3),SUM(x))`` is
    not marked). A call of a function value (``LAMBDA(x,x)(Z1:Z3)``, ``ADDONE(1)`` through a
    defined name) may give an array; an uncalled ``LAMBDA`` marks the formula only as its result,
    through ``LET`` and names, not inside an operation (captures ``lambda_calls``,
    ``lambda_names``: ``=LAMBDA(x,x)+1`` and ``IF(TRUE,LAMBDA(x,x))`` are not marked).
    """
    try:
        root = parse(text).root
    except (FormulaSyntaxError, TokenizeError):
        return False
    stack: list[Node] = [root]
    result: set[int] = {id(root)}  # nodes that are the formula's result, not an operand
    seen: set[str] = set()  # names already followed: a name may refer to itself
    bound: dict[str, Node] = {}  # LET names, by their folded text, to their values
    while stack:
        node = stack.pop()
        at_result = id(node) in result
        if isinstance(node, ArrayConstant | Invoke):
            return True
        if (
            isinstance(node, Call)
            and node.prefix is None
            and node.name.text.upper() == "_XLFN.LAMBDA"
        ):
            if at_result:
                return True
            continue
        if isinstance(node, Reference) and node.token.kind is Kind.NAME and node.prefix is None:
            folded = node.token.text.casefold()
            if folded in bound:
                if folded not in seen:
                    seen.add(folded)
                    stack.append(bound[folded])
                    if at_result:
                        result.add(id(bound[folded]))
                continue
        if isinstance(node, Call) and node.prefix is None and node.name.text.upper() == "_XLFN.LET":
            args = node.args
            pairs = args if len(args) % 2 == 0 else args[:-1]
            for declared, value in zip(pairs[::2], pairs[1::2], strict=True):
                if isinstance(declared, Reference):
                    bound[declared.token.text.casefold()] = value
            if args:
                stack.append(pairs[-2] if len(args) % 2 == 0 and len(pairs) >= 2 else args[-1])
                if at_result:
                    result.add(id(stack[-1]))
            continue
        if isinstance(node, Reference) and node.token.kind is Kind.NAME and name_text is not None:
            folded = node.token.text.casefold()
            definition = name_text(node.token.text) if folded not in seen else None
            seen.add(folded)
            if definition is not None:
                with contextlib.suppress(FormulaSyntaxError, TokenizeError):
                    stack.append(parse(definition.removeprefix("=")).root)
                    if at_result:
                        result.add(id(stack[-1]))
        elif isinstance(node, Group):
            stack.append(node.inner)
            if at_result:
                result.add(id(node.inner))
        elif isinstance(node, Prefix | Postfix):
            stack.append(node.operand)
        elif isinstance(node, Binary):
            stack += (node.left, node.right)
        elif isinstance(node, Call):
            name = ((node.prefix.text if node.prefix is not None else "") + node.name.text).upper()
            spec = SPECS.get(name)
            if spec is None and name_text is not None and name_text(node.name.text) is not None:
                return True  # a call of a defined name, which holds a LAMBDA if it runs
            if spec is None or name == "_XLFN.SINGLE":
                continue
            cycle: int = getattr(REGISTRY[name], "cycle", 1)
            if spec.makes_array:
                return True
            if name == "INDEX" and any(not _nonzero_number(a) for a in node.args[1:3]):
                return True
            if name in ("ROWS", "COLUMNS") and any(_calls_a_function(a) for a in node.args):
                return True  # capture lambda_more: ROWS(MAKEARRAY(...)); not ROWS(SEQUENCE(1))
            if name == "OFFSET" and any(
                not isinstance(a, Missing) and not _nonzero_number(a) for a in node.args[3:5]
            ):
                return True
            carried = _CARRIED.get(name, ())
            stack += (
                a
                for i, a in enumerate(node.args)
                if i in carried or kind_at(spec.kinds, i, cycle) in _LIFTED
            )
    return False


def _nonzero_number(node: Node) -> bool:
    return isinstance(node, Leaf) and node.token.kind is Kind.NUMBER and float(node.token.text) != 0


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
    recalc_on_load: RecalcOnLoad = "excel",
    trust_cached: bool | None = None,
) -> Workbook:
    """Read an xlsx workbook.

    Nothing is evaluated yet: a value is evaluated when it is read. Cached values the file
    holds are trusted (ARCHITECTURE.md §7.4) where Excel trusts them on opening the file, by
    default: every formula is evaluated when the file has no ``calcPr``, asks for a full
    calculation on load (``fullCalcOnLoad``), was last calculated by an older Excel
    (``calcId``), was not saved by Excel (no ``<fileVersion appName="xl">``) or calls a volatile
    function in a cell without Excel's ``ca="1"`` mark; otherwise the cells so marked (volatile
    cells and their dependents) are evaluated; in manual calculation every cached value is
    trusted (``RecalcCheck`` gives the same answer without building an engine).

    Args:
        source: The file, as a path or a binary file object (read into memory).
        quirks: Which Excel quirks to reproduce; all by default.
        limits: Limits on what reading and computing the workbook may cost (``ReadLimits``):
            the package's parts and sizes, the cells an array may hold, the evaluations of one
            run of iterative calculation, and the steps of loops inside functions.
        recalc_on_load: ``"excel"``, Excel's rule above; ``"always"``, evaluate every formula;
            ``"never"``, trust every cached value the file holds but the volatile formulas'
            (``RAND``, ``NOW``..., through a defined name too). A formula with no cached value
            is evaluated in every mode.
        trust_cached: Deprecated, to be removed in 1.0: ``False`` is ``recalc_on_load="always"``,
            ``True`` the default.

    Raises:
        PackageError: The file is not a readable xlsx package.
        LimitError: The file exceeds ``limits``.
    """
    if not isinstance(source, (str, os.PathLike)):
        source = io.BytesIO(source.read())
    mode = recalc_mode(recalc_on_load, "trust_cached", trust_cached, stacklevel=2)
    engine = _load_engine(source, quirks=quirks, limits=limits, recalc_on_load=mode)
    return Workbook(engine, source, limits)


class Loader:
    """Read a workbook whose cells the caller already streams through ``xlio``, reading them once.

    A caller that opens a workbook with ``xlio`` and reads every sheet's cells for its own use would
    read the file a second time with :func:`load`. A ``Loader`` takes the open workbook instead:
    each sheet's cells pass through :meth:`cells`, which records every cell and hands it on, so the
    caller keeps its one loop over the cells, as with ``csv.reader`` over a file the caller owns.
    :meth:`workbook` then gives the :class:`Workbook` that :func:`load` would give, by the same
    code: ``load`` is a ``Loader`` with every sheet fed in order.

    Sheets are fed in any order, and any of them need not be fed: :meth:`workbook` reads any sheet
    not fed, and any sheet whose loop the caller left part-way, again from its start, recording the
    cells not yet recorded. A sheet the loader has read for itself can still be fed: its cells are
    handed on from ``xlio`` again, unrecorded, and a loop the caller is still running when the
    loader reads its sheet hands on the rest of its cells unrecorded. The result is the same
    whatever the order; a sheet read twice costs only the time of that second reading.

    Example::

        with xlio.open(path) as book:
            loader = xlrecalc.Loader(book)
            for sheet in book.sheets:
                for cell in loader.cells(sheet):
                    ...  # the caller's own use of the cell
            workbook = loader.finish()

    Args:
        book: An open ``xlio`` workbook. The caller opens and closes it; it stays open until the
            loader has read what it needs, which :meth:`finish` guarantees. Its archive limits
            are the ones it was opened with.
        recalc_on_load: As for :func:`load`: ``"excel"``, Excel's rule on opening;
            ``"always"``; ``"never"``.
        limits: The computing limits (``max_array_cells``, ``max_iteration_evaluations``,
            ``max_function_steps``); the archive limits are the book's.

    The workbook :meth:`workbook` returns is computed and edited as any other, but has no file
    to copy, so :meth:`Workbook.save` raises; :func:`recalculate` writes a file's values.
    """

    def __init__(
        self,
        book: xlio.Workbook,
        *,
        recalc_on_load: RecalcOnLoad = "excel",
        limits: ReadLimits | None = None,
    ) -> None:
        self._loader = _Loader(book, recalc_on_load=recalc_on_load, limits=limits)
        self._limits = limits
        self._workbook: Workbook | None = None

    def cells(self, sheet: xlio.Sheet, *, rich: bool = False) -> Iterator[xlio.Cell]:
        """Each cell of a sheet, as ``sheet.cells(rich=rich)`` gives it, recorded as it passes.

        Valid for each sheet once, in any order, at any time: also after :meth:`workbook` or
        :meth:`finish`. A sheet the loader has already read is handed on unrecorded.

        Args:
            sheet: One of the book's sheets; a chart sheet passes no cells.
            rich: As for ``xlio``'s ``Sheet.cells``: give string cells their runs and phonetic
                guides (``Cell.rich``). What is recorded is the same either way, and a sheet the
                loader reads for itself is read without them.

        Raises:
            ValueError: The sheet is not the book's, or was fed already.
        """
        return self._loader.cells(sheet, rich=rich)

    @property
    def opening(self) -> Opening | None:
        """What the load computes on opening, ``"all"``, ``"some"`` or ``"none"`` formulas.

        :class:`RecalcCheck`'s answer, from the same rule and the same facts about the formulas,
        so a caller using a ``Loader`` needs no ``RecalcCheck``. ``None`` while the answer
        depends on cells not recorded yet; ``"all"`` may come early, from the workbook's own
        parts or from the first formula that settles it. A formula Excel refuses in a file is
        left out (Excel refuses the file); ``RecalcCheck`` cannot tell one.
        """
        return self._loader.opening

    def workbook(self) -> Workbook:
        """The workbook: every sheet not fed, or left part-way, is read now; the same each call.

        Raises:
            ValueError: A sheet must be read, and the book is closed.
        """
        if self._workbook is None:
            self._workbook = Workbook(self._loader.settle(), None, self._limits)
        return self._workbook

    def finish(self) -> Workbook:
        """:meth:`workbook`, at the end of the caller's reading: the same object.

        Raises:
            ValueError: :meth:`finish` was called already, or a sheet must be read and the
                book is closed.
        """
        self._loader.finish()
        return self.workbook()
