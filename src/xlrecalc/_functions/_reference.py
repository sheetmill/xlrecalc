"""References made during evaluation: OFFSET, INDIRECT, CELL (ARCHITECTURE.md §6.3).

All three are volatile, as in Excel: evaluated again after any change. ``OFFSET`` and
``INDIRECT`` return references the engine checks as they are read (``Engine.dynamic_ref``), so a
formula reading a cell not evaluated yet waits for it, and one reading itself is circular.
Behavior from capture ``dynrefs``.
"""

from __future__ import annotations

import math
import posixpath
import re
from collections.abc import Callable
from typing import Any

from xlrecalc._arrays import check_area
from xlrecalc._coerce import Context, to_number, to_text
from xlrecalc._formula import MISSING, Kind, TokenizeError, tokenize
from xlrecalc._formula._codegen import sheet_text
from xlrecalc._formula._fileform import display_form
from xlrecalc._formula._template import cell_axes, line_axes
from xlrecalc._functions._registry import function
from xlrecalc._report import column_letters, sheet_prefix
from xlrecalc._store import MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Ref

REF, VALUE = ErrorValue.REF, ErrorValue.VALUE
L = "lookup and reference"
VALUE = ErrorValue.VALUE
INFO = "information"
_FORMULATEXT = "the @ the formula bar adds to a legacy formula is not shown"


def _size(arg: Any, default: int, ctx: Context) -> int | ErrorValue:
    if arg is MISSING or arg is None:
        return default
    n = to_number(arg, ctx)
    return n if isinstance(n, ErrorValue) else math.trunc(n)


def _span(start: int, size: int) -> tuple[int, int]:
    """A size from a start: a negative size extends up or left (``OFFSET(Z5,0,0,-2)`` is Z4:Z5)."""
    return (start, start + size - 1) if size > 0 else (start + size + 1, start)


@function(
    "OFFSET",
    "raw",
    "number",
    "number",
    "raw",
    "raw",
    min_args=3,
    volatile=True,
    reference=True,
    category=L,
)
def OFFSET(
    ctx: Any, ref: Any, rows: float, cols: float, height: Any = MISSING, width: Any = MISSING
) -> Any:
    """The reference ``rows`` down and ``cols`` right of ``ref``, ``height`` by ``width``.

    The offsets are truncated; an omitted size is the reference's; a size of 0, or a result off
    the sheet, is ``#REF!``.
    """
    if isinstance(ref, ErrorValue):
        return ref
    if not isinstance(ref, Ref):
        return VALUE
    h = _size(height, ref.rows, ctx)
    if isinstance(h, ErrorValue):
        return h
    w = _size(width, ref.columns, ctx)
    if isinstance(w, ErrorValue):
        return w
    if h == 0 or w == 0:
        return REF
    r1, r2 = _span(ref.r1 + math.trunc(rows), h)
    c1, c2 = _span(ref.c1 + math.trunc(cols), w)
    if r1 < 0 or c1 < 0 or r2 > MAX_ROW or c2 > MAX_COL:
        return REF
    return ctx.dynamic_ref(ref.sheet, r1, c1, r2, c2)


_R1C1_PART = r"R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?"
_R1C1 = re.compile(rf"(?:('(?:[^']|'')+'|[^'!\[\]]+)!)?{_R1C1_PART}(?::{_R1C1_PART})?", re.I)


def _r1c1_axis(part: str | None, base: int) -> int:
    if not part:
        return base
    if part.startswith("["):
        return base + int(part[1:-1])
    return int(part) - 1


def _parse(text: str, a1: bool, row: int, col: int) -> tuple[str | None, tuple[int, ...]] | str:
    """A reference's text as (sheet name or None, (r1, c1, r2, c2)), or a name to look up."""
    if not a1:
        m = _R1C1.fullmatch(text)
        if m is None:
            return ""
        sheet = m[1]
        if sheet is not None and sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        r1, c1 = _r1c1_axis(m[2], row), _r1c1_axis(m[3], col)
        r2, c2 = (_r1c1_axis(m[4], row), _r1c1_axis(m[5], col)) if m[0].count(":") else (r1, c1)
        return sheet, (r1, c1, r2, c2)
    try:
        tokens = [t for t in tokenize(text) if t.kind is not Kind.END]
    except TokenizeError:
        return ""
    sheet = None
    if tokens and tokens[0].kind is Kind.PREFIX:
        if "[" in tokens[0].text or ":" in tokens[0].text[:-1].strip("'"):
            return ""  # another workbook, or several sheets
        sheet = sheet_text(tokens[0])
        tokens = tokens[1:]
    kinds = [t.kind for t in tokens]
    if kinds == [Kind.CELL]:
        r, c = cell_axes(tokens[0].text, 0, 0)
        return sheet, (r.value, c.value, r.value, c.value)
    if kinds == [Kind.CELL, Kind.OPERATOR, Kind.CELL] and tokens[1].text == ":":
        top, left = cell_axes(tokens[0].text, 0, 0)
        bottom, right = cell_axes(tokens[2].text, 0, 0)
        return sheet, (top.value, left.value, bottom.value, right.value)
    if kinds in ([Kind.COLUMNS], [Kind.ROWS]):
        first, second = line_axes(tokens[0], 0, 0)
        if kinds == [Kind.COLUMNS]:
            return sheet, (0, first.value, MAX_ROW, second.value)
        return sheet, (first.value, 0, second.value, MAX_COL)
    if kinds == [Kind.NAME] and sheet is None:
        return tokens[0].text
    return ""


@function(
    "INDIRECT",
    "text",
    "logical",
    min_args=1,
    volatile=True,
    reference=True,
    category=L,
    makes_array=True,
)
def INDIRECT(ctx: Any, text: str, a1: bool = True) -> Any:
    """The reference a text names: A1 (``"Sheet2!A1"``), R1C1 with ``a1`` false, or a name.

    A leading space, another workbook, or text that is no reference gives ``#REF!``.
    """
    if not text or text[0].isspace():
        return REF
    parsed = _parse(text.rstrip(" "), a1, ctx.row, ctx.col)
    if isinstance(parsed, str):
        if not parsed:
            return REF
        definition = ctx.name_text(parsed)  # a name defined as a reference: "S!$Z$2"
        if definition is None:
            return REF
        parsed = _parse(definition.lstrip("="), True, 0, 0)
        if isinstance(parsed, str):
            return REF
    sheet_name, (r1, c1, r2, c2) = parsed
    sheet = ctx.sheet if sheet_name is None else ctx.sheet_index(sheet_name)
    if sheet is None:
        return REF
    r1, r2 = min(r1, r2), max(r1, r2)
    c1, c2 = min(c1, c2), max(c1, c2)
    if r1 < 0 or c1 < 0 or r2 > MAX_ROW or c2 > MAX_COL:
        return REF
    return ctx.dynamic_ref(sheet, r1, c1, r2, c2)


def _book_prefix(ctx: Any, sheet: int) -> str:
    """A sheet as ``CELL`` names it: ``[book.xlsx]Sheet2!``, quoted when the sheet needs it."""
    name = ctx.sheets[sheet].name
    if ctx.path is None:
        return sheet_prefix(name)
    book = f"[{posixpath.basename(ctx.path)}]{name}"
    # After the book name only characters call for quotes: [b.xlsx]S1 (capture cell_noref),
    # '[b.xlsx]My Sheet' (capture dynrefs).
    if re.search(r"[^\w.]", name):
        return "'" + book.replace("'", "''") + "'!"
    return book + "!"


@function(
    "CELL",
    "text",
    "raw",
    min_args=1,
    volatile=True,
    status="partial",
    note="the kinds that depend on formatting are unsupported: styles are not read",
    category=INFO,
)
def CELL(ctx: Any, info: str, ref: Any = MISSING) -> Any:
    """Information about a reference's top-left cell.

    The kinds ``row``, ``col``, ``address``, ``contents``, ``type`` and ``filename``; without
    a reference, the active cell of the sheet the workbook opens on, from any sheet, as Excel
    takes it when it recalculates a file it opens (capture ``cell_noref``). The kinds that
    depend on formatting (``format``, ``width``, ``prefix``, ``protect``, ``color``,
    ``parentheses``) are not implemented; an unknown kind is ``#VALUE!``.
    """
    kind = info.lower()
    if ref is MISSING or ref is None:  # the active sheet's active cell (capture cell_noref)
        sheet, row, col = ctx.active
        ref = Ref(ctx.cells, sheet, row, col, row, col)
    if isinstance(ref, ErrorValue):
        return ref
    if not isinstance(ref, Ref):
        return VALUE
    if kind == "row":
        return float(ref.r1 + 1)
    if kind == "col":
        return float(ref.c1 + 1)
    if kind == "address":
        address = f"${column_letters(ref.c1 + 1)}${ref.r1 + 1}"
        return address if ref.sheet == ctx.sheet else _book_prefix(ctx, ref.sheet) + address
    if kind in ("contents", "type"):
        v = ref.value(0, 0)
        if kind == "contents":
            return 0.0 if v is BLANK else v
        return "b" if v is BLANK else "l" if isinstance(v, str) else "v"
    if kind == "filename":
        if ctx.path is None:
            return ""
        folder, book = posixpath.split(ctx.path)
        return f"{folder}/[{book}]{ctx.sheets[ref.sheet].name}"
    if kind in ("format", "protect", "width", "prefix", "color", "parentheses"):
        return ctx.unsupported(f'CELL("{kind}")')
    return VALUE


@function("TRANSPOSE", "matrix", category=L, makes_array=True)
def TRANSPOSE(ctx: Any, a: Any) -> Any:
    """Rows as columns. A range in a legacy formula is intersected first, as in Excel."""
    if not isinstance(a, tuple):
        return a
    return tuple(zip(*a, strict=True))


_ADDRESS_PLAIN = re.compile(r"(?:[^\W\d]|\[)[\w.\[\]]*")
_CELL_LIKE = re.compile(r"(?i)[A-Z]{1,3}\d+|R\d*C\d*|R|C|TRUE|FALSE")


def _address_sheet(name: str) -> str:
    """A sheet as ADDRESS names it: quoted unless plain (``[b]x`` and ``Résumé`` are plain)."""
    if name == "" or (_ADDRESS_PLAIN.fullmatch(name) and not _CELL_LIKE.fullmatch(name)):
        return name + "!"
    return "'" + name.replace("'", "''") + "'!"


@function("ADDRESS", "number", "number", "number", "logical", "text", min_args=2, category=L)
def ADDRESS(
    ctx: Any, row: float, col: float, kind: float = 1.0, a1: bool = True, sheet: Any = MISSING
) -> Any:
    """A cell's address as text: absolute or relative parts, A1 or R1C1 style, a sheet."""
    r, c, k = math.trunc(row), math.trunc(col), math.trunc(kind)
    if not 1 <= k <= 4:
        return VALUE
    abs_row, abs_col = k in (1, 2), k in (1, 3)
    if a1 or abs_row:
        if not 1 <= r <= MAX_ROW + 1:
            return VALUE
    elif not 0 <= r <= MAX_ROW:  # a relative R1C1 row is an offset: 0 is "R"
        return VALUE
    if a1 or abs_col:
        if not 1 <= c <= MAX_COL + 1:
            return VALUE
    elif not 0 <= c <= MAX_COL:
        return VALUE
    if a1:
        text = ("$" if abs_col else "") + column_letters(c) + ("$" if abs_row else "") + str(r)
    else:
        text = (f"R{r}" if abs_row else (f"R[{r}]" if r else "R")) + (
            f"C{c}" if abs_col else (f"C[{c}]" if c else "C")
        )
    if sheet is MISSING:
        return text
    return _address_sheet(sheet) + text


def _each_cell(ref: Any, ctx: Any, read: Callable[[int, int, int], Any]) -> Any:
    """``read`` of each cell of a reference, as an array (capture ``array_refs``)."""
    if not isinstance(ref, Ref):
        return ref if isinstance(ref, ErrorValue) else VALUE
    if ref.is_cell():
        return read(ref.sheet, ref.r1, ref.c1)
    check_area(ref.rows, ref.columns, getattr(ctx, "max_array_cells", None))
    return tuple(
        tuple(read(ref.sheet, r, c) for c in range(ref.c1, ref.c2 + 1))
        for r in range(ref.r1, ref.r2 + 1)
    )


def _formula_shown(ctx: Any, sheet: int, row: int, col: int) -> Any:
    text = ctx.formula_text(sheet, row, col)
    return ErrorValue.NA if text is None else "=" + display_form(text)


def _isformula_array(ctx: Any, ref: Any) -> Any:
    return _each_cell(ref, ctx, lambda s, r, c: ctx.formula_text(s, r, c) is not None)


def _formulatext_array(ctx: Any, ref: Any) -> Any:
    return _each_cell(ref, ctx, lambda s, r, c: _formula_shown(ctx, s, r, c))


@function("_xlfn.ISFORMULA", "raw", errors=True, category=INFO, in_array=_isformula_array)
def ISFORMULA(ctx: Any, ref: Any) -> Any:
    """Whether a reference's top-left cell holds a formula."""
    if not isinstance(ref, Ref):
        return ref if isinstance(ref, ErrorValue) else VALUE
    return ctx.formula_text(ref.sheet, ref.r1, ref.c1) is not None


@function(
    "_xlfn.FORMULATEXT",
    "raw",
    errors=True,
    category=INFO,
    status="partial",
    note=_FORMULATEXT,
    in_array=_formulatext_array,
)
def FORMULATEXT(ctx: Any, ref: Any) -> Any:
    """A reference's top-left cell's formula as the formula bar shows it; ``#N/A`` without one."""
    if not isinstance(ref, Ref):
        return ref if isinstance(ref, ErrorValue) else VALUE
    return _formula_shown(ctx, ref.sheet, ref.r1, ref.c1)


@function("_xlfn.SHEET", "raw", min_args=0, errors=True, category=INFO)
def SHEET(ctx: Any, v: Any = MISSING) -> Any:
    """The number of a reference's sheet, or of a sheet named by a text."""
    if v is MISSING:
        return float(ctx.sheet + 1)
    if isinstance(v, Ref):
        return float(v.sheet + 1)
    if isinstance(v, ErrorValue):
        return v
    if isinstance(v, str):
        for i, sheet in enumerate(ctx.sheets):
            if sheet.name.casefold() == v.casefold():
                return float(i + 1)
    return ErrorValue.NA


@function("_xlfn.SHEETS", "raw", min_args=0, errors=True, category=INFO)
def SHEETS(ctx: Any, v: Any = MISSING) -> Any:
    """The number of sheets in the workbook, or in a reference (one here: no 3-D references)."""
    if v is MISSING:
        return float(len(ctx.sheets))
    if isinstance(v, Ref):
        return 1.0
    return v if isinstance(v, ErrorValue) else ErrorValue.NA


@function("AREAS", "raw", errors=True, category=L)
def AREAS(ctx: Any, ref: Any) -> Any:
    """The number of areas in a reference; a union is counted where the formula is compiled."""
    if isinstance(ref, Ref):
        return 1.0
    return ref if isinstance(ref, ErrorValue) else VALUE


@function("HYPERLINK", "value", "value", min_args=1, errors=True, category=L)
def HYPERLINK(ctx: Any, link: Any, name: Any = MISSING) -> Any:
    """The text a link shows: its name, or the link itself. Nothing is followed."""
    if isinstance(link, ErrorValue):
        return link
    if name is MISSING:
        return to_text(link, ctx)
    return 0.0 if name is BLANK else name
