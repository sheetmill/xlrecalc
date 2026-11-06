"""The surgical writer: a new package from a source package and new cached values (§9.2).

Every part that does not change is copied byte for byte, compressed data included. A worksheet
whose formula cells get new values is rewritten by a streaming scan of its bytes: expat reports
where each ``<c>`` starts and ends, and only the ``t`` attribute and the ``<v>`` element of a
cell whose value changed are replaced; everything else, the cell's ``<f>`` included, is copied
as written, so namespace prefixes and extensions survive. Whether a value changed is decided
while streaming, against the value the file holds, so no change set is kept in memory. The
output is written to a temporary file next to the target and moved into place.

Values come in file terms, as the reader gives them: ``("n", float)``, ``("str", str)``,
``("b", bool)`` or ``("e", code)``.
"""

from __future__ import annotations

import heapq
import itertools
import os
import posixpath
import re
import shutil
import tempfile
import time
import xml.parsers.expat
import zipfile
from collections.abc import Callable, Mapping, Sequence
from contextlib import ExitStack
from pathlib import Path
from typing import BinaryIO, NamedTuple

from xlrecalc._xlsx._book import Book
from xlrecalc._xlsx._cells import _column, unescape
from xlrecalc._xlsx._package import PackageError, ReadLimits
from xlrecalc._xlsx._zip import ZipWriter

__all__ = [
    "Edit",
    "FileValue",
    "Lookup",
    "Translate",
    "number_text",
    "patch_cells",
    "text_escape",
    "write",
]

FileValue = tuple[str, float | str | bool]
Lookup = Callable[[int, int], FileValue | None]
"""A sheet's new values: ``lookup(row, col)`` (1-based) for a formula cell, ``None`` to keep it."""

_CHUNK = 1 << 16
_FILE_ERRORS = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}
)
_SPACE = " \t\n"  # XML whitespace that can stand in a written text (CR is escaped)
_ZIP64_ABOVE = 2 << 30  # a rewritten part larger than this in the source reserves Zip64 room
_T_ATTR = re.compile(rb"""\s+t\s*=\s*(?:"[^"]*"|'[^']*')""")
_ESCAPED = re.compile(r"_(?=x[0-9A-Fa-f]{4}_)|[\x00-\x08\x0b-\x1f\ud800-\udfff￾￿]")
_CALC = re.compile(rb"<(?P<p>[\w.-]+:)?calcPr\b(?P<a>[^>]*?)(?P<e>/?)>")
_FULL = re.compile(rb"""\bfullCalcOnLoad\s*=\s*(?:"([^"]*)"|'([^']*)')""")
_ROOT = re.compile(rb"<(?P<p>[\w.-]+:)?workbook\b")
# Elements that follow calcPr in a workbook part (ECMA-376 CT_Workbook), and its end tag.
_AFTER_CALC = (
    b"oleSize",
    b"customWorkbookViews",
    b"pivotCaches",
    b"smartTagPr",
    b"smartTagTypes",
    b"webPublishing",
    b"fileRecoveryPr",
    b"webPublishObjects",
    b"extLst",
)


def number_text(value: float) -> str:
    """A number as a file stores it: the shortest exact text, integral values without ``.0``."""
    if value.is_integer() and abs(value) < 2**53:
        return str(int(value))
    return repr(value).replace("e", "E")


def _escape_char(m: re.Match[str]) -> str:
    ch = m[0]
    return "_x005F_" if ch == "_" else f"_x{ord(ch):04X}_"


def text_escape(text: str) -> str:
    """Text as an element's content: XML escapes, and OOXML's ``_xHHHH_`` for what XML cannot hold.

    A carriage return is escaped too (XML would read it as a line feed), and a literal
    ``_xHHHH_`` gets its underscore escaped, so it reads back as written.
    """
    text = _ESCAPED.sub(_escape_char, text)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _same(old_type: str, old: str | None, new: FileValue) -> bool:
    kind, value = new
    if old is None:
        return False
    if kind == "n":
        if old_type != "n" or not old:
            return False
        try:
            return float(old) == value
        except ValueError:
            return False
    if kind == "str":
        return old_type == "str" and unescape(old) == value
    if kind == "b":
        return old_type == "b" and (old in ("1", "true")) == value
    return old_type == kind and old == value


def _value_element(prefix: bytes, new: FileValue) -> bytes:
    kind, value = new
    if kind == "n":
        assert isinstance(value, float)
        text = number_text(value)
    elif kind == "b":
        text = "1" if value else "0"
    elif kind == "str":
        assert isinstance(value, str)
        text = text_escape(value)
    else:
        # Excel caches the newer errors (#SPILL!, #CALC!, ...) as #VALUE!, the real one in a
        # rich value (capture arrays); a written file holds #VALUE!, as other readers see it.
        text = str(value) if value in _FILE_ERRORS else "#VALUE!"
    # Excel drops leading and trailing whitespace of a <v> not marked to preserve it.
    space = b""
    if kind == "str" and text and (text[0] in _SPACE or text[-1] in _SPACE):
        space = b' xml:space="preserve"'
    return b"<" + prefix + b"v" + space + b">" + text.encode("utf-8") + b"</" + prefix + b"v>"


def _check_encoding(head: bytes, name: str) -> None:
    if head.startswith((b"\xff\xfe", b"\xfe\xff")):
        raise PackageError(f"{name}: UTF-16 parts are not written")
    m = re.match(rb"(?:\xef\xbb\xbf)?<\?xml[^>]*encoding\s*=\s*[\"']([^\"']+)", head)
    if m and m[1].lower().replace(b"_", b"-") not in (b"utf-8", b"utf8"):
        raise PackageError(f"{name}: {m[1].decode('ascii', 'replace')} parts are not written")


class Edit(NamedTuple):
    """A cell changed through the API, in file terms.

    ``formula`` is the formula's file-form text, or ``None`` for a constant; ``value`` is the
    constant, or the formula's cached value (``None``: a blank cell, or no cached value).
    ``array_ref`` makes the formula an array formula over that range (``"B1:B5"``), a
    dynamic array with ``dynamic`` (marked with the dynamic-array cell metadata).
    """

    formula: str | None
    value: FileValue | None
    array_ref: str | None = None
    dynamic: bool = False


Translate = Callable[[str, int, int, int, int], str]
"""``translate(text, row, col, to_row, to_col)``: a formula moved between 0-based cells."""

_DROP_ATTRS = re.compile(rb"""\s+(?:t|cm|vm|spans)\s*=\s*(?:"[^"]*"|'[^']*')""")


def _letters(col: int) -> str:
    out = ""
    while col:
        col, rest = divmod(col - 1, 26)
        out = chr(65 + rest) + out
    return out


def _edit_xml(
    prefix: bytes, row: int, col: int, edit: Edit, attrs: bytes = b"", cm: str | None = None
) -> bytes:
    """A ``<c>`` element for an edited cell; ``attrs`` are the original ones to keep (``s``).

    ``cm`` is the cell-metadata index that marks a dynamic array in this workbook.
    """
    p = prefix
    ref = f"{_letters(col)}{row}".encode()
    if not attrs:
        attrs = b' r="' + ref + b'"'
    value = edit.value
    if edit.formula is not None:
        f_attrs = b""
        if edit.array_ref is not None:
            f_attrs = b' t="array" ref="' + edit.array_ref.encode() + b'"'
            if edit.dynamic and cm is not None:
                attrs += b' cm="' + cm.encode() + b'"'
        body = text_escape(edit.formula).encode("utf-8")
        f = b"<" + p + b"f" + f_attrs + b">" + body + b"</" + p + b"f>"
        if value is None:
            return b"<" + p + b"c" + attrs + b">" + f + b"</" + p + b"c>"
        t = b"" if value[0] == "n" else b' t="' + value[0].encode() + b'"'
        v = _value_element(p, value)
        return b"<" + p + b"c" + attrs + t + b">" + f + v + b"</" + p + b"c>"
    if value is None:
        return b"<" + p + b"c" + attrs + b"/>" if b" s=" in attrs else b""
    kind, data = value
    if kind == "str":
        assert isinstance(data, str)
        text = text_escape(data)
        outer = bool(text) and (text[0] in _SPACE or text[-1] in _SPACE)
        space = b' xml:space="preserve"' if outer else b""
        inner = b"<" + p + b"is><" + p + b"t" + space + b">" + text.encode("utf-8")
        inner += b"</" + p + b"t></" + p + b"is>"
        return b"<" + p + b"c" + attrs + b' t="inlineStr">' + inner + b"</" + p + b"c>"
    t = b"" if kind == "n" else b' t="' + kind.encode() + b'"'
    return b"<" + p + b"c" + attrs + t + b">" + _value_element(p, value) + b"</" + p + b"c>"


class Fill(NamedTuple):
    """A range of constant cells to write, as an array formula fills them (1-based, inclusive).

    ``value_at(row, col)`` gives each cell's value when it is written, so the range's cells are
    never held at once; an edit of the same cell takes precedence.
    """

    r1: int
    c1: int
    r2: int
    c2: int
    value_at: Callable[[int, int], FileValue | None]


class _Pending:
    """The cells still to write into a sheet, row by row and column by column: edits and fills."""

    def __init__(self, edits: Mapping[tuple[int, int], Edit], fills: Sequence[Fill]) -> None:
        self.edits = dict(edits)
        self.columns: dict[int, list[int]] = {}
        for r, c in sorted(self.edits):
            self.columns.setdefault(r, []).append(c)
        self.fills = list(fills)
        self._next = [(0, 0)] * len(self.fills)  # per fill: (row, next column not written)
        merged = heapq.merge(self.columns, *(range(f.r1, f.r2 + 1) for f in self.fills))
        self._rows = (r for r, _ in itertools.groupby(merged))
        self._row: int | None = next(self._rows, None)

    def next_row(self) -> int | None:
        """The first row not yet reached that has cells to write."""
        return self._row

    def reach_row(self) -> int | None:
        """Pass that row (its cells are written as the row is copied or inserted)."""
        row, self._row = self._row, next(self._rows, None)
        return row

    def _cursor(self, k: int, row: int) -> int:
        at, column = self._next[k]
        return column if at == row else self.fills[k].c1

    def has(self, row: int) -> bool:
        """Whether row ``row`` has cells left to write."""
        if self.columns.get(row):
            return True
        return any(
            f.r1 <= row <= f.r2 and self._cursor(k, row) <= f.c2 for k, f in enumerate(self.fills)
        )

    def take(self, row: int, below: int | None) -> list[tuple[int, Edit]]:
        """The cells of ``row`` left to write with a column below ``below`` (all for ``None``)."""
        limit = 1 << 30 if below is None else below
        out: dict[int, Edit] = {}
        columns = self.columns.get(row)
        while columns and columns[0] < limit:
            c = columns.pop(0)
            edit = self.edits.pop((row, c), None)
            if edit is not None:
                out[c] = edit
        for k, f in enumerate(self.fills):
            if not f.r1 <= row <= f.r2:
                continue
            start = self._cursor(k, row)
            stop = min(f.c2, limit - 1)
            for c in range(start, stop + 1):
                if c not in out and (row, c) not in self.edits:
                    out[c] = Edit(None, f.value_at(row, c))
            self._next[k] = (row, max(start, stop + 1))
        return sorted(out.items())

    def pop(self, row: int, col: int) -> Edit | None:
        """The edit for a cell the sheet holds, if any, now written."""
        edit = self.edits.pop((row, col), None)
        if edit is not None:
            columns = self.columns.get(row)
            if columns and col in columns:
                columns.remove(col)
        for k, f in enumerate(self.fills):
            if f.r1 <= row <= f.r2 and f.c1 <= col <= f.c2:
                if edit is None and self._cursor(k, row) <= col:
                    edit = Edit(None, f.value_at(row, col))
                self._next[k] = (row, max(self._cursor(k, row), col + 1))
        return edit

    def __bool__(self) -> bool:
        return bool(self.edits) or self._row is not None


def patch_cells(
    source: BinaryIO,
    write: Callable[[bytes], object],
    lookup: Lookup | None,
    name: str = "sheet",
    edits: Mapping[tuple[int, int], Edit] | None = None,
    translate: Translate | None = None,
    cm: str | None = None,
    fills: Sequence[Fill] = (),
) -> int:
    """Copy a worksheet part, replacing the cached values of formula cells that changed.

    ``edits`` replace or add cells, keyed ``(row, col)`` (1-based), and ``fills`` add ranges of
    constant cells, each value made when it is written; a cell or row the part does not hold is
    inserted in order. When the master of a shared formula is edited, its other cells get their
    own formula text, through ``translate``. Returns the number of cells changed.
    """
    remaining = _Pending(edits or {}, fills)
    editing = bool(edits or fills)  # a write of values alone skips what inserts cells and rows
    dissolved: dict[str, tuple[int, int, str]] = {}  # shared index -> master (row, col, text)
    buf = bytearray()
    base = 0  # absolute offset of buf[0]; everything before it has been written out
    patches: list[tuple[int, int, bytes]] = []
    safe = 0  # bytes before this offset are final
    count = 0
    row = col = 0
    prefix = b""
    in_c = False
    c_start = c_tag_end = 0
    c_name = ""
    c_attrs: dict[str, str] = {}
    f_attrs: dict[str, str] | None = None
    f_start = f_end = 0
    f_text: list[str] | None = None
    v_start = v_end = -1
    v_text: list[str] | None = None
    v_preserve = False
    collect = 0  # 1 <v>, 2 <f>

    def end_of(index: int, empty: bool) -> int:
        if empty:
            return index
        return buf.index(b">", index - base) + 1 + base

    def new_cells(r: int, below: int | None) -> bytes:
        """The edited cells of row ``r`` not yet written, with columns below ``below``."""
        nonlocal count
        out = bytearray()  # appended in place: a bytes += would copy all of it each time
        for c, edit in remaining.take(r, below):
            out += _edit_xml(prefix, r, c, edit, cm=cm)
            count += 1
        return bytes(out)

    def new_rows(below: int | None) -> bytes:
        out = bytearray()
        while (r := remaining.next_row()) is not None and (below is None or r < below):
            remaining.reach_row()
            cells = new_cells(r, None)
            if cells:
                out += b"<" + prefix + b'row r="' + str(r).encode() + b'">' + cells
                out += b"</" + prefix + b"row>"
        return bytes(out)

    empties: list[bool] = []

    def start(qname: str, attrs: dict[str, str]) -> None:
        nonlocal safe, row, col, prefix, in_c, c_start, c_tag_end, c_name, c_attrs
        nonlocal f_attrs, f_start, f_end, f_text, v_start, v_end, v_text, v_preserve, collect
        index = p.CurrentByteIndex
        # Where the start tag ends, and whether it is an empty element (inline: once per element).
        end = buf.index(b">", index - base) + 1
        empty = buf[end - 2] == 0x2F  # "/"
        end += base
        empties.append(empty)
        if not in_c:
            safe = index
        n = qname.rpartition(":")[2]
        if n == "c" and not in_c:
            ref = attrs.get("r")
            if ref:
                row, ref_col = _column(ref)
                # A column off the grid is placed by position, as the reader and Excel do.
                col = ref_col if 1 <= ref_col <= 16_384 else col + 1
            else:
                col += 1
            if editing and remaining.has(row):
                cells = new_cells(row, col)
                if cells:
                    patches.append((index, index, cells))
            in_c = True
            c_start, c_tag_end, c_name, c_attrs = index, end, qname, attrs
            f_attrs, f_start, f_end, f_text = None, 0, 0, None
            v_start, v_end, v_text = -1, -1, None
        elif n == "row" and not in_c:
            r = attrs.get("r")
            row = int(r) if r else row + 1
            col = 0
            rows = new_rows(row) if editing else b""
            if rows:
                patches.append((index, index, rows))
            if editing and remaining.next_row() == row:
                remaining.reach_row()
                tag = _DROP_ATTRS.sub(b"", bytes(buf[index - base : end - base]))
                if empty:  # <row r="5"/>: the row's new cells go inside it
                    tag = tag[:-2].rstrip() + b">" + new_cells(row, None)
                    tag += b"</" + qname.encode() + b">"
                patches.append((index, end, tag))
        elif n == "sheetData":
            prefix = qname.rpartition(":")[0].encode()
            prefix = prefix + b":" if prefix else b""
            if empty and remaining.next_row() is not None:
                tag = bytes(buf[index - base : end - 2 - base]).rstrip() + b">"
                patches.append((index, end, tag + new_rows(None) + b"</" + qname.encode() + b">"))
        elif in_c and n == "f":
            f_attrs, f_start, f_end = attrs, index, end
            f_text = []
            collect = 2
        elif in_c and n == "v":
            v_start = index
            v_text = []
            v_preserve = attrs.get("xml:space") == "preserve"
            collect = 1

    def end(qname: str) -> None:
        nonlocal safe, in_c, f_end, v_end, collect, count
        index = p.CurrentByteIndex
        empty = empties.pop()
        n = qname.rpartition(":")[2]
        if not in_c:
            if not editing:
                pass
            elif n == "row":
                cells = new_cells(row, None)
                if cells and not empty:
                    patches.append((index, index, cells))
            elif n == "sheetData" and not empty:
                rows = new_rows(None)
                if rows:
                    patches.append((index, index, rows))
            safe = index
            return
        if n == "f":
            f_end = end_of(index, empty)
            collect = 0
        elif n == "v":
            v_end = end_of(index, empty)
            collect = 0
        elif n == "c" and qname == c_name:
            c_end = end_of(index, empty)
            in_c = False
            safe = c_end
            shared = f_attrs is not None and f_attrs.get("t") == "shared"
            si = f_attrs.get("si") if f_attrs is not None else None
            edit = remaining.pop(row, col) if editing else None
            if edit is not None:
                if shared and si is not None and f_attrs is not None and f_attrs.get("ref"):
                    dissolved[si] = (row, col, "".join(f_text or ()))
                attrs = _DROP_ATTRS.sub(b"", bytes(buf[c_start - base : c_tag_end - base]))
                attrs = attrs[attrs.index(b" ") : -1] if b" " in attrs else b""
                attrs = attrs.rstrip(b"/").rstrip()
                patches.append((c_start, c_end, _edit_xml(prefix, row, col, edit, attrs, cm)))
                count += 1
                return
            new_f = None
            if shared and si in dissolved and not (f_attrs or {}).get("ref"):
                assert translate is not None
                assert si is not None
                r0, c0, text = dissolved[si]
                moved = translate(text, r0 - 1, c0 - 1, row - 1, col - 1)
                new_f = b"<" + prefix + b"f>" + text_escape(moved).encode("utf-8")
                new_f += b"</" + prefix + b"f>"
            new = lookup(row, col) if (f_attrs is not None and lookup is not None) else None
            if new is not None:
                old = "".join(v_text) if v_text is not None else None
                if old is not None and not v_preserve and c_attrs.get("t") == "str":
                    old = old.strip(_SPACE)  # as Excel reads it
                if _same(c_attrs.get("t", "n"), old, new):
                    new = None
            if new is not None or new_f is not None:
                patches.append((c_start, c_end, _patched(new, new_f, c_end)))
                count += 1

    def _patched(new: FileValue | None, new_f: bytes | None, c_end: int) -> bytes:
        tag = bytes(buf[c_start - base : c_tag_end - base])
        if new is not None:
            tag = _T_ATTR.sub(b"", tag)
            if new[0] != "n":
                tag = tag[:-1] + b' t="' + new[0].encode() + b'">'
        pieces: list[tuple[int, int, bytes]] = []
        if new_f is not None:
            pieces.append((f_start, f_end, new_f))
        if new is not None:
            value = _value_element(prefix_of(c_name), new)
            if v_start >= 0:
                pieces.append((v_start, v_end, value))
            else:
                at = f_end or c_tag_end
                pieces.append((at, at, value))
        out = tag
        position = c_tag_end
        for s, e, data in sorted(pieces, key=lambda x: x[0]):
            out += bytes(buf[position - base : s - base]) + data
            position = max(position, e)
        return out + bytes(buf[position - base : c_end - base])

    def text(data: str) -> None:
        if collect == 1 and v_text is not None:
            v_text.append(data)
        elif collect == 2 and f_text is not None:
            f_text.append(data)

    def flush(upto: int) -> None:
        nonlocal base
        position = base
        while patches and patches[0][0] < upto:
            s, e, data = patches.pop(0)
            write(bytes(buf[position - base : s - base]))
            write(data)
            position = e
        write(bytes(buf[position - base : upto - base]))
        del buf[: upto - base]
        base = upto

    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    first = True
    try:
        while chunk := source.read(_CHUNK):
            if first:
                _check_encoding(chunk[:200], name)
                first = False
            buf += chunk
            p.Parse(chunk, False)
            flush(max(safe, base))
        p.Parse(b"", True)
    except xml.parsers.expat.ExpatError as exc:
        raise PackageError(f"{name}: {exc}") from exc
    flush(base + len(buf))
    if remaining:
        raise PackageError(f"{name}: no sheetData to hold the edited cells")
    return count


def prefix_of(qname: str) -> bytes:
    """The namespace prefix of an element name, with its colon (``b""`` for none)."""
    prefix = qname.rpartition(":")[0].encode()
    return prefix + b":" if prefix else b""


def _full_calc_on_load(workbook: bytes) -> bytes | None:
    """The workbook part with ``<calcPr fullCalcOnLoad="1">``; ``None`` when it already has it."""
    m = _CALC.search(workbook)
    if m:
        attrs = m["a"]
        full = _FULL.search(attrs)
        if full:
            value = full[1] if full[1] is not None else full[2]
            if value in (b"1", b"true"):
                return None
            attrs = attrs[: full.start()] + b'fullCalcOnLoad="1"' + attrs[full.end() :]
        else:
            attrs = attrs + b' fullCalcOnLoad="1"'
        tag = b"<" + (m["p"] or b"") + b"calcPr" + attrs + m["e"] + b">"
        return workbook[: m.start()] + tag + workbook[m.end() :]
    root = _ROOT.search(workbook)
    prefix = (root["p"] or b"") if root else b""
    element = b"<" + prefix + b'calcPr fullCalcOnLoad="1"/>'
    names = b"|".join(re.escape(prefix + n) for n in _AFTER_CALC)
    end_tag = rb"</" + re.escape(prefix) + rb"workbook\s*>"
    after = re.search(rb"<(?:" + names + rb")\b|" + end_tag, workbook)
    if after is None:
        raise PackageError("the workbook part has no end tag")
    return workbook[: after.start()] + element + workbook[after.start() :]


def _without(data: bytes, element: bytes, attribute: bytes, value: str) -> bytes:
    """``data`` without the empty elements ``element`` whose ``attribute`` is ``value``."""
    pattern = (
        rb"<(?:[\w.-]+:)?"
        + element
        + rb"\b[^>]*?\b"
        + attribute
        + rb"""\s*=\s*["']"""
        + re.escape(value.encode("utf-8"))
        + rb"""["'][^>]*?/>"""
    )
    return re.sub(pattern, b"", data)


# The metadata part a workbook needs for dynamic-array formulas, as Excel writes it (capture
# arrays): cell metadata 1 is a dynamic array.
METADATA = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    b'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" '
    b'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    b'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    b'<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    b'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    b'</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    b"</metadata>"
)
_METADATA_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
_METADATA_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"


def write(
    source: str | os.PathLike[str] | BinaryIO,
    target: str | os.PathLike[str],
    values: Mapping[str, Lookup],
    *,
    edits: Mapping[str, Mapping[tuple[int, int], Edit]] | None = None,
    fills: Mapping[str, Sequence[Fill]] | None = None,
    translate: Translate | None = None,
    drop_calc_chain: bool = False,
    full_calc_on_load: bool = True,
    limits: ReadLimits | None = None,
) -> int:
    """Write ``source`` to ``target`` with new cached values, atomically; may be the same file.

    ``values`` maps sheet names to lookups for formula results; ``edits`` maps sheet names to
    the cells changed through the API, and ``fills`` to the ranges array formulas fill.
    ``drop_calc_chain`` removes the calculation chain (its part, relationship and content type),
    which Excel needs when a formula cell was removed.
    With ``full_calc_on_load`` the workbook asks Excel to recalculate when it opens the file;
    without it the setting stays as it was. Returns the number of cells changed.
    """
    target = Path(target)
    count = 0
    edits = edits or {}
    fills = fills or {}
    with Book(source, limits) as book:
        package = book.package
        names = set(values) | set(edits) | set(fills)
        parts = {book.sheet(n).part: n for n in names}
        part_values = {book.sheet(n).part: values.get(n) for n in names}
        part_edits = {book.sheet(n).part: edits.get(n) for n in names}
        part_fills = {book.sheet(n).part: fills.get(n, ()) for n in names}
        workbook_part = book.workbook_part
        replaced: dict[str, bytes] = {}
        if full_calc_on_load:
            new_workbook = _full_calc_on_load(package.read(workbook_part))
            if new_workbook is not None:
                replaced[workbook_part] = new_workbook
        skip: set[str] = set()
        if drop_calc_chain:
            rels_part = posixpath.join(
                posixpath.dirname(workbook_part),
                "_rels",
                posixpath.basename(workbook_part) + ".rels",
            )
            for rid, (typ, chain, external) in package.relationships(workbook_part).items():
                if typ.endswith("/calcChain") and not external:
                    skip.add(chain)
                    rels = replaced.get(rels_part, package.read(rels_part))
                    replaced[rels_part] = _without(rels, b"Relationship", b"Id", rid)
                    types = package.read("[Content_Types].xml")
                    replaced["[Content_Types].xml"] = _without(
                        types, b"Override", b"PartName", "/" + chain
                    )
        added: dict[str, bytes] = {}
        cm = None
        if any(e.dynamic for cells in edits.values() for e in cells.values()):
            if book.dynamic_cm:
                cm = min(book.dynamic_cm, key=int)
            elif not any(
                typ.endswith("/sheetMetadata")
                for typ, _, _ in package.relationships(workbook_part).values()
            ):
                cm = "1"
                folder_part = posixpath.dirname(workbook_part)
                meta = posixpath.join(folder_part, "metadata.xml")
                added[meta] = METADATA
                rels_part = posixpath.join(
                    folder_part, "_rels", posixpath.basename(workbook_part) + ".rels"
                )
                rels = replaced.get(rels_part, package.read(rels_part))
                ids = set(re.findall(rb'\bId="([^"]*)"', rels))
                n = 1
                while f"rId{n}".encode() in ids:
                    n += 1
                relationship = (
                    f'<Relationship Id="rId{n}" Type="{_METADATA_TYPE}" Target="metadata.xml"/>'
                ).encode()
                replaced[rels_part] = re.sub(
                    rb"</(\w+:)?Relationships>", lambda m: relationship + m[0], rels, count=1
                )
                types = replaced.get("[Content_Types].xml", package.read("[Content_Types].xml"))
                override = (
                    f'<Override PartName="/{meta}" ContentType="{_METADATA_CONTENT}"/>'
                ).encode()
                replaced["[Content_Types].xml"] = re.sub(
                    rb"</(\w+:)?Types>", lambda m: override + m[0], types, count=1
                )
            # else: a metadata part without the dynamic-array type: legacy array formulas
        infos = package.infos()
        folder = target.parent
        handle, temp = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=folder)
        is_path = isinstance(source, (str, os.PathLike))
        try:
            with os.fdopen(handle, "w+b") as out, ExitStack() as stack:
                raw: BinaryIO
                if isinstance(source, (str, os.PathLike)):
                    raw = stack.enter_context(open(source, "rb"))
                else:
                    raw = source
                zw = ZipWriter(out, package.comment)
                for info in infos:
                    name = info.filename
                    if name in skip:
                        continue
                    if name in replaced:
                        zw.start(info)
                        zw.write(replaced[name])
                        zw.finish()
                    elif name in parts:
                        zw.start(info, zip64=info.file_size > _ZIP64_ABOVE)
                        with package.part(name) as stream:
                            changed = patch_cells(
                                stream,
                                zw.write,
                                part_values[name],
                                name,
                                part_edits[name],
                                translate,
                                cm,
                                part_fills[name],
                            )
                        if changed:
                            zw.finish()
                            count += changed
                        else:
                            zw.discard()
                            zw.copy(raw, info)
                    else:
                        zw.copy(raw, info)
                for name, data in added.items():
                    zw.start(zipfile.ZipInfo(name, time.localtime()[:6]))
                    zw.write(data)
                    zw.finish()
                zw.close()
                out.flush()
                os.fsync(out.fileno())
            if target.exists():
                shutil.copymode(target, temp)
            elif is_path:
                shutil.copymode(source, temp)  # type: ignore[arg-type]
            else:
                os.chmod(temp, 0o666 & ~_umask())
            os.replace(temp, target)
        except BaseException:
            Path(temp).unlink(missing_ok=True)
            raise
    return count


def _umask() -> int:
    mask = os.umask(0)
    os.umask(mask)
    return mask
