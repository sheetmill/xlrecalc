"""Streaming reads of shared strings and worksheet cells with expat (ARCHITECTURE.md §9.1).

Element names are matched by their local part (expat runs without namespace processing, the
fast path), so prefixed and Strict OOXML files read like transitional ones. Text is
unescaped from OOXML's ``_xHHHH_`` form.
"""

from __future__ import annotations

import re
import xml.parsers.expat
from collections.abc import Iterator
from typing import BinaryIO, NamedTuple

from xlrecalc._xlsx._package import PackageError

__all__ = ["Cell", "Formula", "read_cells", "read_shared_strings", "unescape"]

_CHUNK = 1 << 16
_SPACE = " \t\n\r"
_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")
_MAX_ROW = 1_048_576
_MAX_COL = 16_384
_INF = float("inf")


def unescape(text: str) -> str:
    """OOXML's escape of characters XML cannot hold: ``_x000D_`` is a carriage return."""
    if "_x" not in text:
        return text
    return _ESCAPE.sub(lambda m: chr(int(m[1], 16)), text)


_LOCAL = {"row": 1, "c": 2, "v": 3, "f": 4, "is": 5, "t": 6, "rPh": 7, "si": 8}


def _codes() -> dict[str, int]:
    """Element name → code, filled per name on first sight (prefixed names included)."""

    class Codes(dict[str, int]):
        def __missing__(self, name: str) -> int:
            code = _LOCAL.get(name.rpartition(":")[2], 0)
            self[name] = code
            return code

    return Codes()


def read_shared_strings(stream: BinaryIO) -> list[str]:
    """The shared string table: each ``<si>``'s text, phonetic runs (``<rPh>``) left out.

    Each ``<t>`` not marked ``xml:space="preserve"`` loses its leading and trailing whitespace,
    as Excel reads it (capture ``whitespace``).
    """
    strings: list[str] = []
    parts: list[str] = []
    codes = _codes()
    in_t = False
    t_start = 0
    preserve = False
    phonetic = 0

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal in_t, t_start, preserve, phonetic
        code = codes[name]
        if code == 6:
            in_t = True
            t_start = len(parts)
            preserve = attrs.get("xml:space") == "preserve"
        elif code == 7:
            phonetic += 1
        elif code == 8:
            parts.clear()

    def end(name: str) -> None:
        nonlocal in_t, phonetic
        code = codes[name]
        if code == 6:
            in_t = False
            if not preserve and not phonetic:
                parts[t_start:] = ["".join(parts[t_start:]).strip(_SPACE)]
        elif code == 7:
            phonetic -= 1
        elif code == 8:
            strings.append(unescape("".join(parts)))

    def text(data: str) -> None:
        if in_t and not phonetic:
            parts.append(data)

    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    while chunk := stream.read(_CHUNK):
        p.Parse(chunk, False)
    p.Parse(b"", True)
    return strings


class Formula(NamedTuple):
    """A cell's ``<f>``: its text (``None`` for a shared formula's dependent cells) and kind.

    ``kind`` is ``normal``, ``shared``, ``array`` or ``dataTable``; ``ref`` the range of a shared
    formula's master or an array formula; ``si`` the shared-formula index; ``attributes`` all
    attributes as written.
    """

    text: str | None
    kind: str
    ref: str | None
    si: int | None
    attributes: dict[str, str]


class Cell(NamedTuple):
    """One ``<c>``: position (1-based), cached value, its type, formula, and ``cm``/``vm``.

    ``type`` is ``n`` (number, ``value`` a float), ``s`` (text from the shared strings),
    ``str`` (a formula's text result), ``inlineStr``, ``b`` (``value`` a bool), ``e`` (an
    error, ``value`` its code), ``d`` (an ISO 8601 date, ``value`` its text) or ``none`` (no
    value: an uncached formula or a styled empty cell).
    """

    row: int
    column: int
    value: float | str | bool | None
    type: str
    formula: Formula | None
    cm: str | None = None
    vm: str | None = None


def _column(ref: str) -> tuple[int, int]:
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


def _row(text: str) -> int:
    """A row number from the file; Excel refuses one off its grid (capture ``value_forms``)."""
    try:
        row = int(text)
    except ValueError:
        raise PackageError(f"not a row number: {text!r}") from None
    if not 1 <= row <= _MAX_ROW:
        raise PackageError(f"row {row} is off the sheet")
    return row


def _number(text: str, row: int, col: int) -> tuple[float | str, str]:
    """A ``<v>`` of a number cell as Excel reads it (capture ``value_forms``).

    Python's ``float`` takes more than Excel: ``nan``, ``inf``, ``+5`` and ``1_000``, which make
    Excel refuse the file, are refused here too. A number too large for a double is ``#NULL!``,
    as Excel reads it.
    """
    try:
        value = float(text)
    except ValueError:
        raise PackageError(f"row {row}, column {col}: not a number: {text[:40]!r}") from None
    if -_INF < value < _INF and "_" not in text and text[0] != "+" and text[0] not in _SPACE:
        return value, "n"
    if -_INF < value < _INF and "_" not in text and text.lstrip()[:1] != "+":
        return value, "n"  # leading whitespace, which Excel allows
    if (
        "_" in text
        or text.lstrip()[:1] == "+"
        or any(ch.isalpha() and ch not in "eE" for ch in text)
    ):
        raise PackageError(f"row {row}, column {col}: not a number Excel reads: {text[:40]!r}")
    return "#NULL!", "e"


def read_cells(stream: BinaryIO, strings: list[str]) -> Iterator[Cell]:
    """The cells of a worksheet part, in file order.

    The element handlers are the hot path of a load: expat runs without namespace processing
    (names are matched by their local part, so a prefixed file reads the same), and the state
    lives in closure variables.
    """
    out: list[Cell] = []
    buffer: list[str] = []
    codes = _codes()
    row = col = 0
    attrs_c: dict[str, str] = {}
    v: str | None = None
    v_preserve = False
    t_preserve = False
    f_text: str | None = None
    f_attrs: dict[str, str] | None = None
    inline: list[str] | None = None
    collect = 0  # 3 <v>, 4 <f>, 6 <t> inside <is>
    in_is = phonetic = 0

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal row, col, attrs_c, v, v_preserve, t_preserve, f_text, f_attrs, inline
        nonlocal collect, in_is, phonetic
        code = codes[name]
        if code == 2:
            ref = attrs.get("r")
            if ref:
                try:
                    ref_row, ref_col = _column(ref)
                except ValueError:
                    raise PackageError(f"not a cell reference: {ref[:40]!r}") from None
                if not 1 <= ref_row <= _MAX_ROW:
                    raise PackageError(f"row {ref_row} is off the sheet")
                row = ref_row
                # Excel places a cell whose column is off its grid by position (capture
                # value_forms: XFE1 after nothing is A1).
                col = ref_col if 1 <= ref_col <= _MAX_COL else col + 1
            else:
                col += 1
            if col > _MAX_COL:
                raise PackageError(f"row {row}: more cells than columns")
            attrs_c = attrs
            v = f_text = None
            v_preserve = False
            f_attrs = inline = None
        elif code == 3:
            buffer.clear()
            collect = 3
            v_preserve = attrs.get("xml:space") == "preserve"
        elif code == 4:
            buffer.clear()
            collect = 4
            f_attrs = attrs
        elif code == 1:
            r = attrs.get("r")
            row = _row(r) if r else row + 1
            if row > _MAX_ROW:
                raise PackageError(f"row {row} is off the sheet")
            col = 0
        elif code == 5:
            in_is += 1
            inline = []
        elif code == 7:
            phonetic += 1
        elif code == 6 and in_is:
            buffer.clear()
            collect = 6
            t_preserve = attrs.get("xml:space") == "preserve"

    def end(name: str) -> None:
        nonlocal v, f_text, collect, in_is, phonetic
        code = codes[name]
        if code == 2:
            out.append(_cell(row, col, attrs_c, v, v_preserve, f_text, f_attrs, inline, strings))
        elif code == 3:
            v = "".join(buffer)
            collect = 0
        elif code == 4:
            f_text = "".join(buffer)
            collect = 0
        elif code == 6 and collect == 6:
            if not phonetic and inline is not None:
                run = "".join(buffer)
                inline.append(run if t_preserve else run.strip(_SPACE))
            collect = 0
        elif code == 7:
            phonetic -= 1
        elif code == 5:
            in_is -= 1

    def text(data: str) -> None:
        if collect:
            buffer.append(data)

    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler, p.EndElementHandler, p.CharacterDataHandler = start, end, text
    while chunk := stream.read(_CHUNK):
        p.Parse(chunk, False)
        yield from out
        out.clear()
    p.Parse(b"", True)
    yield from out


def _cell(
    row: int,
    col: int,
    attrs: dict[str, str],
    v: str | None,
    v_preserve: bool,
    f_text: str | None,
    f_attrs: dict[str, str] | None,
    inline: list[str] | None,
    strings: list[str],
) -> Cell:
    t = attrs.get("t", "n")
    value: float | str | bool | None
    if t == "n":
        if v:
            value, t = _number(v, row, col)
        else:
            value, t = None, "none"
    elif t == "inlineStr":
        value = unescape("".join(inline)) if inline is not None else None
    elif v is None or (v == "" and t != "str"):
        value, t = None, "none"
    elif t == "s":
        try:
            index = int(v)
        except ValueError:
            index = -1
        if not 0 <= index < len(strings):
            raise PackageError(f"row {row}, column {col}: no shared string {v[:40]!r}")
        value = strings[index]
    elif t == "str":
        # Excel drops leading and trailing whitespace of a <v> not marked xml:space="preserve"
        value = unescape(v if v_preserve else v.strip(_SPACE))
    elif t == "b":
        value = v in ("1", "true")
    else:  # "e", "d"
        value = v
    formula = None
    if f_attrs is not None:
        si = f_attrs.get("si")
        formula = Formula(
            unescape(f_text) if f_text else None,  # Excel decodes escapes in formulas too
            f_attrs.get("t", "normal"),
            f_attrs.get("ref"),
            _shared_index(si, row, col) if si is not None else None,
            f_attrs,
        )
    return Cell(row, col, value, t, formula, attrs.get("cm"), attrs.get("vm"))


def _shared_index(text: str, row: int, col: int) -> int:
    try:
        return int(text)
    except ValueError:
        raise PackageError(
            f"row {row}, column {col}: not a shared formula index: {text[:40]!r}"
        ) from None
