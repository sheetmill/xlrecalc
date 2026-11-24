"""How Excel compares values: operators, exact lookups and criteria (ARCHITECTURE.md §4.2-4.4).

Values of different kinds are never compared with Python's ``==``: every function here
dispatches on the kind first, so the text ``"#N/A"`` and the error ``#N/A`` stay apart even
though ``ErrorValue`` equals its text for interoperation (§4.2). Captured in
``tests/captures/compare.tsv``:

- operators order numbers < text < logicals; text compares without case (``casefold``, so
  ``"straße"="STRASSE"``), marks before digits before letters, in the platform's order of text
  (captures ``modern``, ``collate_chars``, ``collate_words``, on the Mac and on Windows); a blank
  is 0, ``""`` or FALSE by what it meets; an error operand
  gives that error;
- with ``comparison_rounding``, numbers that agree to 15 significant digits are equal, for the
  operators and for criteria; exact lookups (``MATCH``, ``VLOOKUP``) compare exactly;
- an exact lookup's scan ignores case only (``lower``: ``"ß"`` does not find ``"ss"``), while
  a criterion's equality is ``=``'s (``COUNTIF(r,"ß")`` counts ``"ss"``; capture
  ``collate_functions``);
- a lookup never matches a blank cell; ``""`` matches an empty text;
- a criterion text spelled like an error code (``"#N/A"``, ``"#n/a"``) matches that error,
  never the text; ``"="`` matches blank cells only, ``"<>"`` every cell that is not blank,
  ``""`` blank cells and empty texts.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Callable
from decimal import ROUND_HALF_DOWN, Decimal
from typing import Literal

from xlrecalc._numtext import text_to_number
from xlrecalc._quirks import Quirks
from xlrecalc._values import BLANK, ErrorValue, Value

__all__ = [
    "compare",
    "criterion",
    "lookup_equal",
    "lookup_key",
    "numbers_equal",
    "switch_equal",
    "wildcard",
]

Op = Literal["=", "<>", "<", ">", "<=", ">="]
_CODE = re.compile(r"#[A-Za-z0-9_/]+[!?]?")


def numbers_equal(a: float, b: float, quirks: Quirks = Quirks.EXCEL) -> bool:
    """Whether two numbers are equal under ``comparison_rounding``."""
    if a == b:
        return True
    if not quirks.comparison_rounding or a == 0 or b == 0:
        return False
    return _fifteen(a) == _fifteen(b)


def _fifteen(x: float) -> Decimal:
    d = Decimal(x)
    return d.quantize(Decimal(1).scaleb(d.adjusted() - 14), rounding=ROUND_HALF_DOWN)


def _rank(v: Value) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def _resolve_blank(v: Value, other: Value) -> Value:
    if v is not BLANK:
        return v
    if isinstance(other, bool):
        return False
    if isinstance(other, str):
        return ""
    return 0.0


# Excel for Mac's order of text (captures ``modern``, ``collate_chars``, ``collate_words``): after
# folding case, marks and symbols in this order, none ignored, then digits, then letters, with
# ``æ``, ``ð``, ``ø``, ``œ``, ``þ``, ``µ`` letters of their own; ``ª``, ``º`` are ``a``, ``o``;
# an accent counts only between texts otherwise equal (``e`` < ``é`` < ``f``); ``ß`` folds to
# ``ss``. A character the captures leave out sorts after all of these, by code point (our choice,
# CAPTURING.md), and a control character before them.
MAC_ORDER = " _-,;:!¡?¿.'\"«»()[]{}§¶@*/\\&#%`^+<=>|~$0123456789aæbcdðefghijklmnoøœpqrstuvwxyzþμ"
"""Excel for Mac's primary order of the characters captured, as they are after ``casefold``."""

MAC_ALIASES = {"ª": "a", "º": "o"}
"""Characters Excel for Mac orders as another (``ª`` equals ``a``)."""


def collation(order: str, aliases: dict[str, str]) -> Callable[[str], tuple[list[int], list[int]]]:
    """A sort key for folded text: primary weights from ``order``, accents for ties."""
    weights = {ch: i for i, ch in enumerate(order, 1)}
    weights.update({ch: weights[to] for ch, to in aliases.items()})
    beyond = len(order) + 1

    def key(text: str) -> tuple[list[int], list[int]]:
        bases: list[int] = []
        accents: list[int] = []
        for ch in unicodedata.normalize("NFD", text):
            if unicodedata.combining(ch):
                accents.append(ord(ch))
                continue
            weight = weights.get(ch)
            if weight is None:
                code = ord(ch)
                weight = code - 0x20 if code < 0x20 else beyond + code
            bases.append(weight)
        return bases, accents

    return key


COLLATION = collation(MAC_ORDER, MAC_ALIASES)
"""Excel for Mac's order of text."""

# Excel for Windows' order (captures windows/collate_chars, collate_words, modern): its word sort.
# Case is ignored; ``'`` and ``-`` count only between texts otherwise equal, the later position
# first and ``'`` before ``-``; then marks in code order, ``+ < = >``, ``« » § µ ¶``, digits,
# letters. ``æ``, ``œ``, ``þ`` expand to two letters and equal them (``"ae"="æ"``); ``ª``, ``º``,
# ``ð``, ``ø`` are their base letter with a mark of their own, and an accent counts only between
# texts otherwise equal. A character the captures leave out sorts after all of these, by code
# point, and a control character before them, as on the Mac.
WINDOWS_ORDER = ' !"#$%&()*,./:;?@[\\]^_`{|}~¡¿+<=>«»§μ¶0123456789abcdefghijklmnopqrstuvwxyz'
"""Excel for Windows' primary order of the characters captured, after ``casefold``."""

_WINDOWS_WEIGHTS = {ch: i for i, ch in enumerate(WINDOWS_ORDER, 1)}
_WINDOWS_BEYOND = len(WINDOWS_ORDER) + 1
_WINDOWS_SPECIAL = {"'": 0, "-": 1}
_WINDOWS_EXPAND = {"æ": "ae", "œ": "oe", "þ": "th"}
_WINDOWS_MARKED = {"ª": ("a", 1), "º": ("o", 1), "ð": ("d", 9), "ø": ("o", 9)}

WindowsKey = tuple[tuple[int, ...], tuple[tuple[int, ...], ...], tuple[tuple[int, int], ...]]


def windows_key(text: str) -> WindowsKey:
    """Excel for Windows' sort key of folded text: letters, then marks, then ``'`` and ``-``."""
    bases: list[int] = []
    marks: list[tuple[int, ...]] = []
    special: list[tuple[int, int]] = []
    for i, ch in enumerate(text):
        if ch in _WINDOWS_SPECIAL:
            special.append((-i, _WINDOWS_SPECIAL[ch]))
            continue
        for c in _WINDOWS_EXPAND.get(ch, ch):
            marked = _WINDOWS_MARKED.get(c)
            if marked is not None:
                bases.append(_WINDOWS_WEIGHTS[marked[0]])
                marks.append((marked[1],))
                continue
            d = unicodedata.normalize("NFD", c)
            weight = _WINDOWS_WEIGHTS.get(d[0])
            if weight is None:
                code = ord(d[0])
                weight = code - 0x20 if code < 0x20 else _WINDOWS_BEYOND + code
            bases.append(weight)
            marks.append(tuple(2 + ord(m) for m in d[1:]))
    return tuple(bases), tuple(marks), tuple(special)


COLLATIONS: dict[str, Callable[[str], object]] = {"mac": COLLATION, "windows": windows_key}
"""The one place the order of text is chosen, by platform; every comparison of text reads it."""


def compare(
    a: Value, b: Value, op: Op, quirks: Quirks = Quirks.EXCEL, platform: str = "mac"
) -> bool | ErrorValue:
    """``a op b`` as Excel's comparison operators compute it, in the platform's order of text."""
    if isinstance(a, ErrorValue):
        return a
    if isinstance(b, ErrorValue):
        return b
    a, b = _resolve_blank(a, b), _resolve_blank(b, a)
    if a is BLANK:  # both blank
        a = b = 0.0
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        order = -1 if ra < rb else 1
    elif ra == 1:
        x, y = str(a).casefold(), str(b).casefold()
        if x == y:
            order = 0
        elif x.isalnum() and y.isalnum() and x.isascii() and y.isascii():
            order = 1 if x > y else -1  # digits before letters, as the code points are
        else:
            key = COLLATIONS[platform]
            kx, ky = key(x), key(y)
            order = 0 if kx == ky else (1 if kx > ky else -1)  # type: ignore[operator]
    elif ra == 2:
        order = int(bool(a)) - int(bool(b))
    else:
        p, q = float(a), float(b)  # type: ignore[arg-type]
        order = 0 if numbers_equal(p, q, quirks) else (1 if p > q else -1)
    return {
        "=": order == 0,
        "<>": order != 0,
        "<": order < 0,
        ">": order > 0,
        "<=": order <= 0,
        ">=": order >= 0,
    }[op]


def lookup_equal(lookup: Value, cell: Value) -> bool:
    """Whether an exact lookup (``MATCH`` type 0, ``VLOOKUP`` FALSE) matches a cell.

    Text ignores case only, without ``=``'s folding: ``MATCH("ß",r,0)`` passes over ``"ss"``
    (capture ``collate_functions``). The caller has already returned an error lookup value as the
    result.
    """
    if cell is BLANK or isinstance(cell, ErrorValue):
        return False
    if isinstance(lookup, bool) or isinstance(cell, bool):
        return isinstance(lookup, bool) and isinstance(cell, bool) and lookup == cell
    if isinstance(lookup, str) or isinstance(cell, str):
        return (
            isinstance(lookup, str) and isinstance(cell, str) and (lookup.lower() == cell.lower())
        )
    return float(lookup) == float(cell)  # type: ignore[arg-type]


def switch_equal(a: Value, b: Value) -> bool:
    """``SWITCH``'s test: a blank is 0, no text meets a number, and text ignores case."""
    a = 0.0 if a is BLANK else a
    b = 0.0 if b is BLANK else b
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if isinstance(a, str) or isinstance(b, str):
        return isinstance(a, str) and isinstance(b, str) and a.casefold() == b.casefold()
    return a == b


def lookup_key(value: Value) -> tuple[str, object]:
    """A dictionary key for exact-lookup caches, equal exactly when ``lookup_equal`` matches.

    The kind comes first, so the text ``"#N/A"`` and the error ``#N/A`` (equal under Python's
    ``==`` for interoperation) are different keys, as are 1 and TRUE. Blanks and errors get keys
    of their own, though a lookup never matches them.
    """
    if isinstance(value, ErrorValue):
        return ("error", value.code)
    if isinstance(value, bool):
        return ("logical", value)
    if isinstance(value, str):
        return ("text", value.lower())
    if isinstance(value, float):
        return ("number", value)
    return ("blank", None)


_WILD = re.compile(r"[*?~]")
_CRITERION = re.compile(r"(<=|>=|<>|=|<|>)?(.*)", re.S)


def criterion(
    value: Value, quirks: Quirks = Quirks.EXCEL, platform: str = "mac"
) -> Callable[[Value], bool]:
    """The test a ``COUNTIF``-family criterion applies to each cell, in the platform's order."""
    if isinstance(value, ErrorValue):
        code = value.code
        return lambda cell: isinstance(cell, ErrorValue) and cell.code == code
    if value is BLANK:  # a blank cell as the criterion is 0 (capture criteria)
        value = 0.0
    if not isinstance(value, str):
        return _equal_to(value, quirks, platform)
    m = _CRITERION.fullmatch(value)
    assert m is not None
    op, operand = m[1], m[2]
    if op is None and operand == "":
        return lambda cell: cell is BLANK or cell == ""
    if op == "=" and operand == "":
        return lambda cell: cell is BLANK
    if op == "<>" and operand == "":
        return lambda cell: cell is not BLANK
    target = _criterion_value(operand, quirks)
    test = _equal_to(target, quirks, platform)
    if isinstance(target, float):
        # A criterion that reads as a number also matches a text cell spelling it the same
        # way (COUNTIF(r,"3") counts the text "3"; captured).
        number_test, spelled = test, operand.casefold()

        def test(cell: Value) -> bool:
            return number_test(cell) or (isinstance(cell, str) and cell.casefold() == spelled)

    if op in (None, "="):
        return test
    if op == "<>":
        return lambda cell: not test(cell)
    return _ordered(target, op, quirks, platform)


def _criterion_value(operand: str, quirks: Quirks) -> Value:
    if _CODE.fullmatch(operand):
        try:
            return ErrorValue(operand.upper())
        except ValueError:
            pass
    if operand.upper() in ("TRUE", "FALSE"):
        return operand.upper() == "TRUE"
    number = text_to_number(operand, quirks)
    return operand if number is None else number


def _equal_to(target: Value, quirks: Quirks, platform: str) -> Callable[[Value], bool]:
    if isinstance(target, ErrorValue):
        return lambda cell: isinstance(cell, ErrorValue) and cell.code == target.code
    if isinstance(target, bool):
        return lambda cell: isinstance(cell, bool) and cell == target
    if isinstance(target, str):
        if not _WILD.search(target):  # equality as ``=`` has it: ``"ß"`` counts ``"ss"``
            return _text_equal_to(target, platform)
        pattern = wildcard(utf16_units(target))
        return lambda cell: (
            isinstance(cell, str) and pattern.fullmatch(utf16_units(cell)) is not None
        )
    number = float(target)  # type: ignore[arg-type]
    return lambda cell: (
        isinstance(cell, float)
        and not isinstance(cell, bool)
        and numbers_equal(cell, number, quirks)
    )


def _text_equal_to(target: str, platform: str) -> Callable[[Value], bool]:
    """``cell = target`` for text, as ``compare`` has it, with the target folded once."""
    folded = target.casefold()
    ascii_target = folded.isascii()
    collate = COLLATIONS[platform]
    key = collate(folded)

    def test(cell: Value) -> bool:
        if not isinstance(cell, str):
            return False
        text = cell.casefold()
        if text == folded:
            return True
        if ascii_target and text.isascii():  # ASCII keys are equal only when the texts are
            return False
        return collate(text) == key

    return test


def _ordered(target: Value, op: str, quirks: Quirks, platform: str) -> Callable[[Value], bool]:
    def test(cell: Value) -> bool:
        if isinstance(cell, ErrorValue) or cell is BLANK or _rank(cell) != _rank(target):
            return False
        result = compare(cell, target, op, quirks, platform)  # type: ignore[arg-type]
        return result is True

    return test


def utf16_units(text: str) -> str:
    """The text with each character outside the first plane as its two surrogates.

    Excel's text is UTF-16, and a wildcard's ``?`` matches one unit of it, so such a character
    takes ``??`` (capture text_family): patterns and texts are matched in this form.
    """
    if text.isascii() or max(text) <= "\uffff":
        return text
    raw = text.encode("utf-16-le")
    return "".join(chr(raw[i] | raw[i + 1] << 8) for i in range(0, len(raw), 2))


def utf16_length(text: str) -> int:
    """The text's length in UTF-16 units: a character outside the first plane counts two.

    Excel's limit of a text, 32,767, counts these units (capture text_family).
    """
    if text.isascii():
        return len(text)
    return len(text) + sum(1 for ch in text if ch > "\uffff")


def wildcard(text: str) -> re.Pattern[str]:
    """``*`` any run, ``?`` one UTF-16 unit, ``~`` escapes; without case.

    Give it the pattern as ``utf16_units`` gives it, and match texts in that form too.
    """
    out, i = [], 0
    while i < len(text):
        ch = text[i]
        if ch == "~" and i + 1 < len(text) and text[i + 1] in "*?~":
            out.append(re.escape(text[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.S)
