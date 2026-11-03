"""How Excel compares values: operators, exact lookups and criteria (ARCHITECTURE.md §4.2-4.4).

Values of different kinds are never compared with Python's ``==``: every function here
dispatches on the kind first, so the text ``"#N/A"`` and the error ``#N/A`` stay apart even
though ``ErrorValue`` equals its text for interoperation (§4.2). Captured in
``tests/captures/compare.tsv``:

- operators order numbers < text < logicals; text compares without case (``casefold``, so
  ``"straße"="STRASSE"``); a blank is 0, ``""`` or FALSE by what it meets; an error operand
  gives that error;
- with ``comparison_rounding``, numbers that agree to 15 significant digits are equal, for the
  operators and for criteria; exact lookups (``MATCH``, ``VLOOKUP``) compare exactly;
- a lookup never matches a blank cell; ``""`` matches an empty text;
- a criterion text spelled like an error code (``"#N/A"``, ``"#n/a"``) matches that error,
  never the text; ``"="`` matches blank cells only, ``"<>"`` every cell that is not blank,
  ``""`` blank cells and empty texts.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from decimal import ROUND_HALF_DOWN, Decimal
from typing import Literal

from xlrecalc._numtext import text_to_number
from xlrecalc._quirks import Quirks
from xlrecalc._values import BLANK, ErrorValue, Value

__all__ = ["compare", "criterion", "lookup_equal", "lookup_key", "numbers_equal", "wildcard"]

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


def compare(a: Value, b: Value, op: Op, quirks: Quirks = Quirks.EXCEL) -> bool | ErrorValue:
    """``a op b`` as Excel's comparison operators compute it."""
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
        order = (x > y) - (x < y)
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

    The caller has already returned an error lookup value as the result.
    """
    if cell is BLANK or isinstance(cell, ErrorValue):
        return False
    if isinstance(lookup, bool) or isinstance(cell, bool):
        return isinstance(lookup, bool) and isinstance(cell, bool) and lookup == cell
    if isinstance(lookup, str) or isinstance(cell, str):
        return (
            isinstance(lookup, str)
            and isinstance(cell, str)
            and (lookup.casefold() == cell.casefold())
        )
    return float(lookup) == float(cell)  # type: ignore[arg-type]


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
        return ("text", value.casefold())
    if isinstance(value, float):
        return ("number", value)
    return ("blank", None)


_CRITERION = re.compile(r"(<=|>=|<>|=|<|>)?(.*)", re.S)


def criterion(value: Value, quirks: Quirks = Quirks.EXCEL) -> Callable[[Value], bool]:
    """The test a ``COUNTIF``-family criterion applies to each cell."""
    if isinstance(value, ErrorValue):
        code = value.code
        return lambda cell: isinstance(cell, ErrorValue) and cell.code == code
    if value is BLANK:
        value = ""
    if not isinstance(value, str):
        return _equal_to(value, quirks)
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
    test = _equal_to(target, quirks)
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
    return _ordered(target, op, quirks)


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


def _equal_to(target: Value, quirks: Quirks) -> Callable[[Value], bool]:
    if isinstance(target, ErrorValue):
        return lambda cell: isinstance(cell, ErrorValue) and cell.code == target.code
    if isinstance(target, bool):
        return lambda cell: isinstance(cell, bool) and cell == target
    if isinstance(target, str):
        pattern = wildcard(target)
        return lambda cell: isinstance(cell, str) and pattern.fullmatch(cell) is not None
    number = float(target)  # type: ignore[arg-type]
    return lambda cell: (
        isinstance(cell, float)
        and not isinstance(cell, bool)
        and numbers_equal(cell, number, quirks)
    )


def _ordered(target: Value, op: str, quirks: Quirks) -> Callable[[Value], bool]:
    def test(cell: Value) -> bool:
        if isinstance(cell, ErrorValue) or cell is BLANK or _rank(cell) != _rank(target):
            return False
        result = compare(cell, target, op, quirks)  # type: ignore[arg-type]
        return result is True

    return test


def wildcard(text: str) -> re.Pattern[str]:
    """``*`` any run, ``?`` one character, ``~`` escapes; without case."""
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
