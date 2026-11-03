"""A template key computed from the raw formula text, without tokenizing (ARCHITECTURE.md §5.3).

The key is the text with every cell, column-range and row-range reference rewritten relative to
the formula's cell (absolute parts kept), and with strings, quoted sheet names and quoted
external items, and bracketed parts (structured references, book indexes) left as they are.
Only references the tokenizer would read as references are rewritten: a cell whose column is
past ``XFD`` or whose row is past 1,048,576 stays text, as the tokenizer reads it as a name. So
the key can be inverted at any cell, and two formulas with the same key are the same formula
moved: the loader parses a formula only when its key is new and gives the others the same
template.

``tests/test_fastkey.py`` checks that equal raw keys imply equal parsed keys over the corpus.
"""

from __future__ import annotations

import re

from xlrecalc._formula._tokens import column_number

__all__ = ["raw_key", "translate"]

_MAX_COL = 16384
_MAX_ROW = 1048576
_SCAN = re.compile(
    r"""
      (?P<skip>"(?:[^"]|"")*" | '(?:[^']|'')*'!? | \[[^\]]*\] )
    | (?<![\w.$])(?P<cell>(?P<ca>\$?)(?P<col>[A-Za-z]{1,3})(?P<ra>\$?)(?P<row>\d+))(?![\w.(\[!])
    | (?<![\w.$])(?P<cols>(?P<c1a>\$?)(?P<c1>[A-Za-z]{1,3}):(?P<c2a>\$?)(?P<c2>[A-Za-z]{1,3}))
      (?![\w.(\[])
    | (?<![\w.$:])(?P<rows>(?P<r1a>\$?)(?P<r1>\d+):(?P<r2a>\$?)(?P<r2>\d+))(?![\w.])
    """,
    re.VERBOSE,
)


def raw_key(text: str, row: int, col: int) -> str:
    """The raw-text template key of a formula (file form) in the 0-based cell (row, col)."""
    r1, c1 = row + 1, col + 1

    def rewrite(m: re.Match[str]) -> str:
        if m["skip"] is not None:
            return m["skip"]
        if m["cell"] is not None:
            c = column_number(m["col"])
            r = int(m["row"])
            if c > _MAX_COL or not 1 <= r <= _MAX_ROW:
                return m["cell"]
            cpart = f"C${c}" if m["ca"] else f"C[{c - c1}]"
            rpart = f"R${r}" if m["ra"] else f"R[{r - r1}]"
            return f"\x1e{rpart}{cpart}\x1e"
        if m["cols"] is not None:
            a, b = column_number(m["c1"]), column_number(m["c2"])
            if max(a, b) > _MAX_COL:
                return m["cols"]
            first = f"C${a}" if m["c1a"] else f"C[{a - c1}]"
            second = f"C${b}" if m["c2a"] else f"C[{b - c1}]"
            return f"\x1e{first}:{second}\x1e"
        a, b = int(m["r1"]), int(m["r2"])
        if not (1 <= a <= _MAX_ROW and 1 <= b <= _MAX_ROW):
            return m["rows"]
        first = f"R${a}" if m["r1a"] else f"R[{a - r1}]"
        second = f"R${b}" if m["r2a"] else f"R[{b - r1}]"
        return f"\x1e{first}:{second}\x1e"

    return _SCAN.sub(rewrite, text)


def _letters(col: int) -> str:
    out = ""
    while col:
        col, rest = divmod(col - 1, 26)
        out = chr(65 + rest) + out
    return out


_REF_RANGE = re.compile(r"#REF!:\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}\$?\d+:#REF!|#REF!:#REF!")


def translate(text: str, row: int, col: int, to_row: int, to_col: int) -> str:
    """A formula (file form) in the 0-based cell (row, col), moved to (to_row, to_col).

    Relative parts of references shift with the move, absolute parts stay; a reference moved
    off the sheet becomes ``#REF!``, as when Excel fills a formula. This gives the dependent
    cells of a shared formula their own text.
    """
    dr, dc = to_row - row, to_col - col

    def shift(value: int, absolute: str, delta: int, limit: int) -> int | None:
        moved = value if absolute else value + delta
        return moved if 1 <= moved <= limit else None

    def rewrite(m: re.Match[str]) -> str:
        if m["skip"] is not None:
            return m["skip"]
        if m["cell"] is not None:
            c, r = column_number(m["col"]), int(m["row"])
            if c > _MAX_COL or not 1 <= r <= _MAX_ROW:
                return m["cell"]
            c2, r2 = shift(c, m["ca"], dc, _MAX_COL), shift(r, m["ra"], dr, _MAX_ROW)
            if c2 is None or r2 is None:
                return "#REF!"
            return f"{m['ca']}{_letters(c2)}{m['ra']}{r2}"
        if m["cols"] is not None:
            a, b = column_number(m["c1"]), column_number(m["c2"])
            if max(a, b) > _MAX_COL:
                return m["cols"]
            a2, b2 = shift(a, m["c1a"], dc, _MAX_COL), shift(b, m["c2a"], dc, _MAX_COL)
            if a2 is None or b2 is None:
                return "#REF!"
            return f"{m['c1a']}{_letters(a2)}:{m['c2a']}{_letters(b2)}"
        a, b = int(m["r1"]), int(m["r2"])
        if not (1 <= a <= _MAX_ROW and 1 <= b <= _MAX_ROW):
            return m["rows"]
        a2, b2 = shift(a, m["r1a"], dr, _MAX_ROW), shift(b, m["r2a"], dr, _MAX_ROW)
        if a2 is None or b2 is None:
            return "#REF!"
        return f"{m['r1a']}{a2}:{m['r2a']}{b2}"

    return _REF_RANGE.sub("#REF!", _SCAN.sub(rewrite, text))
