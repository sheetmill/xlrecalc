"""Split formula text into tokens that keep every character of the source.

Each token carries the whitespace before it (``space``) and its own text, so joining
``space + text`` over all tokens, the end token included, gives back the source exactly
(ARCHITECTURE.md §5.2). The text is the file form: no leading ``=``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum

__all__ = ["Kind", "Token", "TokenizeError", "tokenize"]


class Kind(Enum):
    """What a token is."""

    NUMBER = "number"
    STRING = "string"
    BOOL = "bool"
    ERROR = "error"
    CELL = "cell"  # A1, $A$1
    COLUMNS = "columns"  # A:C, $A:$C
    ROWS = "rows"  # 1:3, $1:$3
    NAME = "name"  # a defined name, a table name, a LAMBDA parameter, [1]!'quoted item'
    FUNCTION = "function"  # a name directly followed by "("
    STRUCTURED = "structured"  # Table1[Col], Table1[[#This Row],[Col]], [@Col]
    PREFIX = "prefix"  # Sheet1!, 'My Sheet'!, Sheet1:Sheet3!, [1]Sheet1!, [1]!
    OPERATOR = "operator"  # + - * / ^ & = <> < > <= >= % : @ and the spill operator #
    OPEN = "open"  # (
    CLOSE = "close"  # )
    COMMA = "comma"  # , (argument separator, union, array column separator)
    SEMICOLON = "semicolon"  # ; (array row separator)
    ARRAY_OPEN = "array_open"  # {
    ARRAY_CLOSE = "array_close"  # }
    END = "end"  # carries trailing whitespace; its text is empty


@dataclass(slots=True)
class Token:
    """One token: the whitespace before it, its text, and where the text starts."""

    kind: Kind
    text: str
    space: str
    pos: int


class TokenizeError(ValueError):
    """The text has a character no token can start with."""

    def __init__(self, message: str, pos: int) -> None:
        super().__init__(f"{message} at position {pos}")
        self.pos = pos


_SPACE = re.compile(r"[ \t\r\n]*")
_NUMBER = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
_STRING = re.compile(r'"(?:[^"]|"")*"')
# Every error code shape (as xlformat's ExcelError accepts); "#N/A" included.
_ERROR = re.compile(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|[A-Z0-9_]+[!?])")
_ROWS = re.compile(r"\$?\d+:\$?\d+")
_COLUMNS = re.compile(r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w.(\[])")
_CELL = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)(?![\w.(\[!])")
# A name starts with a letter, "_" or "\" and continues with letters, digits, "_", ".", "?", "\".
_NAME = re.compile(r"[^\W\d]?[\w.?\\]*|\\[\w.?\\]*", re.UNICODE)
_NAME_START = re.compile(r"[^\W\d]|_|\\", re.UNICODE)
_QUOTED_PREFIX = re.compile(r"'(?:[^']|'')+'!")
_QUOTED_NAME = re.compile(r"'(?:[^']|'')+'")
_BOOK_ONLY = re.compile(r"\[\d+\]!")
# [1]Sheet1!, [1]!, Sheet1!, Sheet1:Sheet3!, and the same with a book index.
_PLAIN_PREFIX = re.compile(r"(?:\[\d+\])?(?:[^\W\d][\w.]*(?::[^\W\d][\w.]*)?)?!", re.UNICODE)
_OPERATORS = ("<>", "<=", ">=", "+", "-", "*", "/", "^", "&", "=", "<", ">", "%", ":", "@")
_SINGLE = {
    "(": Kind.OPEN,
    ")": Kind.CLOSE,
    ",": Kind.COMMA,
    ";": Kind.SEMICOLON,
    "{": Kind.ARRAY_OPEN,
    "}": Kind.ARRAY_CLOSE,
}
_MAX_COLUMN = 16384
_MAX_ROW = 1048576
# Tokens after which "#" is the spill operator rather than the start of an error literal.
_SPILLABLE = {Kind.CELL, Kind.NAME, Kind.CLOSE, Kind.STRUCTURED}


def column_number(letters: str) -> int:
    """1-based column number of column letters such as ``"AB"``."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _is_cell(match: re.Match[str]) -> bool:
    return column_number(match[1]) <= _MAX_COLUMN and 1 <= int(match[2]) <= _MAX_ROW


def _bracket_end(text: str, i: int) -> int:
    """Index after the bracket group opening at ``i``; ``'`` escapes the next character."""
    depth = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "'":
            i += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    raise TokenizeError("unclosed [", i)


def tokenize(text: str) -> list[Token]:
    """Split a formula (file form, without ``=``) into tokens, ending with an ``END`` token.

    Raises:
        TokenizeError: A character starts no token, or a bracket or quote is not closed.
    """
    tokens: list[Token] = []
    i, n = 0, len(text)
    while True:
        m = _SPACE.match(text, i)
        assert m is not None
        space, i = m.group(), m.end()
        if i >= n:
            tokens.append(Token(Kind.END, "", space, i))
            return tokens
        kind, end = _next(text, i, tokens[-1] if tokens else None, bool(space))
        tokens.append(Token(kind, text[i:end], space, i))
        i = end


def _next(text: str, i: int, previous: Token | None, spaced: bool) -> tuple[Kind, int]:
    ch = text[i]
    if ch == '"':
        m = _STRING.match(text, i)
        if not m:
            raise TokenizeError("unclosed string", i)
        return Kind.STRING, m.end()
    if ch in _SINGLE:
        return _SINGLE[ch], i + 1
    if ch == "#":
        after_reference = previous is not None and previous.kind in _SPILLABLE and not spaced
        m = _ERROR.match(text, i)
        if m and not after_reference:
            return Kind.ERROR, m.end()
        if after_reference:
            return Kind.OPERATOR, i + 1
        raise TokenizeError("unknown error literal", i)
    if ch == "'":
        m = _QUOTED_PREFIX.match(text, i)
        if m:
            return Kind.PREFIX, m.end()
        # [1]!'item': a name in another workbook, or a DDE item, quoted because of its characters.
        if previous is not None and _BOOK_ONLY.fullmatch(previous.text) and not spaced:
            m = _QUOTED_NAME.match(text, i)
            if m:
                return Kind.NAME, m.end()
        raise TokenizeError("unclosed sheet name", i)
    if ch == "[":
        m = _PLAIN_PREFIX.match(text, i)
        if m:
            return Kind.PREFIX, m.end()
        return Kind.STRUCTURED, _bracket_end(text, i)
    for op in _OPERATORS:
        if text.startswith(op, i):
            return Kind.OPERATOR, i + len(op)
    if ch == "$" or ch.isdigit() or ch == ".":
        if m := _ROWS.match(text, i):
            return Kind.ROWS, m.end()
        if (m := _CELL.match(text, i)) and _is_cell(m):
            return Kind.CELL, m.end()
        if m := _COLUMNS.match(text, i):
            return Kind.COLUMNS, m.end()
        if ch != "$" and (m := _NUMBER.match(text, i)):
            return Kind.NUMBER, m.end()
        raise TokenizeError("unexpected character", i)
    if _NAME_START.match(ch):
        if (m := _PLAIN_PREFIX.match(text, i)) and m.end() > i + 1:
            return Kind.PREFIX, m.end()
        if m := _COLUMNS.match(text, i):
            return Kind.COLUMNS, m.end()
        if (m := _CELL.match(text, i)) and _is_cell(m):
            return Kind.CELL, m.end()
        m = _NAME.match(text, i)
        assert m is not None
        end = m.end()
        if end < len(text) and text[end] == "(":
            return Kind.FUNCTION, end
        if end < len(text) and text[end] == "[":
            return Kind.STRUCTURED, _bracket_end(text, end)
        if text[i:end].upper() in ("TRUE", "FALSE"):
            return Kind.BOOL, end
        return Kind.NAME, end
    raise TokenizeError("unexpected character", i)
