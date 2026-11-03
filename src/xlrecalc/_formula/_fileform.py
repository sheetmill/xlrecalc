"""Between the formula bar's form of a formula and the form a file holds (ARCHITECTURE.md §5.2).

``file_form`` is the file-form printer: what the API takes (the formula bar's form, as typed:
``A1#``, ``@``, no ``_xlfn.`` prefixes) becomes the text Excel writes into a file, so Excel
opens it without a repair. ``display_form`` goes the other way, for showing a formula read
from a file. Both work on tokens, so strings, sheet names and everything else pass through
unchanged; neither needs the text to parse.

What ``file_form`` does, as Excel's own file form shows (captures ``formula_forms``,
``formula_bar``):
- function names in upper case, with the prefix the file form needs (``_xlfn.IFNA``,
  ``_xlfn._xlws.SORT``), from ``PREFIXES``; ``LET`` and ``LAMBDA`` parameters with ``_xlpm.``;
- ``A1#`` as ``_xlfn.ANCHORARRAY(A1)``; ``@`` dropped (a formula written by the package is not a
  dynamic-array formula, where implicit intersection is already the rule);
- number literals in Excel's form: the first 15 significant digits (``.5`` as ``0.5``, ``1E3``
  as ``1000``); cell references, logicals and error literals in upper case;
- whitespace at either end dropped, a tab as a space, CR LF and CR as LF, nothing between the
  parentheses of an empty argument list (Excel refuses ``SUM( )`` in a file).
"""

from __future__ import annotations

import re

from xlrecalc._formula._prefixes import PREFIXES
from xlrecalc._formula._tokens import Kind, Token, tokenize
from xlrecalc._numtext import number_to_text, text_to_number
from xlrecalc._report import sheet_prefix

__all__ = ["display_form", "file_form"]

_FN_PREFIXES = ("_XLFN._XLWS.", "_XLFN.", "_XLUDF.")
_PARAM = "_xlpm."


def _bare(name: str) -> str:
    upper = name.upper()
    for prefix in _FN_PREFIXES:
        if upper.startswith(prefix):
            return name[len(prefix) :]
    return name


def _text(tokens: list[Token]) -> str:
    return "".join(t.space + t.text for t in tokens)


def _matching_close(tokens: list[Token], open_index: int) -> int:
    depth = 0
    for k in range(open_index, len(tokens)):
        kind = tokens[k].kind
        if kind in (Kind.OPEN,):
            depth += 1
        elif kind is Kind.CLOSE:
            depth -= 1
            if depth == 0:
                return k
    raise ValueError("unbalanced parentheses")


def _matching_open(tokens: list[Token], close_index: int) -> int:
    depth = 0
    for k in range(close_index, -1, -1):
        kind = tokens[k].kind
        if kind is Kind.CLOSE:
            depth += 1
        elif kind is Kind.OPEN:
            depth -= 1
            if depth == 0:
                return k
    raise ValueError("unbalanced parentheses")


def _primary(tokens: list[Token]) -> bool:
    """Whether the tokens form one operand that needs no parentheses before ``#`` or after ``@``."""
    core = [t for t in tokens if t.kind is not Kind.END]
    if len(core) >= 2 and core[0].kind is Kind.PREFIX:
        core = core[1:]
    if len(core) == 3 and core[1].text == ":" and core[0].kind is core[2].kind is Kind.CELL:
        return True
    if len(core) == 1:
        return core[0].kind in (Kind.CELL, Kind.NAME, Kind.STRUCTURED, Kind.COLUMNS, Kind.ROWS)
    if core and core[0].kind is Kind.FUNCTION and core[-1].kind is Kind.CLOSE:
        return _matching_close(core, 1) == len(core) - 1
    return False


def display_form(text: str) -> str:
    """A formula as a file holds it, in the formula bar's form, without the leading ``=``.

    ``_xlfn.`` and ``_xlpm.`` prefixes go, ``_xlfn.ANCHORARRAY(A1)`` becomes ``A1#`` and
    ``_xlfn.SINGLE(x)`` becomes ``@x``. Text that does not tokenize is returned as it is.
    """
    try:
        tokens = tokenize(text)
    except ValueError:
        return text
    out: list[Token] = []
    k = 0
    while k < len(tokens):
        tok = tokens[k]
        if tok.kind is Kind.FUNCTION:
            bare = _bare(tok.text)
            special = bare.upper()
            if special in ("ANCHORARRAY", "SINGLE") and bare != tok.text:
                close = _matching_close(tokens, k + 1)
                inner = tokens[k + 2 : close]
                inner_text = display_form(_text(inner).strip())
                if not _primary(tokenize(inner_text)):
                    inner_text = f"({inner_text})"
                new = inner_text + "#" if special == "ANCHORARRAY" else "@" + inner_text
                out.append(Token(Kind.NAME, new, tok.space, tok.pos))
                k = close + 1
                continue
            out.append(Token(tok.kind, bare, tok.space, tok.pos))
        elif tok.kind is Kind.NAME and tok.text.lower().startswith(_PARAM):
            out.append(Token(tok.kind, tok.text[len(_PARAM) :], tok.space, tok.pos))
        else:
            out.append(tok)
        k += 1
    return _text(out)


def _parameters(tokens: list[Token]) -> set[str]:
    """Names that ``LET`` or ``LAMBDA`` calls declare, folded."""
    names: set[str] = set()
    for k, tok in enumerate(tokens):
        if tok.kind is not Kind.FUNCTION:
            continue
        which = _bare(tok.text).upper()
        if which not in ("LET", "LAMBDA"):
            continue
        close = _matching_close(tokens, k + 1)
        args: list[list[Token]] = [[]]
        depth = 0
        for t in tokens[k + 2 : close]:
            if t.kind in (Kind.OPEN, Kind.ARRAY_OPEN):
                depth += 1
            elif t.kind in (Kind.CLOSE, Kind.ARRAY_CLOSE):
                depth -= 1
            elif t.kind is Kind.COMMA and depth == 0:
                args.append([])
                continue
            args[-1].append(t)
        declared = args[:-1] if which == "LAMBDA" else args[:-1:2]
        for arg in declared:
            if len(arg) == 1 and arg[0].kind is Kind.NAME:
                name = arg[0].text
                if name.lower().startswith(_PARAM):
                    name = name[len(_PARAM) :]
                names.add(name.casefold())
    return names


def _number(text: str) -> str:
    value = text_to_number(text)
    return text if value is None else number_to_text(value)


def _single(tokens: list[Token]) -> list[Token]:
    """``@x`` as ``_xlfn.SINGLE(x)``, the form a dynamic-array formula holds."""
    out = list(tokens)
    k = 0
    while k < len(out):
        tok = out[k]
        if tok.kind is not Kind.OPERATOR or tok.text != "@":
            k += 1
            continue
        j = k + 1
        if j < len(out) and out[j].kind is Kind.PREFIX:
            j += 1
        if j < len(out) and out[j].kind is Kind.FUNCTION:
            j = _matching_close(out, j + 1) + 1
        elif j < len(out) and out[j].kind is Kind.OPEN:
            j = _matching_close(out, j) + 1
        else:
            j += 1
            if j + 1 < len(out) and out[j].text == ":" and out[j + 1].kind is Kind.CELL:
                j += 2
        operand_first = out[k + 1]
        out[k + 1] = Token(operand_first.kind, operand_first.text, "", operand_first.pos)
        out[k : k + 1] = [
            Token(Kind.FUNCTION, "_xlfn.SINGLE", tok.space, tok.pos),
            Token(Kind.OPEN, "(", "", tok.pos),
        ]
        out.insert(j + 1, Token(Kind.CLOSE, ")", "", tok.pos))
        k += 2
    return out


def file_form(text: str, dynamic: bool = False) -> str:
    """The text Excel writes into a file for a formula typed as ``text`` (with or without ``=``).

    ``dynamic`` prints it for a dynamic-array formula, where ``@x`` is ``_xlfn.SINGLE(x)``;
    otherwise ``@`` is dropped, as a legacy formula intersects anyway.

    Raises:
        TokenizeError: The text holds a character no token can start with.
        ValueError: Unbalanced parentheses around a spill operator.
    """
    if text.startswith("="):
        text = text[1:]
    tokens = tokenize(text)
    if dynamic:
        tokens = _single(tokens)
    for tok in tokens:  # whitespace between tokens only: string literals stay as written
        tok.space = tok.space.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ")
    if tokens:
        tokens[-1].space = ""  # the end token carries the trailing whitespace, which Excel drops
    params = _parameters(tokens)
    known = _known()
    out: list[Token] = []
    previous: Token | None = None
    arrays = 0  # array constants open around the token
    for tok in tokens:
        kind, value, space = tok.kind, tok.text, tok.space
        if arrays or kind in (Kind.COMMA, Kind.SEMICOLON, Kind.ARRAY_CLOSE):
            space = ""  # Excel drops the whitespace in array constants and before a comma
        if kind is Kind.ARRAY_OPEN:
            arrays += 1
        elif kind is Kind.ARRAY_CLOSE:
            arrays -= 1
        if kind is Kind.FUNCTION:
            bare = _bare(value)
            upper = bare.upper()
            typed = _canonical_prefix(value[: len(value) - len(bare)])
            if upper in known and not (previous is not None and previous.kind is Kind.PREFIX):
                prefix = PREFIXES.get(upper, typed if typed.startswith("_xlfn.") else "")
                value = prefix + upper
        elif kind is Kind.PREFIX:
            value = _unquoted(value)
        elif kind is Kind.STRUCTURED:
            value = _this_row(value)
        elif (
            kind is Kind.OPERATOR
            and value == "+"
            and (previous is None or arrays)
            and (previous is None or previous.kind is not Kind.NUMBER)
        ):
            continue  # a leading "+", and a sign "+" in an array constant
        elif kind is Kind.NAME and value.casefold().removeprefix(_PARAM) in params:
            bare = value[len(_PARAM) :] if value.lower().startswith(_PARAM) else value
            value = _PARAM + bare
        elif kind in (Kind.CELL, Kind.COLUMNS, Kind.BOOL, Kind.ERROR):
            value = value.upper()
        elif kind is Kind.NUMBER:
            value = _number(value)
        elif kind is Kind.CLOSE and previous is not None and previous.kind is Kind.OPEN:
            space = ""
        elif kind is Kind.OPERATOR and value == "@":
            previous = tok
            continue
        elif kind is Kind.OPERATOR and value == "#" and out:
            start = len(out) - 1
            if out[start].kind is Kind.CLOSE:
                start = _matching_open(out, start)
                if start > 0 and out[start - 1].kind is Kind.FUNCTION:
                    start -= 1
            if start > 0 and out[start - 1].kind is Kind.PREFIX:
                start -= 1
            operand = out[start:]
            lead = operand[0].space
            operand[0] = Token(operand[0].kind, operand[0].text, "", operand[0].pos)
            del out[start:]
            out.append(Token(Kind.FUNCTION, "_xlfn.ANCHORARRAY", lead, tok.pos))
            out.append(Token(Kind.OPEN, "(", "", tok.pos))
            out.extend(operand)
            out.append(Token(Kind.CLOSE, ")", "", tok.pos))
            previous = tok
            continue
        out.append(Token(kind, value, space, tok.pos))
        previous = tok
    return _text(out)


_THIS_ROW = re.compile(r"^(?P<table>[^\[]*)\[@(?P<column>\[.*\]|[^\[\]]*)\]$", re.S)


def _this_row(text: str) -> str:
    """``Table1[@Price]`` as a file writes it, ``Table1[[#This Row],[Price]]``."""
    m = _THIS_ROW.match(text)
    if m is None:
        return text
    table, column = m["table"], m["column"]
    if not column:
        return f"{table}[#This Row]"
    if not column.startswith("["):
        column = f"[{column}]"
    return f"{table}[[#This Row],{column}]"


def _unquoted(prefix: str) -> str:
    """A quoted sheet prefix without quotes Excel does not need (``'Résumé'!`` as ``Résumé!``)."""
    if not (prefix.startswith("'") and prefix.endswith("'!")):
        return prefix
    inner = prefix[1:-2].replace("''", "'")
    parts = inner.split(":")
    if len(parts) > 2 or any(sheet_prefix(part) != part + "!" for part in parts):
        return prefix
    return inner + "!"


_KNOWN: set[str] = set()


def _known() -> set[str]:
    """The function names written in upper case: those the package knows, by bare name."""
    if not _KNOWN:
        from xlrecalc._functions import REGISTRY

        _KNOWN.update(PREFIXES)
        _KNOWN.update(_bare(name).upper() for name in REGISTRY)
    return _KNOWN


def _canonical_prefix(prefix: str) -> str:
    """``_xlfn.``, ``_xlfn._xlws.`` or ``_xludf.`` as Excel writes them."""
    return {"_XLFN.": "_xlfn.", "_XLFN._XLWS.": "_xlfn._xlws.", "_XLUDF.": "_xludf.", "": ""}.get(
        prefix.upper(), prefix
    )
