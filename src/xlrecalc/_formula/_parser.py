"""Parse formula tokens into a syntax tree with Excel's operator precedence.

Precedence, tightest first: the spill operator ``#``; range ``:``; ``@``; intersection (a
space between two references); negation; ``%``; ``^``; ``*`` and ``/``; ``+`` and ``-``;
``&``; comparisons. Binary operators are left-associative (``=2^3^2`` is 64) and negation
binds tighter than ``^`` (``=-2^2`` is 4). A ``,`` inside parentheses that are not a call's
is the union operator.
"""

from __future__ import annotations

from dataclasses import dataclass

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
    Union,
    print_tokens,
    tokens,
)
from xlrecalc._formula._tokens import Kind, Token, tokenize

__all__ = ["Formula", "FormulaSyntaxError", "parse"]


class FormulaSyntaxError(ValueError):
    """The formula is not well formed."""

    def __init__(self, message: str, pos: int) -> None:
        super().__init__(f"{message} at position {pos}")
        self.pos = pos


@dataclass(slots=True)
class Formula:
    """A parsed formula: its tree and the end token (which holds trailing whitespace)."""

    root: Node
    end: Token

    def text(self) -> str:
        """The formula's text, exactly as parsed unless a token was rewritten."""
        return print_tokens([*tokens(self.root), self.end])


# Binding powers: (left, right) for infix operators.
_INFIX = {
    "=": (10, 11),
    "<>": (10, 11),
    "<": (10, 11),
    ">": (10, 11),
    "<=": (10, 11),
    ">=": (10, 11),
    "&": (20, 21),
    "+": (30, 31),
    "-": (30, 31),
    "*": (40, 41),
    "/": (40, 41),
    "^": (50, 51),
    ":": (90, 91),
}
_PERCENT = 55
_NEGATION = 60
_INTERSECTION = (70, 71)
_AT = 80
_SPILL = 100
_REFERENCE_START = {
    Kind.CELL,
    Kind.COLUMNS,
    Kind.ROWS,
    Kind.NAME,
    Kind.STRUCTURED,
    Kind.PREFIX,
    Kind.FUNCTION,
    Kind.OPEN,
}
_ARRAY_VALUE = {Kind.NUMBER, Kind.STRING, Kind.BOOL, Kind.ERROR}
# Nesting of parentheses, calls and prefix operators. Excel allows 64 levels of nested
# functions; the parser and the code generator recurse once per level.
MAX_DEPTH = 200


def parse(text: str) -> Formula:
    """Parse a formula in file form (no leading ``=``).

    Raises:
        TokenizeError: A character starts no token.
        FormulaSyntaxError: The tokens do not form a formula.
    """
    parser = _Parser(tokenize(text))
    root = parser.expression(0)
    end = parser.take()
    if end.kind is not Kind.END:
        raise FormulaSyntaxError(f"unexpected {end.text!r}", end.pos)
    return Formula(root, end)


class _Parser:
    __slots__ = ("depth", "i", "toks")

    def __init__(self, toks: list[Token]) -> None:
        self.toks = toks
        self.i = 0
        self.depth = 0

    def peek(self) -> Token:
        return self.toks[self.i]

    def take(self) -> Token:
        tok = self.toks[self.i]
        if tok.kind is not Kind.END:
            self.i += 1
        return tok

    def expect(self, kind: Kind) -> Token:
        tok = self.take()
        if tok.kind is not kind:
            raise FormulaSyntaxError(f"expected {kind.value}, found {tok.text!r}", tok.pos)
        return tok

    def expression(self, min_bp: int) -> Node:
        self.depth += 1
        if self.depth > MAX_DEPTH:
            raise FormulaSyntaxError("nested too deeply", self.peek().pos)
        try:
            return self.operators(self.prefix(), min_bp)
        finally:
            self.depth -= 1

    def operators(self, left: Node, min_bp: int) -> Node:
        while True:
            tok = self.peek()
            if tok.kind is Kind.OPEN and not tok.space and isinstance(left, (Call, Invoke)):
                left = self.invoke(left, self.take())
                continue
            if tok.kind is Kind.OPERATOR:
                if tok.text == "#":
                    if min_bp > _SPILL:
                        break
                    left = Postfix(left, self.take())
                    continue
                if tok.text == "%":
                    if min_bp > _PERCENT:
                        break
                    left = Postfix(left, self.take())
                    continue
                if tok.text not in _INFIX:
                    break
                lbp, rbp = _INFIX[tok.text]
                if lbp < min_bp:
                    break
                op = self.take()
                left = Binary(left, op, self.expression(rbp))
                continue
            if tok.space and tok.kind in _REFERENCE_START and _INTERSECTION[0] >= min_bp:
                left = Binary(left, None, self.expression(_INTERSECTION[1]))
                continue
            break
        return left

    def prefix(self) -> Node:
        tok = self.take()
        kind = tok.kind
        if kind is Kind.OPERATOR and tok.text in ("-", "+"):
            return Prefix(tok, self.expression(_NEGATION))
        if kind is Kind.OPERATOR and tok.text == "@":
            return Prefix(tok, self.expression(_AT))
        if kind in (Kind.NUMBER, Kind.STRING, Kind.BOOL, Kind.ERROR):
            return Leaf(tok)
        if kind in (Kind.CELL, Kind.COLUMNS, Kind.ROWS, Kind.NAME, Kind.STRUCTURED):
            return Reference(tok)
        if kind is Kind.PREFIX:
            target = self.take()
            if target.space:
                raise FormulaSyntaxError("space after a sheet name", target.pos)
            if target.kind is Kind.FUNCTION:
                return self.call(target, tok)
            if target.kind in (Kind.CELL, Kind.COLUMNS, Kind.ROWS, Kind.NAME, Kind.ERROR):
                return Reference(target, tok)
            raise FormulaSyntaxError(f"expected a reference after {tok.text!r}", target.pos)
        if kind is Kind.FUNCTION:
            return self.call(tok, None)
        if kind is Kind.OPEN:
            return self.group(tok)
        if kind is Kind.ARRAY_OPEN:
            return self.array(tok)
        raise FormulaSyntaxError(f"unexpected {tok.text or 'end of formula'!r}", tok.pos)

    def call(self, name: Token, prefix: Token | None, opened: Token | None = None) -> Call:
        open_ = opened if opened is not None else self.expect(Kind.OPEN)
        args: list[Node] = []
        commas: list[Token] = []
        if self.peek().kind is Kind.CLOSE:
            return Call(name, open_, args, commas, self.take(), prefix)
        while True:
            if self.peek().kind in (Kind.COMMA, Kind.CLOSE):
                args.append(Missing())
            else:
                args.append(self.expression(0))
            tok = self.take()
            if tok.kind is Kind.CLOSE:
                return Call(name, open_, args, commas, tok, prefix)
            if tok.kind is not Kind.COMMA:
                raise FormulaSyntaxError(f"expected , or ), found {tok.text!r}", tok.pos)
            commas.append(tok)

    def invoke(self, callee: Node, open_: Token) -> Invoke:
        call = self.call(open_, None, opened=open_)
        return Invoke(callee, call.open, call.args, call.commas, call.close)

    def group(self, open_: Token) -> Group:
        first = self.expression(0)
        if self.peek().kind is Kind.COMMA:
            items, commas = [first], []
            while self.peek().kind is Kind.COMMA:
                commas.append(self.take())
                items.append(self.expression(0))
            inner: Node = Union(items, commas)
        else:
            inner = first
        return Group(open_, inner, self.expect(Kind.CLOSE))

    def array(self, open_: Token) -> ArrayConstant:
        rows: list[list[list[Token]]] = [[]]
        separators: list[Token] = []
        while True:
            element: list[Token] = []
            tok = self.take()
            if tok.kind is Kind.OPERATOR and tok.text in ("-", "+"):
                element.append(tok)
                tok = self.take()
            if tok.kind not in _ARRAY_VALUE or (element and tok.kind is not Kind.NUMBER):
                raise FormulaSyntaxError(f"not an array constant value: {tok.text!r}", tok.pos)
            element.append(tok)
            rows[-1].append(element)
            sep = self.take()
            if sep.kind is Kind.ARRAY_CLOSE:
                return ArrayConstant(open_, rows, separators, sep)
            if sep.kind is Kind.SEMICOLON:
                rows.append([])
            elif sep.kind is not Kind.COMMA:
                raise FormulaSyntaxError(f"expected , ; or }}, found {sep.text!r}", sep.pos)
            separators.append(sep)
