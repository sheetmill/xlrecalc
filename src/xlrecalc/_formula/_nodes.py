"""The syntax tree of a formula. Every node keeps the tokens it was parsed from.

``tokens(node)`` yields a node's tokens in source order, so printing the tree is joining
``space + text`` of each token: the printer cannot lose or reorder anything (ARCHITECTURE.md
§5.2). A reference rewrite replaces a token's text and leaves every other token as it was.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field

from xlrecalc._formula._tokens import Token

__all__ = [
    "ArrayConstant",
    "Binary",
    "Call",
    "Group",
    "Invoke",
    "Leaf",
    "Missing",
    "Node",
    "Postfix",
    "Prefix",
    "Reference",
    "Union",
    "print_tokens",
    "tokens",
]


@dataclass(slots=True)
class Leaf:
    """A number, string, logical or error literal."""

    token: Token


@dataclass(slots=True)
class Reference:
    """A cell, column range, row range, defined name or structured reference.

    ``prefix`` is the sheet (or book and sheet) token in front of it, if any. A name token
    holding an error literal, such as ``Sheet1!#REF!``, is also a reference.
    """

    token: Token
    prefix: Token | None = None


@dataclass(slots=True)
class Missing:
    """An omitted argument, as in ``IF(A1,,1)``. It has no tokens."""


@dataclass(slots=True)
class Prefix:
    """A prefix operator: ``-``, ``+`` or ``@``."""

    op: Token
    operand: Node


@dataclass(slots=True)
class Postfix:
    """A postfix operator: ``%`` or the spill operator ``#``."""

    operand: Node
    op: Token


@dataclass(slots=True)
class Binary:
    """A binary operator. ``op`` is ``None`` for intersection, written as whitespace."""

    left: Node
    op: Token | None
    right: Node


@dataclass(slots=True)
class Group:
    """An expression in parentheses."""

    open: Token
    inner: Node
    close: Token


@dataclass(slots=True)
class Union:
    """References joined by ``,`` inside parentheses: ``(A1,B2:C3)``."""

    items: list[Node]
    commas: list[Token]


@dataclass(slots=True)
class Call:
    """A function call. ``args`` holds a ``Missing`` for each omitted argument."""

    name: Token
    open: Token
    args: list[Node]
    commas: list[Token]
    close: Token
    prefix: Token | None = None


@dataclass(slots=True)
class Invoke:
    """A call of a function value, such as ``_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)``."""

    callee: Node
    open: Token
    args: list[Node]
    commas: list[Token]
    close: Token


@dataclass(slots=True)
class ArrayConstant:
    """An array constant ``{1,2;3,4}``. ``rows`` holds each element's tokens (sign and value)."""

    open: Token
    rows: list[list[list[Token]]]
    separators: list[Token] = field(default_factory=list)
    close: Token | None = None


Node = (
    Leaf
    | Reference
    | Missing
    | Prefix
    | Postfix
    | Binary
    | Group
    | Union
    | Call
    | Invoke
    | ArrayConstant
)


def tokens(node: Node) -> Iterator[Token]:
    """The node's tokens in source order."""
    stack: list[Node | Token] = [node]
    while stack:
        item = stack.pop()
        if isinstance(item, Token):
            yield item
            continue
        parts: list[Node | Token | None]
        match item:
            case Leaf(token):
                parts = [token]
            case Reference(token, prefix):
                parts = [prefix, token]
            case Missing():
                parts = []
            case Prefix(op, operand):
                parts = [op, operand]
            case Postfix(operand, op):
                parts = [operand, op]
            case Binary(left, op, right):
                parts = [left, op, right]
            case Group(open_, inner, close):
                parts = [open_, inner, close]
            case Union(items, commas):
                parts = _interleave(items, commas)
            case Call(name, open_, args, commas, close, prefix):
                parts = [prefix, name, open_, *_interleave(args, commas), close]
            case Invoke(callee, open_, args, commas, close):
                parts = [callee, open_, *_interleave(args, commas), close]
            case ArrayConstant(open_, rows, separators, close):
                elements = [t for row in rows for element in row for t in element]
                parts = [open_, *sorted([*elements, *separators], key=_pos), close]
        stack.extend(p for p in reversed(parts) if p is not None)


def _pos(token: Token) -> int:
    return token.pos


def _interleave(items: list[Node], separators: list[Token]) -> list[Node | Token | None]:
    out: list[Node | Token | None] = []
    for k, item in enumerate(items):
        if k:
            out.append(separators[k - 1])
        out.append(item)
    return out


def print_tokens(toks: Iterator[Token] | list[Token]) -> str:
    """Join tokens back into formula text."""
    return "".join(t.space + t.text for t in toks)
