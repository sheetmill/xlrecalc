"""Compile a formula's tree into a Python function ``fn(ctx, r, c)`` (ARCHITECTURE.md §5.4).

The function is straight-line code: every operation is one statement assigning a temporary
(``t3 = ops.add(t1, t2)``), so the source nests at most two levels whatever the formula's depth
or length, and CPython's own limits on nested parentheses never apply.

**No text from the workbook enters the generated source.** Literals, sheet names, defined
names, structured references and array constants go into the constants table ``K``; functions
go into the table ``F``, resolved from the registry by name, so the source refers to them only
by index. What remains is names chosen here (``ALLOWED_NAMES``, and temporaries ``t0``,
``t1``, ...), integers formatted from Python ``int`` values, and punctuation.
``tests/test_codegen.py`` checks every generated source against that rule with Python's own
tokenizer.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Any

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
)
from xlrecalc._formula._parser import Formula
from xlrecalc._formula._template import Axis, cell_axes, line_axes
from xlrecalc._formula._tokens import Kind, Token
from xlrecalc._numtext import literal_number
from xlrecalc._quirks import Quirks
from xlrecalc._values import ErrorValue

__all__ = ["ALLOWED_NAMES", "MISSING", "Compiled", "RefSpec", "compile_formula"]

_LIFTED = frozenset({"number", "text", "logical", "value", "branch"})
# Functions whose later arguments are evaluated only when taken, and their operator helpers.
_LAZY = {
    "IF": "lazy_if",
    "IFERROR": "lazy_iferror",
    "_XLFN.IFNA": "lazy_ifna",
    "CHOOSE": "lazy_choose",
}
_SCALAR = frozenset({"number", "text", "logical", "value"})

MAX_ROW = 1048575  # 0-based limits of a sheet
MAX_COL = 16383

_BINARY = {
    "+": "add",
    "-": "sub",
    "*": "mul",
    "/": "div",
    "^": "pow",
    "&": "concat",
    "=": "eq",
    "<>": "ne",
    "<": "lt",
    ">": "gt",
    "<=": "le",
    ">=": "ge",
    ":": "range",
}
_PREFIX = {"-": "neg", "+": "pos", "@": "single"}
_CONTEXT = (
    "arr",
    "cell",
    "cref",
    "box",
    "columns",
    "rows",
    "name",
    "structured",
    "spill",
    "unsupported",
)
ALLOWED_NAMES = frozenset(
    {
        "def",
        "return",
        "_formula",
        "ctx",
        "r",
        "c",
        "K",
        "F",
        "ops",
        "MISSING",
        "None",
        "intersect",
        "add_top",
        "sub_top",
        "union",
        "percent",
        "invoke",
        *_LAZY.values(),
        *_BINARY.values(),
        *_PREFIX.values(),
        *_CONTEXT,
    }
)


class _MissingType:
    """An omitted function argument."""

    __slots__ = ()

    def __repr__(self) -> str:
        return "MISSING"


MISSING = _MissingType()


@dataclass(frozen=True, slots=True)
class RefSpec:
    """A reference a template makes, for its dependencies (ARCHITECTURE.md §6).

    ``kind`` is ``box`` (a cell is a one-cell box; whole columns and rows are boxes over every
    row or column) or ``name`` (a defined name, resolved by the engine). ``sheet`` is the
    prefix's sheet text or ``None`` for the formula's own sheet; the axes are relative to the
    formula's cell unless absolute.
    """

    kind: str
    sheet: str | None
    r1: Axis | None = None
    c1: Axis | None = None
    r2: Axis | None = None
    c2: Axis | None = None
    name: str | None = None


@dataclass(frozen=True, slots=True)
class Compiled:
    """A compiled formula: its source, constants, functions, references and the callable."""

    source: str
    constants: tuple[Any, ...]
    functions: tuple[Callable[..., Any], ...]
    fn: Callable[[Any, int, int], Any]
    refs: tuple[RefSpec, ...] = ()
    calls: frozenset[str] = frozenset()
    features: frozenset[str] = frozenset()
    """Features the formula uses that are not implemented yet."""
    intersects: bool = False
    """Whether, as a legacy formula, it takes an implicit intersection somewhere: an operator or
    a one-value argument given a range outside an array argument (§7.7)."""


def compile_formula(
    formula: Formula,
    row: int,
    col: int,
    functions: Mapping[str, Callable[..., Any]],
    ops: Any,
    quirks: Quirks = Quirks.EXCEL,
    array_mode: bool = False,
) -> Compiled:
    """Compile a formula in the cell (row, col), 0-based.

    Args:
        formula: The parsed formula.
        row: The formula's row.
        col: The formula's column.
        functions: Registry: upper-case function name (file form, ``_xlfn.`` included) to the
            callable, called as ``fn(ctx, *args)``. A name not in it calls
            ``ctx.unsupported(name, *args)``.
        ops: Object with the operator helpers (``add``, ``neg``, ``range``, ...).
        quirks: The quirks in force: number literals keep their first 15 significant digits
            under ``fifteen_digit_precision``, as Excel computes them (capture ``literals``).
        array_mode: Compile an array formula (a legacy array formula, or a dynamic-array
            formula): ranges are arrays everywhere, and the result may be an array.
    """
    gen = _Generator(row, col, functions, quirks)
    gen.arrays = 1 if array_mode else 0
    gen.array_formula = array_mode
    result = gen.expr(formula.root, top=True)
    if array_mode:
        result = gen.emit(f"ctx.arr({result})")
    body = "".join(f"    {line}\n" for line in gen.lines)
    source = "".join(gen.branches) + f"def _formula(ctx, r, c):\n{body}    return {result}\n"
    namespace: dict[str, Any] = {
        "__builtins__": {},
        "K": tuple(gen.constants),
        "F": tuple(gen.functions),
        "ops": ops,
        "MISSING": MISSING,
    }
    exec(compile(source, "<formula>", "exec"), namespace)
    fn = namespace["_formula"]
    return Compiled(
        source,
        tuple(gen.constants),
        tuple(gen.functions),
        fn,
        tuple(gen.refs),
        frozenset(gen.calls),
        frozenset(gen.features),
        gen.intersects,
    )


class _Generator:
    def __init__(
        self,
        row: int,
        col: int,
        registry: Mapping[str, Callable[..., Any]],
        quirks: Quirks = Quirks.EXCEL,
    ) -> None:
        self.quirks = quirks
        self.row = row
        self.col = col
        self.registry = registry
        self.constants: list[Any] = []
        self.constant_index: dict[tuple[type, Any], int] = {}
        self.functions: list[Callable[..., Any]] = []
        self.function_index: dict[int, int] = {}
        self.lines: list[str] = []
        self.refs: list[RefSpec] = []
        self.calls: set[str] = set()
        self.features: set[str] = set()
        self.arrays = 0  # array contexts around the node being compiled (§7.7)
        self.array_formula = False  # the formula itself is an array formula
        self.branches: list[str] = []  # functions for the branches of IF and friends
        self.intersects = False

    def operand(self, node: Node, code: str) -> str:
        """An operator's operand: an array in an array context, else implicitly intersected."""
        if self.arrays:
            return code if isinstance(node, Leaf) else self.emit(f"ctx.arr({code})")
        if _is_range(node, ranges_only=True):
            self.intersects = True
        return code

    def emit(self, expression: str) -> str:
        """Assign an expression to a new temporary and return its name."""
        name = f"t{len(self.lines):d}"
        self.lines.append(f"{name} = {expression}")
        return name

    def constant(self, value: Any) -> str:
        key = (type(value), value)
        try:
            index = self.constant_index.get(key)
        except TypeError:  # unhashable: never shared
            index = None
            key = (object, object())
        if index is None:
            index = len(self.constants)
            self.constants.append(value)
            self.constant_index[key] = index
        return f"K[{index:d}]"

    def function(self, fn: Callable[..., Any]) -> str:
        index = self.function_index.get(id(fn))
        if index is None:
            index = len(self.functions)
            self.functions.append(fn)
            self.function_index[id(fn)] = index
        return f"F[{index:d}]"

    def sheet(self, prefix: Token | None) -> str:
        return "None" if prefix is None else self.constant(sheet_text(prefix))

    def axis(self, axis: Axis, var: str) -> str:
        if axis.absolute:
            return f"{axis.value:d}"
        if axis.value == 0:
            return var
        sign = "+" if axis.value > 0 else "-"
        return f"({var} {sign} {abs(axis.value):d})"

    def expr(self, node: Node, top: bool = False) -> str:
        """Code for a node; ``top`` marks the formula's outermost operation (§4.4 snapping)."""
        if isinstance(node, Binary) and _boxable(node):
            return self.box(node)
        match node:
            case Leaf(token):
                return self.constant(_literal(token, self.quirks))
            case Reference(token, prefix):
                return self.reference(token, prefix)
            case Missing():
                return "MISSING"
            case Prefix(op, operand):
                code = self.expr(operand)
                if op.text != "@":  # @ takes the implicit intersection of the reference itself
                    code = self.operand(operand, code)
                return self.emit(f"ops.{_PREFIX[op.text]}({code})")
            case Postfix(operand, op):
                if op.text == "%":
                    code = self.operand(operand, self.expr(operand))
                    return self.emit(f"ops.percent({code})")
                return self.emit(f"ctx.spill({self.expr(operand)})")
            case Binary():
                return self.chain(node, top)
            case Group(_, inner, _):
                return self.expr(inner)
            case Union(items, _):
                return self.emit(f"ops.union({', '.join(self.expr(i) for i in items)})")
            case Call(name, _, args, _, _, prefix):
                return self.call(name, prefix, args)
            case Invoke(callee, _, args):
                argv = "".join(f", {self.expr(a)}" for a in args)
                return self.emit(f"ops.invoke(ctx, {self.expr(callee)}{argv})")
            case ArrayConstant(_, rows):
                return self.constant(
                    tuple(tuple(_element(e, self.quirks) for e in row) for row in rows)
                )
        raise AssertionError(node)

    def chain(self, node: Binary, top: bool = False) -> str:
        """A left-deep run of binary operators (``A1+A2+...``), walked without recursion.

        The formula's own last operation, when it is ``+`` or ``-`` outside any parentheses,
        goes through ``add_top``/``sub_top``, which apply cancellation snapping.
        """
        spine: list[Binary] = []
        current: Node = node
        while isinstance(current, Binary) and not _boxable(current):
            spine.append(current)
            current = current.left
        first_range = bool(spine) and spine[-1].op is not None and spine[-1].op.text == ":"
        # The operands of ":" are references (A1:INDEX(...)): a cell compiles as one.
        result = self.argument(current) if first_range else self.expr(current)
        for item in reversed(spine):
            is_range = item.op is not None and item.op.text == ":"
            right = self.argument(item.right) if is_range else self.expr(item.right)
            name = "intersect" if item.op is None else _BINARY[item.op.text]
            if name not in ("intersect", "range"):
                if item.left is current:
                    result = self.operand(current, result)
                right = self.operand(item.right, right)
            if top and item is node and name in ("add", "sub"):
                name += "_top"
            result = self.emit(f"ops.{name}({result}, {right})")
        return result

    def box(self, node: Binary) -> str:
        """``A1:B2`` between two cell references: one box, on the left one's sheet."""
        left, right = node.left, node.right
        assert isinstance(left, Reference)
        assert isinstance(right, Reference)
        sheet = self.sheet(left.prefix)
        r1, c1 = cell_axes(left.token.text, self.row, self.col)
        r2, c2 = cell_axes(right.token.text, self.row, self.col)
        where = sheet_text(left.prefix) if left.prefix is not None else None
        self.refs.append(RefSpec("box", where, r1, c1, r2, c2))
        rows = f"{self.axis(r1, 'r')}, {self.axis(c1, 'c')}, {self.axis(r2, 'r')}"
        return self.emit(f"ctx.box({sheet}, {rows}, {self.axis(c2, 'c')})")

    def reference(self, token: Token, prefix: Token | None) -> str:
        sheet = self.sheet(prefix)
        where = sheet_text(prefix) if prefix is not None else None
        if token.kind is Kind.CELL:
            r, c = cell_axes(token.text, self.row, self.col)
            self.refs.append(RefSpec("box", where, r, c, r, c))
            return self.emit(f"ctx.cell({sheet}, {self.axis(r, 'r')}, {self.axis(c, 'c')})")
        if token.kind in (Kind.COLUMNS, Kind.ROWS):
            var = "r" if token.kind is Kind.ROWS else "c"
            first, second = line_axes(token, self.row, self.col)
            method = token.kind.value
            zero = Axis(0, True)
            if token.kind is Kind.ROWS:
                spec = RefSpec("box", where, first, zero, second, Axis(MAX_COL, True))
            else:
                spec = RefSpec("box", where, zero, first, Axis(MAX_ROW, True), second)
            self.refs.append(spec)
            return self.emit(
                f"ctx.{method}({sheet}, {self.axis(first, var)}, {self.axis(second, var)})"
            )
        if token.kind is Kind.STRUCTURED:
            return self.emit(f"ctx.structured({self.constant(token.text)})")
        if token.kind is Kind.ERROR:
            return self.constant(ErrorValue(token.text))
        self.refs.append(RefSpec("name", where, name=token.text))
        return self.emit(f"ctx.name({sheet}, {self.constant(token.text)})")

    def branch(self, node: Node, kind: str) -> str:
        """A lazy argument, compiled as a function of its own (``b0(ctx, r, c)``).

        The call evaluates it only if it is taken: ``IF(FALSE,A1,5)`` never reads A1.
        """
        saved = self.lines
        self.lines = []
        code = self.argument(node)
        if self.arrays and kind in _LIFTED and not isinstance(node, Leaf):
            code = self.emit(f"ctx.arr({code})")
        body = "".join(f"    {line}\n" for line in self.lines)
        self.lines = saved
        name = f"b{len(self.branches):d}"
        self.branches.append(f"def {name}(ctx, r, c):\n{body}    return {code}\n")
        return name

    def lazy(self, fn: Callable[..., Any], upper: str, args: list[Node]) -> str:
        """``IF``, ``IFERROR``, ``IFNA`` and ``CHOOSE`` (``ops.lazy_*``).

        The first argument is evaluated, the others only when taken.
        """
        kinds: tuple[str, ...] = getattr(fn, "kinds", ())
        first = self.argument(args[0]) if args else "MISSING"
        if args and self.arrays and not isinstance(args[0], Leaf):
            first = self.emit(f"ctx.arr({first})")
        rest = [self.branch(a, kinds[min(k + 1, len(kinds) - 1)]) for k, a in enumerate(args[1:])]
        helper = _LAZY[upper]
        names = "".join(f", {b}" for b in rest)
        return self.emit(f"ops.{helper}({self.function(fn)}, {first}, r, c{names})")

    def call(self, name: Token, prefix: Token | None, args: list[Node]) -> str:
        fn = self.registry.get(name.text.upper()) if prefix is None else None
        if fn is not None and name.text.upper() in _LAZY and args:
            self.calls.add(name.text.upper())
            return self.lazy(fn, name.text.upper(), args)
        if self.arrays:
            fn = getattr(fn, "in_array", fn)
        kinds: tuple[str, ...] = getattr(fn, "kinds", ())
        cycle: int = getattr(fn, "cycle", 1)
        parts = []
        for i, arg in enumerate(args):
            kind = kind_at(kinds, i, cycle) if kinds else "raw"
            if kind == "matrix":  # follows the formula's own mode, not an array argument's
                saved, self.arrays = self.arrays, int(self.array_formula)
                code = self.argument(arg)
                if self.arrays and not isinstance(arg, Leaf):
                    code = self.emit(f"ctx.arr({code})")
                elif not self.arrays and _is_range(arg, ranges_only=True):
                    self.intersects = True
                self.arrays = saved
                parts.append(code)
                continue
            if kind == "key":  # follows the formula's own mode, and is one value outside it
                saved, self.arrays = self.arrays, int(self.array_formula)
                code = self.argument(arg)
                if self.arrays:
                    code = code if isinstance(arg, Leaf) else self.emit(f"ctx.arr({code})")
                else:
                    self.intersects = self.intersects or _is_range(arg, ranges_only=True)
                    code = code if isinstance(arg, Leaf) else self.emit(f"ops.single({code})")
                self.arrays = saved
                parts.append(code)
                continue
            if kind == "array":
                self.arrays += 1
                parts.append(self.argument(arg))
                self.arrays -= 1
                continue
            code = self.argument(arg)
            if self.arrays and kind in _LIFTED and not isinstance(arg, Leaf):
                code = self.emit(f"ctx.arr({code})")
            elif not self.arrays and kind in _SCALAR and _is_range(arg, ranges_only=True):
                self.intersects = True
            parts.append(code)
        argv = "".join(f", {code}" for code in parts)
        self.calls.add(name.text.upper() if prefix is None else prefix.text + name.text)
        if fn is None:
            label = name.text if prefix is None else prefix.text + name.text
            return self.emit(f"ctx.unsupported({self.constant(label)}{argv})")
        return self.emit(f"{self.function(fn)}(ctx{argv})")

    def argument(self, node: Node) -> str:
        """A function argument: a single cell is passed as a reference, not its value."""
        if isinstance(node, Reference) and node.token.kind is Kind.CELL:
            sheet = self.sheet(node.prefix)
            r, c = cell_axes(node.token.text, self.row, self.col)
            where = sheet_text(node.prefix) if node.prefix is not None else None
            self.refs.append(RefSpec("box", where, r, c, r, c))
            return self.emit(f"ctx.cref({sheet}, {self.axis(r, 'r')}, {self.axis(c, 'c')})")
        return self.expr(node)


def kind_at(kinds: tuple[str, ...], i: int, cycle: int = 1) -> str:
    """The kind of argument ``i``: past the declared ones, the last ``cycle`` kinds repeat."""
    n = len(kinds)
    return kinds[i] if i < n else kinds[n - cycle + (i - n) % cycle]


def _is_range(node: Node, ranges_only: bool = False) -> bool:
    """A node that denotes more than one cell or value.

    A range, a union, a name (which may be a range) or, unless ``ranges_only``, an array
    constant.
    """
    while isinstance(node, Group):
        node = node.inner
    if isinstance(node, Binary):
        return node.op is not None and node.op.text == ":"
    if isinstance(node, Reference):
        return node.token.kind in (Kind.COLUMNS, Kind.ROWS, Kind.NAME)
    if isinstance(node, ArrayConstant):
        return not ranges_only
    return isinstance(node, Union)


def _boxable(node: Binary) -> bool:
    left, right = node.left, node.right
    return (
        node.op is not None
        and node.op.text == ":"
        and isinstance(left, Reference)
        and isinstance(right, Reference)
        and left.token.kind is Kind.CELL
        and right.token.kind is Kind.CELL
        and right.prefix is None
    )


def sheet_text(prefix: Token) -> str:
    """The sheet part of a prefix token without the ``!`` and without quoting.

    ``'It''s'!`` gives ``It's``; ``Sheet1:Sheet3!`` gives ``Sheet1:Sheet3``; a book index such
    as ``[1]`` is kept.
    """
    text = prefix.text[:-1]
    if text.startswith("'") and text.endswith("'"):
        return text[1:-1].replace("''", "'")
    return text


def _literal(token: Token, quirks: Quirks = Quirks.EXCEL) -> Any:
    kind = token.kind
    if kind is Kind.NUMBER:
        return literal_number(token.text, quirks)
    if kind is Kind.STRING:
        return token.text[1:-1].replace('""', '"')
    if kind is Kind.BOOL:
        return token.text.upper() == "TRUE"
    return ErrorValue(token.text)


def _element(tokens: list[Token], quirks: Quirks = Quirks.EXCEL) -> Any:
    value = _literal(tokens[-1], quirks)
    if len(tokens) == 2 and tokens[0].text == "-":
        return -value
    return value
