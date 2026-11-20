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
    "_XLFN.IFS": "lazy_ifs",
    "_XLFN.SWITCH": "lazy_switch",
}
_SCALAR = frozenset({"number", "text", "logical", "value"})
_PARAM = "_xlpm."  # a LET or LAMBDA name, as a file writes it
_OPTIONAL = "_xlop."  # an optional LAMBDA parameter in its list (``[y]`` typed)
# The generator versions of the lazy helpers, for a branch that calls a function value.
_GLAZY = {helper: "g" + helper for helper in _LAZY.values()}

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
    "arr_padded",
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
        "E",
        "scope",
        "force",
        "bind",
        "closure",
        "request",
        "omitted",
        "gforce",
        "yield",
        "from",
        *_LAZY.values(),
        *_GLAZY.values(),
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
    forms: frozenset[str] = frozenset()
    """Forms in its text that are not computed when evaluation reaches them (a union of ranges, a
    structured reference, ``CELL`` given a formatting kind as text), named for ``Formula``."""
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
    as_reference: bool = False,
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
        as_reference: Compile a defined name's definition: one cell gives a reference, as a
            function argument does, not its value (capture ``names``).
    """
    gen = _Generator(row, col, functions, quirks)
    gen.arrays = 1 if array_mode else 0
    gen.array_formula = array_mode
    if as_reference:
        root = formula.root
        while isinstance(root, Group):
            root = root.inner
        result = gen.argument(root)
    else:
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
        intersects=gen.intersects,
        forms=frozenset(gen.forms),
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
        self.forms: set[str] = set()
        self.arrays = 0  # array contexts around the node being compiled (§7.7)
        # Only operators lie between the node and a function taught to read a padded array
        # (``padded`` on the function): a range made into an array here may be padded (§7.7).
        self.padded = False
        self.array_formula = False  # the formula itself is an array formula
        self.branches: list[str] = []  # functions for the branches of IF and friends
        self.intersects = False
        self.branch_yields = False  # the last branch compiled calls a function value
        self.scopes: list[dict[str, int]] = []  # LET names in scope, innermost last
        self.values: dict[str, Node] = {}  # what each LET name in scope is bound to
        self.env = "None"  # the variable holding the innermost LET's environment
        # Per scope: names whose binding is a generator (it calls a function value), and
        # whether the scope is a LAMBDA's parameters.
        self.generators: list[set[str]] = []
        self.parameters: list[bool] = []
        self.bodies = 0  # LAMBDA bodies around the node being compiled
        self.let_names = 0  # LET names the innermost LAMBDA body binds (its cost a call)
        self.yields = False  # the function being generated calls a function value

    def ranged(self, node: Node) -> bool:
        """A range a legacy formula intersects; a ``LET`` name is what it is bound to."""
        seen = 0
        while (
            isinstance(node, Reference)
            and node.prefix is None
            and node.token.kind is Kind.NAME
            and node.token.text.casefold() in self.values
            and seen < 64  # a name bound to the name before it, at most so deep
        ):
            node, seen = self.values[node.token.text.casefold()], seen + 1
        return _is_range(node, ranges_only=True)

    def operand(self, node: Node, code: str) -> str:
        """An operator's operand: an array in an array context, else implicitly intersected."""
        if self.arrays:
            if isinstance(node, Leaf):
                return code
            return self.emit(f"ctx.{'arr_padded' if self.padded else 'arr'}({code})")
        if self.ranged(node):
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

    def sheet(self, prefix: Token | None, *, name: bool = False) -> str:
        if prefix is None:
            return "None"
        text = sheet_text(prefix)
        if reaches_out(text, name=name):  # kept, as a function not implemented is (§7.6)
            self.features.add(EXTERNAL)
        return self.constant(text)

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
                self.forms.add("spill reference")
                return self.emit(f"ctx.spill({self.opaque(operand)})")
            case Binary():
                return self.chain(node, top)
            case Group(_, inner, _):
                return self.expr(inner)
            case Union(items, _):
                self.forms.add("a union of ranges")
                return self.emit(f"ops.union({', '.join(self.opaque(i) for i in items)})")
            case Call(name, _, args, _, _, prefix):
                # An element-wise function passes the outer mark on to its lifted argument.
                fn = self.registry.get(name.text.upper()) if prefix is None else None
                keep = self.padded and bool(getattr(fn, "elementwise", False))
                saved, self.padded = self.padded, keep
                try:
                    return self.call(name, prefix, args)
                finally:
                    self.padded = saved
            case Invoke(callee, _, args):
                if _calls_a_value(callee, self.registry):  # SUM(1)(2): Excel refuses the file
                    self.features.add("a call of a function's result that is no LAMBDA")
                saved, self.padded = self.padded, False
                try:
                    return self.invoke(self.expr(callee), args)
                finally:
                    self.padded = saved
            case ArrayConstant(_, rows):
                return self.constant(
                    tuple(tuple(_element(e, self.quirks) for e in row) for row in rows)
                )
        raise AssertionError(node)

    def opaque(self, node: Node) -> str:
        """Code for a node read by something other than an operator: nothing in it is padded."""
        saved, self.padded = self.padded, False
        try:
            return self.expr(node)
        finally:
            self.padded = saved

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
        first_range = bool(spine) and (spine[-1].op is None or spine[-1].op.text == ":")
        # The operands of ":" and of the intersection (a space) are references (A1:INDEX(...),
        # A1 A2): a cell compiles as one.
        result = self.argument(current) if first_range else self.expr(current)
        for item in reversed(spine):
            is_range = item.op is None or item.op.text == ":"
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
        sheet = self.sheet(prefix, name=token.kind is Kind.NAME)
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
            self.forms.add("structured reference")
            return self.emit(f"ctx.structured({self.constant(token.text)})")
        if token.kind is Kind.ERROR:
            return self.constant(ErrorValue(token.text))
        if prefix is None and token.text[:6].casefold() == _PARAM:
            return self.bound(token.text)
        self.refs.append(RefSpec("name", where, name=token.text))
        return self.emit(f"ctx.name({sheet}, {self.constant(token.text)})")

    def bound(self, text: str) -> str:
        """A ``LET`` name: its binding, evaluated once when first read (capture ``dyn_let``).

        Resolved here, innermost scope first, into a depth and an index in the environment; a
        name not bound at this point (bound later, or nowhere) is ``#NAME?``, as in Excel.
        """
        found = self.lookup(text)
        if found is None:
            return self.constant(ErrorValue.NAME)
        depth, index, generator = found
        if generator:  # its binding calls a function value: evaluated on the call stack
            self.yields = True
            return self.emit(f"yield from ops.gforce({self.env}, {depth:d}, {index:d})")
        return self.emit(f"ops.force({self.env}, {depth:d}, {index:d})")

    def lookup(self, text: str) -> tuple[int, int, bool] | None:
        """A name in scope: its depth, its index and whether its binding is a generator."""
        folded = text.casefold()
        for depth, names in enumerate(reversed(self.scopes)):
            index = names.get(folded)
            if index is not None:
                return depth, index, folded in self.generators[-1 - depth]
        return None

    def let(self, args: list[Node]) -> str:
        """``LET``: each binding a function of the environment, evaluated only when read.

        With no calculation after the bindings the last name is the calculation, as Excel
        writes it back. A binding sees the names before it, the calculation all of them.
        """
        self.calls.add("_XLFN.LET")
        self.let_names += len(args) // 2
        pairs = args if len(args) % 2 == 0 else args[:-1]
        body = pairs[-2] if len(args) % 2 == 0 else args[-1]
        names: dict[str, int] = {}
        generators: set[str] = set()
        self.scopes.append(names)
        self.generators.append(generators)
        self.parameters.append(False)
        thunks = []
        outer: list[tuple[str, Node | None]] = []  # what the names shadowed, restored after
        for k in range(0, len(pairs), 2):
            name = pairs[k]
            assert isinstance(name, Reference)  # refused in a file otherwise (_load)
            thunk, generator = self.function_of_env(pairs[k + 1])
            thunks.append(thunk)
            names[name.token.text.casefold()] = k // 2
            if generator:
                generators.add(name.token.text.casefold())
            outer.append((name.token.text.casefold(), self.values.get(name.token.text.casefold())))
            self.values[name.token.text.casefold()] = pairs[k + 1]
        listed = "".join(f"{t}, " for t in thunks)
        saved = self.env
        self.env = self.emit(f"ops.scope(ctx, r, c, {saved}, ({listed}))")
        result = self.argument(body)
        self.env = saved
        self.scopes.pop()
        self.generators.pop()
        self.parameters.pop()
        for folded, value in reversed(outer):
            if value is None:
                self.values.pop(folded, None)
            else:
                self.values[folded] = value
        return result

    def function_of_env(self, node: Node) -> tuple[str, bool]:
        """A top-level function ``bK(ctx, r, c, E)`` of a value, ``E`` the environment.

        Also whether it is a generator: inside a ``LAMBDA`` body, a function that calls a
        function value yields the call to the call stack (``Ops.invoke``) instead of making it.
        """
        saved_lines, saved_env, saved_yields = self.lines, self.env, self.yields
        self.lines, self.env, self.yields = [], "E", False
        code = self.argument(node)
        body = "".join(f"    {line}\n" for line in self.lines)
        generator = self.yields
        self.lines, self.env, self.yields = saved_lines, saved_env, saved_yields
        name = f"b{len(self.branches):d}"
        self.branches.append(f"def {name}(ctx, r, c, E):\n{body}    return {code}\n")
        return name, generator

    def branch(self, node: Node, kind: str) -> str:
        """A lazy argument, compiled as a function of its own (``b0(ctx, r, c)``).

        The call evaluates it only if it is taken: ``IF(FALSE,A1,5)`` never reads A1.
        """
        saved, saved_env, saved_yields = self.lines, self.env, self.yields
        scoped = bool(self.scopes)
        self.lines, self.yields = [], False
        if scoped:  # a branch inside LET reads its names through the environment
            self.env = "E"
        code = self.argument(node)
        if self.arrays and kind in _LIFTED and not isinstance(node, Leaf):
            code = self.emit(f"ctx.arr({code})")
        body = "".join(f"    {line}\n" for line in self.lines)
        self.branch_yields = self.yields
        self.lines, self.env, self.yields = saved, saved_env, saved_yields
        name = f"b{len(self.branches):d}"
        params = "ctx, r, c, E" if scoped else "ctx, r, c"
        self.branches.append(f"def {name}({params}):\n{body}    return {code}\n")
        return self.emit(f"ops.bind({name}, {self.env})") if scoped else name

    def lazy(self, fn: Callable[..., Any], upper: str, args: list[Node]) -> str:
        """``IF``, ``IFERROR``, ``IFNA`` and ``CHOOSE`` (``ops.lazy_*``).

        The first argument is evaluated, the others only when taken.
        """
        kinds: tuple[str, ...] = getattr(fn, "kinds", ())
        first = self.argument(args[0]) if args else "MISSING"
        if args and self.arrays and not isinstance(args[0], Leaf):
            first = self.emit(f"ctx.arr({first})")
        rest = []
        generator = False
        for k, a in enumerate(args[1:]):
            rest.append(self.branch(a, kinds[min(k + 1, len(kinds) - 1)]))
            generator = generator or self.branch_yields
        helper = _LAZY[upper]
        names = "".join(f", {b}" for b in rest)
        call = f"ops.{helper}({self.function(fn)}, {first}, r, c{names})"
        if generator:  # a branch calls a function value: run on the call stack (ops.glazy_*)
            self.yields = True
            return self.emit(f"yield from ops.g{call[4:]}")
        return self.emit(call)

    def call(self, name: Token, prefix: Token | None, args: list[Node]) -> str:
        if prefix is None and name.text.upper() == "_XLFN.LET" and len(args) >= 2:
            return self.let(args)
        if prefix is None and name.text.upper() == "_XLFN.LAMBDA":
            return self.closure(args)
        if prefix is None and name.text[:6].casefold() == _PARAM:  # a name bound to a function
            return self.invoke(self.bound(name.text), args)
        if prefix is None and name.text.upper() == "_XLFN.ISOMITTED" and len(args) == 1:
            omitted = self.omitted(args[0])
            if omitted is not None:
                self.calls.add("_XLFN.ISOMITTED")
                return omitted
        fn = self.registry.get(name.text.upper()) if prefix is None else None
        if fn is not None and name.text.upper() in _LAZY and args:
            self.calls.add(name.text.upper())
            return self.lazy(fn, name.text.upper(), args)
        if name.text.upper() in ("ROW", "COLUMN") and len(args) == 1 and _is_union(args[0]):
            self.calls.add(name.text.upper())
            return self.constant(ErrorValue.REF)  # of several areas (capture array_refs)
        if name.text.upper() in ("AREAS", "ISREF") and len(args) == 1 and _is_union(args[0]):
            self.calls.add(name.text.upper())  # known from the formula alone (capture info_family)
            return self.constant(float(_areas(args[0])) if name.text.upper() == "AREAS" else True)
        kind = _text_literal(args[0]) if name.text.upper() == "CELL" and args else None
        if kind is not None and kind.lower() in CELL_FORMATTING:
            self.forms.add(f'CELL("{kind.lower()}")')
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
                elif not self.arrays and self.ranged(arg):
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
                    self.intersects = self.intersects or self.ranged(arg)
                    code = code if isinstance(arg, Leaf) else self.emit(f"ops.single({code})")
                self.arrays = saved
                parts.append(code)
                continue
            if kind == "array":
                self.arrays += 1
                saved, self.padded = self.padded, bool(getattr(fn, "padded", False))
                parts.append(self.argument(arg))
                self.arrays -= 1
                self.padded = saved
                continue
            outer = self.padded  # set here only for an element-wise function, as expr() decides
            if kind == "raw" and getattr(fn, "padded", False):
                self.padded = True
            elif not (kind in _LIFTED and getattr(fn, "elementwise", False)):
                self.padded = False
            code = self.argument(arg)
            self.padded = outer
            if self.arrays and kind in _LIFTED and not isinstance(arg, Leaf):
                code = self.emit(f"ctx.{'arr_padded' if self.padded else 'arr'}({code})")
            elif not self.arrays and kind in _SCALAR and self.ranged(arg):
                self.intersects = True
            parts.append(code)
        argv = "".join(f", {code}" for code in parts)
        self.calls.add(name.text.upper() if prefix is None else prefix.text + name.text)
        if fn is None:
            label = name.text if prefix is None else prefix.text + name.text
            if prefix is None and "." not in label:  # a defined name holding a LAMBDA, if any
                self.refs.append(RefSpec("name", None, name=label))
                callee = self.emit(f"ctx.name(None, {self.constant(label)})")
                return self.invoke(callee, args, parts)
            return self.emit(f"ctx.unsupported({self.constant(label)}{argv})")
        return self.emit(f"{self.function(fn)}(ctx{argv})")

    def closure(self, args: list[Node]) -> str:
        """``LAMBDA``: a function value over the environment it is evaluated in (a closure).

        Its body is a function of its parameters' environment (``bK(ctx, r, c, E)``).
        Parameters are ``_xlpm.`` names, optional ones ``_xlop.`` (``[y]`` typed), which the
        body reads as ``_xlpm.``; a body that calls a function value is a generator, run on
        the call stack in ``Ops`` (§5.4).
        """
        self.calls.add("_XLFN.LAMBDA")
        names: dict[str, int] = {}
        required = 0
        for k, param in enumerate(args[:-1]):
            plain = isinstance(param, Reference) and param.prefix is None
            text = param.token.text.casefold() if isinstance(param, Reference) else ""
            kind, key = text[:6], _PARAM + text[6:]
            if (
                not plain
                or kind not in (_PARAM, _OPTIONAL)
                or len(text) == len(kind)
                or key in names
                or (kind == _PARAM and required < k)  # a required one after an optional one
            ):
                names = {}
                break
            names[key] = k
            required += kind == _PARAM
        if not args or len(names) != len(args) - 1:
            self.features.add("a LAMBDA whose parameters Excel would not take")
            return self.constant(ErrorValue.VALUE)
        self.scopes.append(names)
        self.generators.append(set())
        self.parameters.append(True)
        outer = [(folded, self.values.pop(folded, None)) for folded in names]
        self.bodies += 1
        saved_names, self.let_names = self.let_names, 0
        body, _ = self.function_of_env(args[-1])
        bound, self.let_names = self.let_names, saved_names
        self.bodies -= 1
        for folded, value in outer:
            if value is not None:
                self.values[folded] = value
        self.scopes.pop()
        self.generators.pop()
        self.parameters.pop()
        total = len(names)
        return self.emit(
            f"ops.closure({body}, {self.env}, r, c, {required:d}, {total:d}, {bound:d})"
        )

    def invoke(self, callee: str, args: list[Node], parts: list[str] | None = None) -> str:
        """A call of a function value: inside a ``LAMBDA`` body, yielded to the call stack.

        Arguments are passed as they are, a cell as a reference (capture ``lambda_calls``).
        """
        if parts is None:
            parts = [self.argument(a) for a in args]
        argv = "".join(f", {code}" for code in parts)
        if self.bodies:
            self.yields = True
            return self.emit(f"yield {self.emit(f'ops.request({callee}{argv})')}")
        return self.emit(f"ops.invoke({callee}{argv})")

    def omitted(self, node: Node) -> str | None:
        """``ISOMITTED`` of a ``LAMBDA`` parameter, or ``None`` for anything else."""
        while isinstance(node, Group):
            node = node.inner
        if not (isinstance(node, Reference) and node.prefix is None):
            return None
        found = self.lookup(node.token.text) if node.token.text[:6].casefold() == _PARAM else None
        if found is None or not self.parameters[-1 - found[0]]:
            return None
        return self.emit(f"ops.omitted({self.env}, {found[0]:d}, {found[1]:d})")

    def argument(self, node: Node) -> str:
        """A function argument: a single cell is passed as a reference, not its value."""
        if isinstance(node, Reference) and node.token.kind is Kind.CELL:
            sheet = self.sheet(node.prefix)
            r, c = cell_axes(node.token.text, self.row, self.col)
            where = sheet_text(node.prefix) if node.prefix is not None else None
            self.refs.append(RefSpec("box", where, r, c, r, c))
            return self.emit(f"ctx.cref({sheet}, {self.axis(r, 'r')}, {self.axis(c, 'c')})")
        return self.expr(node)


# Functions whose result may be a function value, called where they are (capture lambda_calls).
_MAY_GIVE_A_FUNCTION = frozenset(
    {"_XLFN.LAMBDA", "_XLFN.LET", "IF", "CHOOSE", "_XLFN.IFS", "_XLFN.SWITCH", "IFERROR"}
    | {"_XLFN.IFNA"}
)


def _calls_a_value(callee: Node, registry: Mapping[str, Callable[..., Any]]) -> bool:
    """A call of a registered function's result where that result is never a function value."""
    while isinstance(callee, Group):
        callee = callee.inner
    if not isinstance(callee, Call) or callee.prefix is not None:
        return False
    name = callee.name.text.upper()
    return name in registry and name not in _MAY_GIVE_A_FUNCTION


CELL_FORMATTING = ("format", "protect", "width", "prefix", "color", "parentheses")
"""``CELL``'s kinds that read a cell's formatting, which this package does not compute."""


def _text_literal(node: Node) -> str | None:
    """A text literal's value (``"format"`` gives ``format``), or ``None`` for anything else."""
    while isinstance(node, Group):
        node = node.inner
    if isinstance(node, Leaf) and node.token.kind is Kind.STRING:
        return node.token.text[1:-1].replace('""', '"')
    return None


def _is_union(node: Node) -> bool:
    while isinstance(node, Group):
        node = node.inner
    return isinstance(node, Union)


def _areas(node: Node) -> int:
    """The areas a union names, nested unions included."""
    while isinstance(node, Group):
        node = node.inner
    return sum(_areas(item) for item in node.items) if isinstance(node, Union) else 1


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


EXTERNAL = "external or 3-D reference"
"""The report's label for a reference into another workbook or across sheets."""


def reaches_out(sheet: str, *, name: bool = False) -> bool:
    """A prefix naming another workbook (``[1]``) or several sheets (``S1:S3``).

    ``[0]!Rate`` is this workbook's own name (capture ``workbook_name``), so for a name ``[0]``
    is not another workbook.
    """
    if ":" in sheet:
        return True
    return "[" in sheet and not (name and sheet == "[0]")


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
