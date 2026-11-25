"""The operators compiled formulas call as ``ops.add`` and the like (ARCHITECTURE.md §4.3, §5.4).

One ``Ops`` object serves a workbook; it reads the cell being evaluated from the engine (for
implicit intersection) through ``context``. An error operand gives that error, the left one
first; results without subnormals; overflow is ``#NUM!``.
"""

from __future__ import annotations

import math
from types import GeneratorType
from typing import Any

from xlrecalc._arith import flush, snap
from xlrecalc._arrays import Padded, broadcast, pair
from xlrecalc._coerce import Context, Raw, scalar, to_bool, to_number, to_text
from xlrecalc._compare import compare, switch_equal, utf16_length
from xlrecalc._formula import MISSING
from xlrecalc._numtext import DBL_MIN
from xlrecalc._values import BLANK, Areas, ErrorValue, Lambda, Ref, Value

__all__ = ["Ops"]

NUM, DIV0, VALUE, NULL = ErrorValue.NUM, ErrorValue.DIV0, ErrorValue.VALUE, ErrorValue.NULL
NA = ErrorValue.NA

STACK_LIMIT = 65536
"""Excel's budget for calls of function values in progress: each costs 8, plus 4 a parameter
(optional ones included, given or not) and 4 a ``LET`` name its body binds (captures
``lambda_depth``, ``lambda_more``: 5,461 calls of one parameter, 4,095 of two or of one and a
``LET`` name, 3,276 of three or of one and two names)."""
NESTED_HELPERS = 20
"""Calls of a function value made by ``MAP`` and its siblings, one inside another, that Excel
takes; the next is ``#VALUE!`` (capture ``lambda_more``: recursion through ``MAP`` 20 deep
computes, 21 does not). Each such call is a call stack of its own, so this bounds Python's."""


def _around(refs: tuple[Ref, ...]) -> Ref | ErrorValue:
    """The one range around references on one sheet."""
    first = refs[0]
    if any(r.sheet != first.sheet for r in refs):
        return VALUE
    return Ref(
        first.source,
        first.sheet,
        min(r.r1 for r in refs),
        min(r.c1 for r in refs),
        max(r.r2 for r in refs),
        max(r.c2 for r in refs),
    )


_ARRAYS = (tuple, Padded)
"""What an operator works on element by element: an array, dense or padded."""


def _finite(x: float) -> float | ErrorValue:
    return flush(x) if math.isfinite(x) else NUM


_isfinite = math.isfinite


class Ops:
    """Operators for compiled formulas.

    Each arithmetic operator first takes two floats, the common case, without coercion; the
    result is the one the general path gives (``_finite``, ``snap``), written out inline.
    """

    def __init__(self, context: Context) -> None:
        self.context = context
        self.stack_cost = 0  # of the calls of function values in progress
        self.helpers = 0  # calls by MAP and its siblings in progress, one inside another

    def _limit(self) -> int | None:
        return getattr(self.context, "max_array_cells", None)

    def _pair(self, a: Raw, b: Raw) -> tuple[float, float] | ErrorValue:
        x = to_number(a, self.context)
        if isinstance(x, ErrorValue):
            return x
        y = to_number(b, self.context)
        if isinstance(y, ErrorValue):
            return y
        return x, y

    def add(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a + b
            return (0.0 if -DBL_MIN < r < DBL_MIN else r) if _isfinite(r) else NUM
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.add, a, b)
        p = self._pair(a, b)
        return p if isinstance(p, ErrorValue) else _finite(p[0] + p[1])

    def sub(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a - b
            return (0.0 if -DBL_MIN < r < DBL_MIN else r) if _isfinite(r) else NUM
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.sub, a, b)
        p = self._pair(a, b)
        return p if isinstance(p, ErrorValue) else _finite(p[0] - p[1])

    def add_top(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a + b
            return snap(a, r, self.context.quirks) if _isfinite(r) else NUM
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.add_top, a, b)
        p = self._pair(a, b)
        if isinstance(p, ErrorValue):
            return p
        r = p[0] + p[1]
        return snap(p[0], r, self.context.quirks) if math.isfinite(r) else NUM

    def sub_top(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a - b
            return snap(a, r, self.context.quirks) if _isfinite(r) else NUM
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.sub_top, a, b)
        p = self._pair(a, b)
        if isinstance(p, ErrorValue):
            return p
        r = p[0] - p[1]
        return snap(p[0], r, self.context.quirks) if math.isfinite(r) else NUM

    def mul(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a * b
            return (0.0 if -DBL_MIN < r < DBL_MIN else r) if _isfinite(r) else NUM
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.mul, a, b)
        p = self._pair(a, b)
        return p if isinstance(p, ErrorValue) else _finite(p[0] * p[1])

    def div(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            if b == 0:
                return DIV0
            r = a / b
            return (0.0 if -DBL_MIN < r < DBL_MIN else r) if _isfinite(r) else NUM
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.div, a, b)
        p = self._pair(a, b)
        if isinstance(p, ErrorValue):
            return p
        if p[1] == 0:
            return DIV0
        return _finite(p[0] / p[1])

    def pow(self, a: Raw, b: Raw) -> Value:
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.pow, a, b)
        p = self._pair(a, b)
        if isinstance(p, ErrorValue):
            return p
        from xlrecalc._functions._math import power

        return power(*p)

    def neg(self, a: Raw) -> Value:
        if type(a) is Padded:
            return a.map(self.neg)  # type: ignore[return-value]
        if type(a) is tuple:
            return broadcast(self.neg, (a,), (True,), self._limit())  # type: ignore[return-value]
        x = to_number(a, self.context)
        return x if isinstance(x, ErrorValue) else -x if x else 0.0

    def pos(self, a: Raw) -> Value:
        return a if type(a) in _ARRAYS else scalar(a, self.context)  # type: ignore[return-value]

    def percent(self, a: Raw) -> Value:
        if type(a) is Padded:
            return a.map(self.percent)  # type: ignore[return-value]
        if type(a) is tuple:
            return broadcast(self.percent, (a,), (True,), self._limit())  # type: ignore[return-value]
        x = to_number(a, self.context)
        return x if isinstance(x, ErrorValue) else flush(x / 100)

    def concat(self, a: Raw, b: Raw) -> Value:
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(self.concat, a, b)
        x = to_text(a, self.context)
        if isinstance(x, ErrorValue):
            return x
        y = to_text(b, self.context)
        if isinstance(y, ErrorValue):
            return y
        if len(x) + len(y) > 16383 and utf16_length(x) + utf16_length(y) > 32767:
            return VALUE  # Excel's limit counts UTF-16 units (capture text_family)
        return x + y

    def _cmp(self, a: Raw, b: Raw, op: str) -> Value:
        if type(a) in _ARRAYS or type(b) in _ARRAYS:
            return self._each(lambda x, y: self._cmp(x, y, op), a, b)
        if type(a) is Lambda or type(b) is Lambda:
            return VALUE
        return compare(
            scalar(a, self.context),
            scalar(b, self.context),
            op,  # type: ignore[arg-type]
            self.context.quirks,
            self.context.platform,
        )

    def eq(self, a: Raw, b: Raw) -> Value:
        return self._cmp(a, b, "=")

    def ne(self, a: Raw, b: Raw) -> Value:
        return self._cmp(a, b, "<>")

    def lt(self, a: Raw, b: Raw) -> Value:
        return self._cmp(a, b, "<")

    def gt(self, a: Raw, b: Raw) -> Value:
        return self._cmp(a, b, ">")

    def le(self, a: Raw, b: Raw) -> Value:
        return self._cmp(a, b, "<=")

    def ge(self, a: Raw, b: Raw) -> Value:
        return self._cmp(a, b, ">=")

    def single(self, a: Raw) -> Value:
        """``@``: implicit intersection of a reference; an array's first element."""
        if type(a) is tuple:
            return a[0][0]  # type: ignore[no-any-return]
        return scalar(a, self.context)

    # Lazy functions (codegen ``_LAZY``): ``fn`` is the registered function, the branches are
    # compiled functions of (ctx, r, c), called only when taken. An array first argument takes
    # every branch, element by element, through ``fn``.

    def lazy_if(self, fn: Any, test: Raw, r: int, c: int, *branches: Any) -> Value:
        ctx = self.context
        if type(test) in _ARRAYS:
            return fn(ctx, test, *(b(ctx, r, c) for b in branches))  # type: ignore[no-any-return]
        t = to_bool(test, ctx)
        if isinstance(t, ErrorValue):
            return t
        k = 0 if t else 1
        if k >= len(branches):
            return fn(ctx, t)  # type: ignore[no-any-return]
        chosen = branches[k](ctx, r, c)
        return fn(ctx, t, chosen, chosen)  # type: ignore[no-any-return]

    def _lazy_error(self, fn: Any, value: Raw, r: int, c: int, alt: Any, na_only: bool) -> Value:
        ctx = self.context
        if type(value) is tuple:
            return fn(ctx, value, alt(ctx, r, c))  # type: ignore[no-any-return]
        v = scalar(value, ctx)
        caught = isinstance(v, ErrorValue) and (not na_only or v is NA)
        return fn(ctx, value, alt(ctx, r, c) if caught else BLANK)  # type: ignore[no-any-return]

    def lazy_iferror(self, fn: Any, value: Raw, r: int, c: int, alt: Any) -> Value:
        return self._lazy_error(fn, value, r, c, alt, False)

    def lazy_ifna(self, fn: Any, value: Raw, r: int, c: int, alt: Any) -> Value:
        return self._lazy_error(fn, value, r, c, alt, True)

    def lazy_ifs(self, fn: Any, first: Raw, r: int, c: int, *rest: Any) -> Value:
        """``IFS``: conditions in turn, and only the value after the one that holds."""
        ctx = self.context
        condition = first
        k = 0
        while True:
            if type(condition) is tuple:  # an array: every argument, element by element
                done = [first, *(b(ctx, r, c) for b in rest)]
                return fn(ctx, *done)  # type: ignore[no-any-return]
            t = to_bool(condition, ctx)
            if isinstance(t, ErrorValue):
                return t
            if t:
                return rest[k](ctx, r, c) if k < len(rest) else VALUE
            if k + 1 >= len(rest):
                return NA if k < len(rest) else VALUE
            condition = rest[k + 1](ctx, r, c)
            k += 2

    def lazy_switch(self, fn: Any, first: Raw, r: int, c: int, *rest: Any) -> Value:
        """``SWITCH``: values in turn, and only the result after the one that matches."""
        ctx = self.context
        if type(first) is tuple:
            return fn(ctx, first, *(b(ctx, r, c) for b in rest))  # type: ignore[no-any-return]
        v = scalar(first, ctx)
        if isinstance(v, ErrorValue):
            return v
        for k in range(0, len(rest) - len(rest) % 2, 2):
            candidate = rest[k](ctx, r, c)
            if type(candidate) is tuple:  # an array: every argument, element by element
                return fn(ctx, first, *(b(ctx, r, c) for b in rest))  # type: ignore[no-any-return]
            value = scalar(candidate, ctx)
            if isinstance(value, ErrorValue):
                return value
            if switch_equal(v, value):
                return rest[k + 1](ctx, r, c)  # type: ignore[no-any-return]
        return rest[-1](ctx, r, c) if len(rest) % 2 else NA

    def lazy_choose(self, fn: Any, index: Raw, r: int, c: int, *choices: Any) -> Value:
        ctx = self.context
        if type(index) is tuple:
            return fn(ctx, index, *(b(ctx, r, c) for b in choices))  # type: ignore[no-any-return]
        n = to_number(index, ctx)
        if isinstance(n, ErrorValue):
            return n
        k = math.trunc(n)
        args = [b(ctx, r, c) if i == k - 1 else MISSING for i, b in enumerate(choices)]
        return fn(ctx, n, *args)  # type: ignore[no-any-return]

    def _each(self, op: Any, a: Raw, b: Raw) -> Value:
        """A binary operation over arrays, element by element (§7.7)."""
        x = a if type(a) in _ARRAYS else scalar(a, self.context)
        y = b if type(b) in _ARRAYS else scalar(b, self.context)
        if type(x) is Padded or type(y) is Padded:
            kept = pair(op, x, y)
            if kept is not None:
                return kept  # type: ignore[return-value]
            x = x.dense(self._limit()) if type(x) is Padded else x
            y = y.dense(self._limit()) if type(y) is Padded else y
        return broadcast(op, (x, y), (True, True), self._limit())  # type: ignore[return-value]

    def range(self, a: Raw, b: Raw) -> Value:
        # Between computed references (A1:INDEX(...), a name): an error in either is the result,
        # and references on two sheets or anything else are a value error (capture names). A
        # union's areas are read as one range around them all (capture unions).
        if isinstance(a, ErrorValue):
            return a
        if isinstance(b, ErrorValue):
            return b
        if type(a) is Areas and a.union:
            a = _around(a.refs)
        if type(b) is Areas and b.union:
            b = _around(b.refs)
        if isinstance(a, Ref) and isinstance(b, Ref) and a.sheet == b.sheet:
            return Ref(  # type: ignore[return-value]
                a.source,
                a.sheet,
                min(a.r1, b.r1),
                min(a.c1, b.c1),
                max(a.r2, b.r2),
                max(a.c2, b.c2),
            )
        return VALUE

    def intersect(self, a: Raw, b: Raw) -> Value:
        # A union intersects area by area, and the areas that meet are the result (capture
        # unions: SUM((X1:X3,Z1:Z3) X2:Z2) is X2+Z2).
        left = a.refs if type(a) is Areas and a.union else (a,)
        right = b.refs if type(b) is Areas and b.union else (b,)
        met = []
        for x in left:
            for y in right:
                if not (isinstance(x, Ref) and isinstance(y, Ref) and x.sheet == y.sheet):
                    return VALUE
                r1, c1 = max(x.r1, y.r1), max(x.c1, y.c1)
                r2, c2 = min(x.r2, y.r2), min(x.c2, y.c2)
                if r1 <= r2 and c1 <= c2:
                    met.append(Ref(x.source, x.sheet, r1, c1, r2, c2))
        if not met:
            return NULL
        return met[0] if len(met) == 1 else Areas(tuple(met), union=True)  # type: ignore[return-value]

    def union(self, *items: Raw) -> Value:
        """A union of references: its areas as written, a nested union's flattened (capture unions).

        An error among them is the result; areas on two sheets are ``#VALUE!``, and a 3-D
        reference among them ``#NAME?``, as Excel reads such a union in a file.
        """
        refs: list[Ref] = []
        for item in items:
            if type(item) is Areas:
                if not item.union:
                    return ErrorValue.NAME
                refs += item.refs
            elif isinstance(item, Ref):
                refs.append(item)
            elif isinstance(item, ErrorValue):
                return item
            else:
                return VALUE
        if any(r.sheet != refs[0].sheet for r in refs):
            return VALUE
        return Areas(tuple(refs), union=True)  # type: ignore[return-value]

    # LAMBDA (captures lambda_calls, lambda_names, lambda_depth) ---------------------------------

    @staticmethod
    def closure(
        body: Any, env: Env | None, r: int, c: int, required: int, total: int, names: int
    ) -> Lambda:
        """``LAMBDA``: a function value over the environment it is evaluated in.

        ``names``: the ``LET`` names its body binds, each costing a call 4, as a parameter does.
        """
        return Lambda(body, env, r, c, required, total, 8 + 4 * (total + names))

    def invoke(self, fn: Any, *args: Any) -> Any:
        """A call of a function value in a formula, outside any ``LAMBDA`` body."""
        return self.call(fn, args)

    @staticmethod
    def request(fn: Any, *args: Any) -> tuple[Any, tuple[Any, ...]]:
        """A call inside a ``LAMBDA`` body, as its generator yields it to ``call``."""
        return fn, args

    def call(self, fn: Any, args: tuple[Any, ...]) -> Any:
        """A call of a function value, run to its result on an explicit stack.

        A body that calls a function value is a generator that yields each call as
        ``(function, arguments)``; the loop runs the call and sends its result back, so
        recursion as deep as Excel's budget never uses Python's stack. A call past the budget
        (``STACK_LIMIT``) is ``#NUM!``, as in Excel.
        """
        stack: list[tuple[Any, int]] = []
        start = self.stack_cost
        try:
            value = self._enter(fn, args, stack)
            while stack:
                try:
                    fn, args = stack[-1][0].send(value)
                except StopIteration as done:
                    self.stack_cost -= stack.pop()[1]
                    value = done.value
                    continue
                value = self._enter(fn, args, stack)
            return value
        finally:
            self.stack_cost = start

    def helper_call(self, fn: Any, args: tuple[Any, ...]) -> Any:
        """A call by ``MAP`` or a sibling: ``#VALUE!`` past ``NESTED_HELPERS`` such calls."""
        if self.helpers >= NESTED_HELPERS:
            return VALUE
        self.helpers += 1
        try:
            return self.call(fn, args)
        finally:
            self.helpers -= 1

    def _enter(self, fn: Any, args: tuple[Any, ...], stack: list[tuple[Any, int]]) -> Any:
        """Start a call: its result, or ``None`` with its body pushed on ``stack`` to run."""
        if type(fn) is not Lambda:  # an error is the result, another value #VALUE! (lambda_more)
            fn = scalar(fn, self.context)
            return fn if isinstance(fn, ErrorValue) else VALUE
        n = len(args)
        if n < fn.required or n > fn.total:  # capture lambda_calls
            return VALUE
        cost = fn.cost
        if self.stack_cost + cost >= STACK_LIMIT:
            return NUM
        env = Env(self.context, fn.r, fn.c, fn.env, ())
        env.values = [BLANK if a is MISSING else a for a in args]
        if n < fn.total or MISSING in args:  # an omitted argument reads as a blank
            env.values.extend([BLANK] * (fn.total - n))
            env.omitted = frozenset(i for i in range(fn.total) if i >= n or args[i] is MISSING)
        self.stack_cost += cost
        try:
            result = fn.body(self.context, fn.r, fn.c, env)
        except BaseException:
            self.stack_cost -= cost
            raise
        if type(result) is GeneratorType:
            stack.append((result, cost))
            return None
        self.stack_cost -= cost
        return result

    @staticmethod
    def omitted(env: Env, depth: int, index: int) -> bool:
        """``ISOMITTED`` of a parameter: no argument, or an empty one, was given for it."""
        for _ in range(depth):
            env = env.parent  # type: ignore[assignment]
        return index in env.omitted

    @staticmethod
    def gforce(env: Env, depth: int, index: int) -> Any:
        """``force`` for a binding that calls a function value: a generator, as its caller."""
        for _ in range(depth):
            env = env.parent  # type: ignore[assignment]
        value = env.values[index]
        if value is _UNSET:
            value = env.thunks[index](env.ctx, env.r, env.c, env)
            if type(value) is GeneratorType:
                value = yield from value
            env.values[index] = value
        return value

    # The lazy helpers again, as generators, for branches that call a function value: a
    # branch's result may be a generator, run from here so its calls reach the call stack.

    def glazy_if(self, fn: Any, test: Raw, r: int, c: int, *branches: Any) -> Any:
        ctx = self.context
        if type(test) in _ARRAYS:
            done = yield from _taken(branches, ctx, r, c)
            return fn(ctx, test, *done)
        t = to_bool(test, ctx)
        if isinstance(t, ErrorValue):
            return t
        k = 0 if t else 1
        if k >= len(branches):
            return fn(ctx, t)
        (chosen,) = yield from _taken(branches[k : k + 1], ctx, r, c)
        return fn(ctx, t, chosen, chosen)

    def _glazy_error(self, fn: Any, value: Raw, r: int, c: int, alt: Any, na_only: bool) -> Any:
        ctx = self.context
        if type(value) is tuple:
            (other,) = yield from _taken((alt,), ctx, r, c)
            return fn(ctx, value, other)
        v = scalar(value, ctx)
        if isinstance(v, ErrorValue) and (not na_only or v is NA):
            (other,) = yield from _taken((alt,), ctx, r, c)
            return fn(ctx, value, other)
        return fn(ctx, value, BLANK)

    def glazy_iferror(self, fn: Any, value: Raw, r: int, c: int, alt: Any) -> Any:
        return (yield from self._glazy_error(fn, value, r, c, alt, False))

    def glazy_ifna(self, fn: Any, value: Raw, r: int, c: int, alt: Any) -> Any:
        return (yield from self._glazy_error(fn, value, r, c, alt, True))

    def glazy_ifs(self, fn: Any, first: Raw, r: int, c: int, *rest: Any) -> Any:
        ctx = self.context
        condition = first
        k = 0
        while True:
            if type(condition) is tuple:
                done = yield from _taken(rest, ctx, r, c)
                return fn(ctx, first, *done)
            t = to_bool(condition, ctx)
            if isinstance(t, ErrorValue):
                return t
            if t:
                if k >= len(rest):
                    return VALUE
                (chosen,) = yield from _taken(rest[k : k + 1], ctx, r, c)
                return chosen
            if k + 1 >= len(rest):
                return NA if k < len(rest) else VALUE
            (condition,) = yield from _taken(rest[k + 1 : k + 2], ctx, r, c)
            k += 2

    def glazy_switch(self, fn: Any, first: Raw, r: int, c: int, *rest: Any) -> Any:
        ctx = self.context
        if type(first) is tuple:
            done = yield from _taken(rest, ctx, r, c)
            return fn(ctx, first, *done)
        v = scalar(first, ctx)
        if isinstance(v, ErrorValue):
            return v
        for k in range(0, len(rest) - len(rest) % 2, 2):
            (candidate,) = yield from _taken(rest[k : k + 1], ctx, r, c)
            if type(candidate) is tuple:
                done = yield from _taken(rest, ctx, r, c)
                return fn(ctx, first, *done)
            value = scalar(candidate, ctx)
            if isinstance(value, ErrorValue):
                return value
            if switch_equal(v, value):
                (chosen,) = yield from _taken(rest[k + 1 : k + 2], ctx, r, c)
                return chosen
        if len(rest) % 2:
            (chosen,) = yield from _taken(rest[-1:], ctx, r, c)
            return chosen
        return NA

    def glazy_choose(self, fn: Any, index: Raw, r: int, c: int, *choices: Any) -> Any:
        ctx = self.context
        if type(index) is tuple:
            done = yield from _taken(choices, ctx, r, c)
            return fn(ctx, index, *done)
        n = to_number(index, ctx)
        if isinstance(n, ErrorValue):
            return n
        k = math.trunc(n)
        chosen: Any = MISSING
        if 1 <= k <= len(choices):
            (chosen,) = yield from _taken(choices[k - 1 : k], ctx, r, c)
        args = [chosen if i == k - 1 else MISSING for i in range(len(choices))]
        return fn(ctx, n, *args)

    # LET (capture dyn_let) ----------------------------------------------------------------------

    @staticmethod
    def scope(ctx: object, r: int, c: int, parent: Env | None, thunks: tuple[Any, ...]) -> Env:
        """A ``LET``'s environment: its bindings, none evaluated yet."""
        return Env(ctx, r, c, parent, thunks)

    @staticmethod
    def force(env: Env, depth: int, index: int) -> Any:
        """The value of a binding ``depth`` scopes out, evaluated on its first read only."""
        for _ in range(depth):
            env = env.parent  # type: ignore[assignment]
        value = env.values[index]
        if value is _UNSET:
            value = env.thunks[index](env.ctx, env.r, env.c, env)
            env.values[index] = value
        return value

    @staticmethod
    def bind(branch: Any, env: Env) -> Any:
        """A lazy branch inside ``LET``, called as other branches are: ``branch(ctx, r, c)``."""
        return lambda ctx, r, c: branch(ctx, r, c, env)


_UNSET = object()  # a binding not evaluated yet


def _taken(branches: tuple[Any, ...], ctx: Any, r: int, c: int) -> Any:
    """The values of lazy branches, each run on the call stack if it is a generator."""
    out = []
    for b in branches:
        v = b(ctx, r, c)
        if type(v) is GeneratorType:
            v = yield from v
        out.append(v)
    return out


class Env:
    """The bindings of one ``LET`` call, or the arguments of one call of a function value.

    It lives while it is evaluated, and links to the scope around it.
    """

    __slots__ = ("c", "ctx", "omitted", "parent", "r", "thunks", "values")

    def __init__(
        self, ctx: object, r: int, c: int, parent: Env | None, thunks: tuple[Any, ...]
    ) -> None:
        self.ctx, self.r, self.c, self.parent, self.thunks = ctx, r, c, parent, thunks
        self.values: list[Any] = [_UNSET] * len(thunks)
        self.omitted: frozenset[int] = frozenset()  # arguments not given (ISOMITTED)
