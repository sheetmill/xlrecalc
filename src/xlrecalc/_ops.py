"""The operators compiled formulas call as ``ops.add`` and the like (ARCHITECTURE.md §4.3, §5.4).

One ``Ops`` object serves a workbook; it reads the cell being evaluated from the engine (for
implicit intersection) through ``context``. An error operand gives that error, the left one
first; results without subnormals; overflow is ``#NUM!``.
"""

from __future__ import annotations

import math
from typing import Any

from xlrecalc._arith import flush, snap
from xlrecalc._arrays import broadcast
from xlrecalc._coerce import Context, Raw, scalar, to_bool, to_number, to_text
from xlrecalc._compare import compare
from xlrecalc._formula import MISSING
from xlrecalc._numtext import DBL_MIN
from xlrecalc._values import BLANK, ErrorValue, Ref, Value

__all__ = ["Ops"]

NUM, DIV0, VALUE, NULL = ErrorValue.NUM, ErrorValue.DIV0, ErrorValue.VALUE, ErrorValue.NULL
NA = ErrorValue.NA


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
        if type(a) is tuple or type(b) is tuple:
            return self._each(self.add, a, b)
        p = self._pair(a, b)
        return p if isinstance(p, ErrorValue) else _finite(p[0] + p[1])

    def sub(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a - b
            return (0.0 if -DBL_MIN < r < DBL_MIN else r) if _isfinite(r) else NUM
        if type(a) is tuple or type(b) is tuple:
            return self._each(self.sub, a, b)
        p = self._pair(a, b)
        return p if isinstance(p, ErrorValue) else _finite(p[0] - p[1])

    def add_top(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            r = a + b
            return snap(a, r, self.context.quirks) if _isfinite(r) else NUM
        if type(a) is tuple or type(b) is tuple:
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
        if type(a) is tuple or type(b) is tuple:
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
        if type(a) is tuple or type(b) is tuple:
            return self._each(self.mul, a, b)
        p = self._pair(a, b)
        return p if isinstance(p, ErrorValue) else _finite(p[0] * p[1])

    def div(self, a: Raw, b: Raw) -> Value:
        if type(a) is float and type(b) is float:
            if b == 0:
                return DIV0
            r = a / b
            return (0.0 if -DBL_MIN < r < DBL_MIN else r) if _isfinite(r) else NUM
        if type(a) is tuple or type(b) is tuple:
            return self._each(self.div, a, b)
        p = self._pair(a, b)
        if isinstance(p, ErrorValue):
            return p
        if p[1] == 0:
            return DIV0
        return _finite(p[0] / p[1])

    def pow(self, a: Raw, b: Raw) -> Value:
        if type(a) is tuple or type(b) is tuple:
            return self._each(self.pow, a, b)
        p = self._pair(a, b)
        if isinstance(p, ErrorValue):
            return p
        from xlrecalc._functions._math import power

        return power(*p)

    def neg(self, a: Raw) -> Value:
        if type(a) is tuple:
            return broadcast(self.neg, (a,), (True,))  # type: ignore[return-value]
        x = to_number(a, self.context)
        return x if isinstance(x, ErrorValue) else -x if x else 0.0

    def pos(self, a: Raw) -> Value:
        return a if type(a) is tuple else scalar(a, self.context)  # type: ignore[return-value]

    def percent(self, a: Raw) -> Value:
        if type(a) is tuple:
            return broadcast(self.percent, (a,), (True,))  # type: ignore[return-value]
        x = to_number(a, self.context)
        return x if isinstance(x, ErrorValue) else flush(x / 100)

    def concat(self, a: Raw, b: Raw) -> Value:
        if type(a) is tuple or type(b) is tuple:
            return self._each(self.concat, a, b)
        x = to_text(a, self.context)
        if isinstance(x, ErrorValue):
            return x
        y = to_text(b, self.context)
        if isinstance(y, ErrorValue):
            return y
        return x + y if len(x) + len(y) <= 32767 else VALUE

    def _cmp(self, a: Raw, b: Raw, op: str) -> Value:
        if type(a) is tuple or type(b) is tuple:
            return self._each(lambda x, y: self._cmp(x, y, op), a, b)
        return compare(scalar(a, self.context), scalar(b, self.context), op, self.context.quirks)  # type: ignore[arg-type]

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
        if type(test) is tuple:
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
        x = a if type(a) is tuple else scalar(a, self.context)
        y = b if type(b) is tuple else scalar(b, self.context)
        return broadcast(op, (x, y), (True, True))  # type: ignore[return-value]

    def range(self, a: Raw, b: Raw) -> Value:
        # A range between computed references (A1:INDEX(...)) needs functions that return
        # references; until then it is a value error.
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
        if isinstance(a, Ref) and isinstance(b, Ref) and a.sheet == b.sheet:
            r1, c1 = max(a.r1, b.r1), max(a.c1, b.c1)
            r2, c2 = min(a.r2, b.r2), min(a.c2, b.c2)
            if r1 > r2 or c1 > c2:
                return NULL
            return Ref(a.source, a.sheet, r1, c1, r2, c2)  # type: ignore[return-value]
        return VALUE

    def union(self, *items: Raw) -> Value:
        return VALUE  # multi-area references come with the functions that take them

    def invoke(self, ctx: object, fn: Raw, *args: Raw) -> Value:
        return VALUE
