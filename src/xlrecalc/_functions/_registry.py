"""The function registry (ARCHITECTURE.md §8.1).

``@function("ROUND", "number", "number")`` registers a function with the kinds of its
arguments; the wrapper checks the argument count, coerces each argument, and returns the first
error among scalar arguments unless the function inspects errors itself (``errors=True``). The
core function receives the evaluation context first, then the coerced arguments:

- ``number``, ``text``, ``logical``: one value converted (``_coerce.py``), an omitted argument
  as 0, ``""`` or FALSE;
- ``value``: one value as it is (references intersected, an omitted argument blank);
- ``raw``: the argument untouched, a ``Ref`` for references, ``MISSING`` when omitted;
- ``array``: untouched too, and evaluated as an array even in a legacy formula, where an
  operator over a range would otherwise take its implicit intersection (``SUMPRODUCT``'s
  arguments, ``INDEX``'s array: capture ``arrays``);
- ``branch``: untouched, like ``raw``, but taken element by element when the call is (``IF``'s
  results);
- ``matrix``: a whole array, never taken element by element; a reference in a legacy formula
  is intersected, as for ``value``, and in an array context it arrives as an array
  (``TRANSPOSE``: capture ``misc``).

When an argument of a one-value kind (or a ``branch``) is an array, the function is applied to
each element and returns an array (§7.7).

The last kind repeats for variadic functions. Names are the file form, upper case, with
``_xlfn.`` for functions newer than Excel 2007.

Results are compared with Excel's bit for bit. A function whose results could not be made
bit-exact within a time-boxed attempt declares a small tolerance in units in the last place
(``ulps``) with its reason (``ulps_reason``); the capture tests count rows within it separately
from exact matches, and ``QUIRKS.md`` lists it as a known deviation. An iterative function
(IRR, XIRR, RATE, YIELD), whose result Excel itself guarantees only to its documented stopping
tolerance, may instead declare a relative tolerance (``rel``, 1E-9, with ``rel_reason``),
counted apart from both.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from xlrecalc._arrays import broadcast, to_array
from xlrecalc._coerce import Context, scalar, to_bool, to_number, to_text
from xlrecalc._values import ErrorValue, Ref

__all__ = ["REGISTRY", "SPECS", "Spec", "function"]

VALUE = ErrorValue.VALUE

_LIFTED = frozenset({"number", "text", "logical", "value", "branch"})
_CONVERT: dict[str, Callable[[Any, Context], Any]] = {
    "number": to_number,
    "text": to_text,
    "logical": to_bool,
    "value": scalar,
    "matrix": lambda a, ctx: scalar(a, ctx) if isinstance(a, Ref) else a,
}


@dataclass(frozen=True, slots=True)
class Spec:
    """A registered function: its kinds, argument counts and flags."""

    name: str
    kinds: tuple[str, ...]
    min_args: int
    max_args: int | None
    errors: bool
    volatile: bool
    category: str
    ulps: int = 0
    ulps_reason: str = ""
    status: str = "implemented"
    note: str = ""
    rel: float = 0.0
    rel_reason: str = ""


REGISTRY: dict[str, Callable[..., Any]] = {}
SPECS: dict[str, Spec] = {}


def function(
    name: str,
    *kinds: str,
    min_args: int | None = None,
    variadic: bool = False,
    errors: bool = False,
    volatile: bool = False,
    category: str = "",
    ulps: int = 0,
    ulps_reason: str = "",
    status: str = "implemented",
    note: str = "",
    rel: float = 0.0,
    rel_reason: str = "",
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register a function under its file-form name."""
    if (ulps and not ulps_reason) or (rel and not rel_reason):
        raise ValueError(f"{name}: a tolerance needs its reason")
    if status not in ("implemented", "partial") or (status == "partial") != bool(note):
        raise ValueError(f"{name}: status 'implemented', or 'partial' with a note")
    lo = len(kinds) if min_args is None else min_args
    hi = None if variadic else len(kinds)

    def register(core: Callable[..., Any]) -> Callable[..., Any]:
        converters = [_CONVERT.get(k) for k in kinds]
        last = converters[-1] if converters else None

        lifted_kinds = [k in _LIFTED for k in kinds]
        last_lifted = lifted_kinds[-1] if lifted_kinds else False

        plain = not any(lifted_kinds) and all(c is None for c in converters)

        def call(ctx: Context, *args: Any) -> Any:
            n = len(args)
            if n < lo or (hi is not None and n > hi):
                return VALUE
            if plain:  # every argument passed as it is (SUM, COUNT, ...): nothing to coerce
                return core(ctx, *args)
            lifted = [(lifted_kinds[i] if i < len(lifted_kinds) else last_lifted) for i in range(n)]
            if any(lifted[i] and type(a) is tuple for i, a in enumerate(args)):
                spread = [
                    to_array(a) if lifted[i] and isinstance(a, Ref) and not a.is_cell() else a
                    for i, a in enumerate(args)
                ]
                limit = getattr(ctx, "max_array_cells", None)
                return broadcast(lambda *xs: call(ctx, *xs), spread, lifted, limit)
            coerced = []
            for i, arg in enumerate(args):
                convert = converters[i] if i < len(converters) else last
                if convert is None:
                    coerced.append(arg)
                    continue
                value = convert(arg, ctx)
                if isinstance(value, ErrorValue) and not errors:
                    return value
                coerced.append(value)
            return core(ctx, *coerced)

        call.kinds = kinds  # type: ignore[attr-defined]
        call.__name__ = core.__name__
        call.__doc__ = core.__doc__
        upper = name.upper()
        REGISTRY[upper] = call
        SPECS[upper] = Spec(
            upper, kinds, lo, hi, errors, volatile, category, ulps, ulps_reason, status, note,
            rel, rel_reason,
        )  # fmt: skip
        return core

    return register
