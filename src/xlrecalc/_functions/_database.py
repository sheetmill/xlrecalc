"""Database functions: DSUM, DAVERAGE, DCOUNT, DMAX, DMIN (capture ``misc``).

A database is a range whose first row names its columns; a criteria range has names in its
first row and conditions below. A record matches when every condition of some criteria row
holds; a blank criteria row matches every record, and a condition under a name the database
lacks matches none. A text condition without an operator matches text that starts with it
(``Apple`` matches ``Applesauce``), as in Excel; ``=Apple`` matches only ``Apple``.
Conditions computed by a formula (a name the database lacks over a formula) are not supported:
they match no record.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from typing import Any

from xlrecalc._arrays import to_array
from xlrecalc._coerce import Context, scalar
from xlrecalc._compare import criterion
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._values import BLANK, ErrorValue, Ref

VALUE, DIV0 = ErrorValue.VALUE, ErrorValue.DIV0
DB = "database"


def _key(v: Any) -> str | None:
    return v.casefold() if isinstance(v, str) else None


def _condition(value: Any, ctx: Context) -> Callable[[Any], bool]:
    if isinstance(value, str) and value and value[0] not in "<>=":
        value += "*"
    return criterion(value, ctx.quirks)


def _matching(db: Any, field: Any, crit: Any, ctx: Context) -> list[Any] | ErrorValue:
    """The field's values in the records that match, or the error the arguments give."""
    if not isinstance(db, Ref) or not isinstance(crit, Ref):
        return VALUE
    table, conditions = to_array(db), to_array(crit)
    if len(table) < 2 or len(conditions) < 2:
        return VALUE
    names = [_key(v) for v in table[0]]
    if field is MISSING:
        column = None
    else:
        f = scalar(field, ctx)
        if isinstance(f, ErrorValue):
            return f
        if isinstance(f, bool):  # a logical is a column number (captured: TRUE is column 1)
            f = float(f)
        if isinstance(f, str):
            if f.casefold() not in names:
                return VALUE
            column = names.index(f.casefold())
        elif isinstance(f, float):
            k = math.trunc(f)
            if not 1 <= k <= len(names):
                return VALUE
            column = k - 1
        else:
            return VALUE
    rows: list[list[tuple[int, Callable[[Any], bool]] | None]] = []
    for crow in conditions[1:]:
        tests: list[tuple[int, Callable[[Any], bool]] | None] = []
        for name, value in zip(conditions[0], crow, strict=True):
            if value is BLANK:
                continue
            key = _key(name)
            tests.append((names.index(key), _condition(value, ctx)) if key in names else None)
        rows.append(tests)
    return [
        record[column] if column is not None else record
        for record in table[1:]
        if any(all(t is not None and t[1](record[t[0]]) for t in tests) for tests in rows)
    ]


def _numbers(values: list[Any]) -> list[float] | ErrorValue:
    out = []
    for v in values:
        if isinstance(v, ErrorValue):
            return v
        if type(v) is float:
            out.append(v)
    return out


def _database(name: str, reduce: Callable[[list[float]], Any], doc: str) -> None:
    @function(
        name,
        "raw",
        "raw",
        "raw",
        status="partial",
        note="computed criteria are not supported",
        category=DB,
    )
    def fn(ctx: Context, db: Any, field: Any, crit: Any) -> Any:
        values = _matching(db, field, crit, ctx)
        if isinstance(values, ErrorValue):
            return values
        if field is MISSING:
            return float(len(values)) if name == "DCOUNT" else VALUE
        nums = _numbers(values)
        return nums if isinstance(nums, ErrorValue) else reduce(nums)

    fn.__doc__ = doc


def _plain_sum(values: list[float]) -> float:
    total = 0.0
    for v in values:
        total += v
    return total


_database("DSUM", _plain_sum, "Sum of a field over the matching records.")
_database(
    "DAVERAGE",
    lambda v: _plain_sum(v) / len(v) if v else DIV0,
    "Mean of a field over the matching records.",
)
_database("DCOUNT", lambda v: float(len(v)), "Numbers in a field of the matching records.")
_database("DMAX", lambda v: max(v) if v else 0.0, "Largest number of a field, matching records.")
_database("DMIN", lambda v: min(v) if v else 0.0, "Smallest number of a field, matching records.")
