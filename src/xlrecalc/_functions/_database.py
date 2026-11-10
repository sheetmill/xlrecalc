"""Database functions: the D functions (captures ``misc`` and ``database_family``).

A database is a range whose first row names its columns; a criteria range has names in its
first row and conditions below. A record matches when every condition of some criteria row
holds; a blank criteria row matches every record, and a condition under a name the database
lacks matches none. A text condition without an operator matches text that starts with it
(``Bolt`` matches ``Boltcutter``), as in Excel; ``=Bolt`` matches only ``Bolt``.
Conditions computed by a formula (a name the database lacks over a formula) are not supported:
the formula is reported (``Report.unsupported``) and the file's value kept.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from typing import Any

from xlrecalc._coerce import Context, scalar
from xlrecalc._compare import criterion
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._functions._statistical import _VARIANCE, _variance
from xlrecalc._values import BLANK, ErrorValue, Ref

VALUE, DIV0 = ErrorValue.VALUE, ErrorValue.DIV0
DB = "database"


def _key(v: Any) -> str | None:
    return v.casefold() if isinstance(v, str) else None


def _condition(value: Any, ctx: Context) -> Callable[[Any], bool]:
    if isinstance(value, str) and value and value[0] not in "<>=":
        value += "*"
    return criterion(value, ctx.quirks)


def _by_row(ref: Ref) -> dict[int, dict[int, Any]]:
    """A range's non-blank cells, as {row offset: {column offset: value}}, in row order."""
    rows: dict[int, dict[int, Any]] = {}
    for r, c, v in ref.cells():
        rows.setdefault(r, {})[c] = v
    return rows


def _matching(db: Any, field: Any, crit: Any, ctx: Context) -> tuple[list[Any], int] | ErrorValue:
    """The field's values in the records that match, and how many more match with it blank.

    This walks the non-blank cells: a blank record and a blank criteria row are each tested once
    and counted, so a range over whole columns costs what its cells cost. An error given as an
    argument is the result, the first in argument order (capture ``database_family``).
    """
    for arg in (db, field, crit):
        if isinstance(arg, ErrorValue):
            return arg
    if not isinstance(db, Ref) or not isinstance(crit, Ref):
        return VALUE
    if db.rows < 2 or crit.rows < 2:
        return VALUE
    table, conditions = _by_row(db), _by_row(crit)
    header = table.pop(0, {})
    names: dict[str | None, int] = {}
    for c in sorted(header):
        names.setdefault(_key(header[c]), c)
    if len(header) < db.columns:  # a blank name, as a condition under a blank name finds it
        names.setdefault(None, min(set(range(len(header) + 1)) - set(header)))
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
            column = names[f.casefold()]
        elif isinstance(f, float):
            k = math.trunc(f)
            if not 1 <= k <= db.columns:
                return VALUE
            column = k - 1
        else:
            return VALUE
    wanted = conditions.pop(0, {})
    rows: list[list[tuple[int, Callable[[Any], bool]] | None]] = []
    if len(conditions) < crit.rows - 1:
        rows.append([])  # a blank criteria row matches every record
    templates = ctx.sheets[crit.sheet].templates  # type: ignore[attr-defined]
    for r, crow in conditions.items():
        tests: list[tuple[int, Callable[[Any], bool]] | None] = []
        for c, value in crow.items():
            key = _key(wanted.get(c, BLANK))
            if key not in names and templates.at(crit.r1 + r, crit.c1 + c) is not None:
                return ctx.unsupported("a computed criterion")  # type: ignore[attr-defined,no-any-return]
            tests.append((names[key], _condition(value, ctx)) if key in names else None)
        rows.append(tests)

    def matches(record: dict[int, Any]) -> bool:
        return any(
            all(t is not None and t[1](record.get(t[0], BLANK)) for t in tests) for tests in rows
        )

    values, blanks = [], 0
    for record in table.values():
        if matches(record):
            if column is None:
                blanks += 1
            elif column in record:
                values.append(record[column])
            else:
                blanks += 1
    if matches({}):  # Excel counts blank records only up to the sheet's last used row
        columns = ctx.sheets[db.sheet].columns.values()  # type: ignore[attr-defined]
        used = max((filled[-1] for filled in columns if filled), default=-1)
        blanks += max(min(db.r2, used) - db.r1 - len(table), 0)
    return values, blanks


def _numbers(values: list[Any]) -> list[float] | ErrorValue:
    out = []
    for v in values:
        if isinstance(v, ErrorValue):
            return v
        if type(v) is float:
            out.append(v)
    return out


def _database(
    name: str, reduce: Callable[[list[float]], Any], doc: str, ulps: int = 0, reason: str = ""
) -> None:
    @function(
        name,
        "raw",
        "raw",
        "raw",
        status="partial",
        note="computed criteria are not supported",
        category=DB,
        ulps=ulps,
        ulps_reason=reason,
    )
    def fn(ctx: Context, db: Any, field: Any, crit: Any) -> Any:
        found = _matching(db, field, crit, ctx)
        if isinstance(found, ErrorValue):
            return found
        values, blanks = found
        if field is MISSING:
            return float(blanks) if name == "DCOUNT" else VALUE
        if name == "DCOUNT":  # it counts numbers, past any error
            return float(sum(type(v) is float for v in values))
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


def _product(values: list[float]) -> float | ErrorValue:
    if not values:
        return 0.0
    total = 1.0
    for v in values:
        total *= v
    return total if math.isfinite(total) else ErrorValue.NUM


def _spread(values: list[float], sample: bool, root: bool) -> float | ErrorValue:
    var = _variance(values, sample)
    return math.sqrt(var) if root and not isinstance(var, ErrorValue) else var


_database("DPRODUCT", _product, "Product of a field over the matching records.")
_database("DVAR", lambda v: _spread(v, True, False), "Sample variance.", 4, _VARIANCE)
_database("DVARP", lambda v: _spread(v, False, False), "Population variance.", 4, _VARIANCE)
_database("DSTDEV", lambda v: _spread(v, True, True), "Sample deviation.", 4, _VARIANCE)
_database("DSTDEVP", lambda v: _spread(v, False, True), "Population deviation.", 4, _VARIANCE)


_COMPUTED = "computed criteria are not supported"


@function("DCOUNTA", "raw", "raw", "raw", status="partial", note=_COMPUTED, category=DB)
def DCOUNTA(ctx: Context, db: Any, field: Any, crit: Any) -> Any:
    """Non-blank cells of a field in the matching records."""
    found = _matching(db, field, crit, ctx)
    if isinstance(found, ErrorValue):
        return found
    return float(found[1]) if field is MISSING else float(len(found[0]))


@function("DGET", "raw", "raw", "raw", status="partial", note=_COMPUTED, category=DB)
def DGET(ctx: Context, db: Any, field: Any, crit: Any) -> Any:
    """The field's value in the one matching record: ``#VALUE!`` for none, ``#NUM!`` for more."""
    found = _matching(db, field, crit, ctx)
    if isinstance(found, ErrorValue):
        return found
    if field is MISSING:
        return VALUE
    values, blanks = found
    count = len(values) + blanks
    if count != 1:
        return VALUE if count == 0 else ErrorValue.NUM
    return values[0] if values else 0.0
