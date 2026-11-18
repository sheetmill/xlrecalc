"""A conditional format's formula, against capture ``cf_rules`` (Excel 16.115 for Mac).

Every case of ``capture/cf_probe.py`` is replayed: the rule's formula is parsed at its anchor with
``Formula.parse`` and evaluated at each cell read, with the case's constants, and the result is
compared with what Excel showed (its fill and bold applied or not). The anchor, the truth of a
result and the shape of a range or array result are the rules this capture settles; until it is
made, they are the ones below and the test is skipped.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from capture.cf_probe import CASES, OTHER, SHEET_OF, Case, Rule

from xlrecalc import Formula, UnsupportedError
from xlrecalc._compare import compare
from xlrecalc._values import ErrorValue

CAPTURE = Path(__file__).with_name("captures") / "cf_rules.tsv"
SKIPPED = {"refused", "unreadable", "kept", "recalculated"}
"""Rows not replayed here: a file Excel refused, a reading the control could not confirm, and
group G's opening rows (their own test)."""


def _rows() -> list[list[str]]:
    if not CAPTURE.is_file():
        return []
    lines = CAPTURE.read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith("#")]


def _anchor(sqref: str) -> str:
    """The cell relative references are measured from: the first area's top-left (provisional).

    A whole column (``C:C``) starts at row 1.
    """
    corner = sqref.split()[0].split(":")[0]
    return corner if corner[-1].isdigit() else f"{corner}1"


def _truth(value: Any) -> bool:
    """Whether a result applies the format (provisional): TRUE or a nonzero number; a range or
    array by its first element."""
    while isinstance(value, tuple):
        value = value[0] if value else False
    if isinstance(value, bool):
        return value
    if isinstance(value, float):
        return value != 0
    return False


def _values(case: Case) -> dict[str, Any]:
    sheet = SHEET_OF[case.name]
    values: dict[str, Any] = {f"{sheet}!{ref}": v for ref, v in case.cells.items()}
    values |= {f"{OTHER}!A1": 5.0, "ctl!C1": 5.0}
    return values


def _names(case: Case) -> dict[str, Any]:
    return {name: f"={text}" for name, text in {**case.names, **case.local_names}.items()}


def _applies(case: Case, rule: Rule, cell: str) -> bool:
    sheet = SHEET_OF[case.name]
    values, names = _values(case), _names(case)

    def at(formula: str) -> Any:
        parsed = Formula.parse(formula, sheet=sheet, origin=_anchor(rule.sqref))
        return parsed.evaluate(values, at=cell, names=names)

    if rule.kind == "cellIs":
        own = values.get(f"{sheet}!{cell}", 0.0)
        bounds = [at(f) for f in rule.formulas]
        if rule.operator == "between":
            low, high = sorted(bounds, key=lambda b: (isinstance(b, str), b))  # type: ignore[arg-type]
            return compare(own, low, ">=") is True and compare(own, high, "<=") is True
        op = {"greaterThan": ">", "lessThan": "<", "equal": "=", "notEqual": "<>"}[
            rule.operator or ""
        ]
        return compare(own, bounds[0], op) is True  # type: ignore[arg-type]
    result = at(rule.formulas[0])
    return not isinstance(result, ErrorValue) and _truth(result)


@pytest.mark.skipif(not CAPTURE.is_file(), reason="capture cf_rules not made yet")
@pytest.mark.parametrize("row", _rows(), ids=lambda r: f"{r[0]}:{r[4]}")
def test_a_rule_applies_where_excel_applies_it(row: list[str]) -> None:
    name, outcome, cell = row[0], row[8], row[4]
    if outcome in SKIPPED:
        pytest.skip(outcome)
    case = next(c for c in CASES if c.name == name)
    rules = [r for r in case.rules if cell in _cells(r.sqref)]
    try:
        got = any(_applies(case, rule, cell) for rule in rules)
    except UnsupportedError:
        got = False
    assert got is (outcome == "applied"), (name, cell, outcome)


def _cells(sqref: str) -> set[str]:
    from capture.cf_probe import _expand

    return {cell for area in sqref.split() for cell in _expand(area)}
