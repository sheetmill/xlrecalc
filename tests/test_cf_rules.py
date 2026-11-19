"""A conditional format's formula, against capture ``cf_rules`` (Excel 16.115 for Mac).

Every case of ``capture/cf_probe.py`` is replayed: the rule's formula is parsed at its anchor with
``Formula.parse`` and evaluated at each cell read, with the case's constants, and the result is
compared with what Excel showed (its fill and bold applied or not). What the capture settles:

- the anchor is the top-left corner of the box around all the rule's areas, whatever their
  order (group A);
- the rule is evaluated as a formula is today, not as a legacy one, and a range or array result
  counts by its top-left element (group C);
- the format applies where ``IF`` would take its true branch on that element (``to_bool``: TRUE,
  a nonzero number, the text ``"TRUE"`` in any case); an error, other text and a blank do not
  (group B).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from capture.cf_probe import CASES, OTHER, SHEET_OF, Case, Rule

from xlrecalc import Formula, UnsupportedError
from xlrecalc._coerce import to_bool
from xlrecalc._compare import compare
from xlrecalc._values import BLANK

CAPTURE = Path(__file__).with_name("captures") / "cf_rules.tsv"
SKIPPED = {"refused", "unreadable", "kept", "recalculated", "mixed"}
"""Rows not replayed here: a file Excel refused, a reading the control could not confirm, group
G's opening rows and the two stacked rules (their own tests)."""


def _rows() -> list[list[str]]:
    if not CAPTURE.is_file():
        return []
    lines = CAPTURE.read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith("#")]


def _anchor(sqref: str) -> str:
    """The top-left corner of the box around every area of the range (a whole column from row 1)."""
    rows, cols = [], []
    for area in sqref.split():
        for corner in area.split(":"):
            letters = corner.rstrip("0123456789")
            cols.append(letters)
            rows.append(int(corner[len(letters) :] or 1))
    return f"{min(cols, key=lambda c: (len(c), c))}{min(rows)}"


def _truth(value: Any) -> bool:
    """Whether a result applies the format: its top-left element, as ``IF``'s condition."""
    while isinstance(value, tuple):
        value = value[0] if value else BLANK
    test = to_bool(BLANK if value is None else value, None)  # type: ignore[arg-type]
    return test is True


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
    return _truth(at(rule.formulas[0]))


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


def _row(name: str) -> list[str]:
    return next(r for r in _rows() if r[0] == name)


@pytest.mark.skipif(not CAPTURE.is_file(), reason="capture cf_rules not made yet")
def test_what_excel_refuses_and_stacks() -> None:
    """A file holding a rule with an array constant or a structured reference is one Excel offers
    to repair; a rule with ``stopIfTrue`` that applies keeps the next rule's format off."""
    refused = {r[0] for r in _rows() if r[8] == "refused"}
    assert refused == {
        "shape_array_false_first",
        "shape_array_true_first",
        "table_this_row",
        "table_column_sum",
    }
    stop = _row("stop_if_true")
    assert (stop[6], stop[7]) == ("255,255,255", "true")  # the first rule's bold, no fill


@pytest.mark.skipif(not CAPTURE.is_file(), reason="capture cf_rules not made yet")
def test_a_volatile_function_in_a_rule_recalculates_nothing_on_opening() -> None:
    """Planted cached values stay on opening and in Excel's saved copy, whatever the rule calls."""
    for name in ("open_rule_rand", "open_rule_offset", "open_rule_plain"):
        row = _row(name)
        assert (row[5], row[6], row[8]) == ("999 888", "A1=2 A2=999 A3=888", "kept"), row
