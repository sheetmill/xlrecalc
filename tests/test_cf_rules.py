"""A conditional format's formula, against capture ``cf_rules`` (Excel 16.115 for Mac).

Every case of ``capture/cf_probe.py`` is replayed through ``ConditionalFormula``: the rule is made
from its formula (or ``cellIs`` operator and bounds) and range as the file stores them, and asked
at each cell read, with the case's constants; the answer is compared with what Excel showed (its
fill and bold applied or not). What the capture settles, which the class holds:

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

from xlrecalc import ConditionalFormula, FormulaError, UnsupportedError

CAPTURE = Path(__file__).with_name("captures") / "cf_rules.tsv"
SKIPPED = {"refused", "unreadable", "kept", "recalculated", "mixed"}
"""Rows not replayed here: a file Excel refused, a reading the control could not confirm, group
G's opening rows and the two stacked rules (their own tests)."""


def _rows() -> list[list[str]]:
    if not CAPTURE.is_file():
        return []
    lines = CAPTURE.read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith("#")]


def _values(case: Case) -> dict[str, Any]:
    sheet = SHEET_OF[case.name]
    values: dict[str, Any] = {f"{sheet}!{ref}": v for ref, v in case.cells.items()}
    values |= {f"{OTHER}!A1": 5.0, "ctl!C1": 5.0}
    return values


def _names(case: Case) -> dict[str, Any]:
    return {name: f"={text}" for name, text in {**case.names, **case.local_names}.items()}


def _rule(case: Case, rule: Rule) -> ConditionalFormula:
    sheet = SHEET_OF[case.name]
    if rule.kind == "cellIs":
        return ConditionalFormula.cell_is(
            rule.operator or "", rule.formulas, sqref=rule.sqref, sheet=sheet
        )
    return ConditionalFormula.parse(rule.formulas[0], sqref=rule.sqref, sheet=sheet)


def _applies(case: Case, rule: Rule, cell: str) -> bool:
    return _rule(case, rule).applies(cell, _values(case), names=_names(case))


@pytest.mark.skipif(not CAPTURE.is_file(), reason="capture cf_rules not made yet")
@pytest.mark.parametrize("row", _rows(), ids=lambda r: f"{r[0]}:{r[4]}")
def test_a_rule_applies_where_excel_applies_it(row: list[str]) -> None:
    name, outcome, cell = row[0], row[8], row[4]
    if outcome in SKIPPED:
        pytest.skip(outcome)
    case = next(c for c in CASES if c.name == name)
    rules = [r for r in case.rules if cell in _cells(r.sqref)]
    got = any(_applies(case, rule, cell) for rule in rules)
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
    for case in (c for c in CASES if c.name in refused):
        with pytest.raises(FormulaError, match=r"array constant|structured reference"):
            _rule(case, case.rules[0])
    stop = _row("stop_if_true")
    assert (stop[6], stop[7]) == ("255,255,255", "true")  # the first rule's bold, no fill


@pytest.mark.skipif(not CAPTURE.is_file(), reason="capture cf_rules not made yet")
def test_a_volatile_function_in_a_rule_recalculates_nothing_on_opening() -> None:
    """Planted cached values stay on opening and in Excel's saved copy, whatever the rule calls."""
    for name in ("open_rule_rand", "open_rule_offset", "open_rule_plain"):
        row = _row(name)
        assert (row[5], row[6], row[8]) == ("999 888", "A1=2 A2=999 A3=888", "kept"), row


def test_the_anchor_is_the_corner_of_the_box_not_of_the_first_area() -> None:
    """``D2:D3 B4:C5`` and ``B4:C5 D2:D3`` both measure from B2 (capture ``cf_rules``, group A).

    The first area's corner (D2 or B4) would give other cells: ROW(A1)=1 marks row 2 only, and
    COLUMN(A1)=1 column B only.
    """
    for sqref in ("D2:D3 B4:C5", "B4:C5 D2:D3"):
        row = ConditionalFormula.parse("ROW(A1)=1", sqref=sqref)
        col = ConditionalFormula.parse("COLUMN(A1)=1", sqref=sqref)
        assert row.anchor == col.anchor == "B2"
        cells = ("D2", "D3", "B4", "C4", "B5", "C5")
        assert [c for c in cells if row.applies(c)] == ["D2"]
        assert [c for c in cells if col.applies(c)] == ["B4", "B5"]
    assert ConditionalFormula.parse("TRUE", sqref="D5 B2").anchor == "B2"
    assert ConditionalFormula.parse("TRUE", sqref="C:C").anchor == "C1"
    assert ConditionalFormula.parse("TRUE", sqref="3:4").anchor == "A3"


def test_mode_and_truth() -> None:
    """A dynamic array whose top-left element is taken, then ``IF``'s condition."""
    data = {"D1": 0.0, "D2": 1.0, "D3": 1.0}
    compared = ConditionalFormula.parse("$D$1:$D$3>0", sqref="C1:C3")
    assert [compared.applies(f"C{r}", data) for r in (1, 2, 3)] == [False] * 3  # no intersection
    at = ConditionalFormula.parse("_xlfn.SINGLE($D$1:$D$3)>0", sqref="C1:C3")
    assert [at.applies(f"C{r}", data) for r in (1, 2, 3)] == [False, True, True]
    for formula, shown in (
        ("TRUE", True),
        ("-1", True),
        ("1E-300", True),
        ('"TRUE"', True),
        ('"true"', True),
        ("FALSE", False),
        ("0", False),
        ('"1"', False),
        ('""', False),
        ("NA()", False),
        ("1/0", False),
        ("Z1", False),  # a blank cell
    ):
        assert ConditionalFormula.parse(formula, sqref="B1").applies("B1") is shown, formula


def test_what_the_class_refuses() -> None:
    with pytest.raises(FormulaError, match="array constant"):
        ConditionalFormula.parse("INDEX({1,2},1)>1", sqref="A1")
    with pytest.raises(FormulaError, match="structured reference"):
        ConditionalFormula.parse("[@Col]>1", sqref="A1")
    assert ConditionalFormula.parse('"{x}"="{x}"', sqref="A1").applies("A1")  # braces in text
    with pytest.raises(UnsupportedError):
        ConditionalFormula.parse('RTD("a","","b")>0', sqref="A1").applies("A1")
    rule = ConditionalFormula.parse("TRUE", sqref="A1:A3")
    with pytest.raises(ValueError, match="not in the range"):
        rule.applies("B1")
    with pytest.raises(ValueError, match="not a range"):
        ConditionalFormula.parse("TRUE", sqref="A1:B")
    with pytest.raises(ValueError, match="not a cellIs operator"):
        ConditionalFormula.cell_is("beginsWith", ["1"], sqref="A1")
    with pytest.raises(ValueError, match="takes 2"):
        ConditionalFormula.cell_is("between", ["1"], sqref="A1")


def test_cell_is() -> None:
    """Excel's comparison against bounds evaluated at the cell; an error compares as no format."""
    values = {"A1": 1.0, "A2": 3.0, "A3": "x"}
    rule = ConditionalFormula.cell_is("notBetween", ["4", "2"], sqref="A1:A4")
    assert [rule.applies(f"A{r}", values) for r in (1, 2, 3)] == [True, False, True]
    greater = ConditionalFormula.cell_is("greaterThan", ["B1*0+2"], sqref="A1:A3")
    assert [greater.applies(f"A{r}", values) for r in (1, 2, 3)] == [False, True, True]
    error = ConditionalFormula.cell_is("equal", ["NA()"], sqref="A1")
    assert error.applies("A1", values) is False
