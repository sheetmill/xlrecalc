"""One formula evaluated outside a workbook: ``xlrecalc.evaluate`` and ``xlrecalc.Formula``.

A captured formula evaluated alone, its inputs read through a resolver, gives what the engine
gives in the capture's sheet (which ``test_captures.py`` compares with Excel's value), and no
function raises on a captured row: in every suite whose inputs are values (not formulas or fixed
cells), unsupported and volatile formulas left out; every 25th row by default, every row in the
slow tier.
"""

from __future__ import annotations

import re

import pytest
from tools.excel_names import module_text

import xlrecalc
from tests.captures import Formula as CaptureFormula
from tests.captures import load as load_suite
from tests.test_captures import SUITES, formula_text, run, same
from xlrecalc import ErrorValue, Formula, FormulaError, Reference, UnsupportedError, evaluate
from xlrecalc._functions._excel_names import EXCEL_NAMES
from xlrecalc._store import COLS


def test_values_by_cell_and_range() -> None:
    assert evaluate("=A1+B1", {"A1": 1, "B1": 2}) == 3.0
    assert evaluate("SUM(A1:A3)", {"A1:A3": [[1], [2], [3]]}) == 6.0
    assert evaluate("=Sheet2!B2*2", {"Sheet2!B2": 10}) == 20.0
    assert evaluate("='My Sheet'!B2", {"'My Sheet'!B2": "x"}) == "x"
    assert evaluate("=A1", {"a1": True}) is True  # keys ignore case


def test_a_missing_reference_is_blank() -> None:
    """A cell given no value is blank, as an empty cell is in Excel."""
    assert evaluate("=A1") == 0.0
    assert evaluate('=A1&""') == ""
    assert evaluate("=ISBLANK(A1)") is True


def test_forms_accepted() -> None:
    """As typed (``=``: the formula bar's form) or in file form (no ``=``: as a file stores it)."""
    for text in ("=XLOOKUP(2,A1:A2,B1:B2)", "_xlfn.XLOOKUP(2,A1:A2,B1:B2)"):
        assert evaluate(text, {"A1:B2": [[1, "a"], [2, "b"]]}) == "b"
    assert Formula.parse("=XLOOKUP(1,A1,B1)").text == "_xlfn.XLOOKUP(1,A1,B1)"
    assert evaluate("=1.2345678901234567") == 1.23456789012345  # typed: the first 15 digits
    assert evaluate("1.2345678901234567E-300*1") == 1.23456789012345e-300  # capture literals


def test_dynamic_by_default_legacy_on_request() -> None:
    values = {"A1:A3": [[1], [2], [3]]}
    assert evaluate("=A1:A3*2", values) == ((2.0,), (4.0,), (6.0,))
    assert evaluate("=A1:A3*2", values, legacy=True, at="B2") == 4.0  # implicit intersection
    assert evaluate("=SEQUENCE(2,2)") == ((1.0, 2.0), (3.0, 4.0))


def test_relative_references_move_with_the_cell() -> None:
    """A conditional format's formula, written for A1, evaluated at each cell of its range."""
    rule = Formula.parse("=A1>5", origin="A1")
    values = {"A1:A3": [[3], [9], [6]]}
    assert [rule.evaluate(values, at=f"A{r}") for r in (1, 2, 3)] == [False, True, True]
    assert rule.references == (Reference(None, 1, 1, 1, 1),)
    assert evaluate("=ROW()+COLUMN()", at="C5") == 8.0


def test_resolver_gets_each_reference_once() -> None:
    asked: list[str] = []

    def resolver(ref: Reference) -> object:
        asked.append(ref.a1)
        return [[1], [2], [3]] if not ref.is_cell else 7

    assert evaluate("=SUM(A:A)+B1+B1", resolver=resolver) == 20.0
    assert asked == ["A1:A1048576", "B1"]
    assert (
        evaluate('=INDIRECT("Data!C3")', resolver=lambda r: 42 if r.a1 == "'Data'!C3" else None)
        == 42.0
    )


def test_the_mapping_comes_before_the_resolver() -> None:
    assert evaluate("=SUM(A1:A2)", {"A1": 100}, resolver=lambda r: [[1], [2]]) == 102.0


def test_names() -> None:
    assert evaluate("=B2*Rate", {"B2": 10}, names={"Rate": 0.5}) == 5.0
    assert evaluate("=Twice(4)", names={"Twice": "=LAMBDA(x,x*2)"}) == 8.0
    assert evaluate("=Missing") == ErrorValue("#NAME?")


def test_errors_are_values() -> None:
    assert evaluate("=1/0") == ErrorValue("#DIV/0!")
    assert evaluate("=A1#") == ErrorValue("#REF!")  # no dynamic array at A1


def test_a_name_excel_does_not_know_is_name_error() -> None:
    """As in Excel: ``#NAME?``, even in an untaken branch it is not reached."""
    assert evaluate("=NOSUCHFN(1)") == ErrorValue("#NAME?")
    assert evaluate('=JIS("a")') == ErrorValue("#NAME?")  # no en-US Excel knows it (function_names)
    assert evaluate("=IF(TRUE,2,NOSUCHFN())") == 2.0
    assert Formula.parse("=NOSUCHFN(1)").unsupported == frozenset()


def test_a_function_excel_has_and_the_package_does_not_compute_raises() -> None:
    f = Formula.parse('=RTD("a","","b")+1')
    assert f.unsupported == frozenset({"RTD"})
    with pytest.raises(UnsupportedError, match="RTD"):
        f.evaluate()
    with pytest.raises(UnsupportedError, match="structured reference"):
        evaluate("=Table1[Price]")
    assert issubclass(UnsupportedError, FormulaError)


def test_a_union_is_computed_where_a_capture_shows_excel_reading_one() -> None:
    values = {"A1": 1.0, "A2": 2.0, "C1": 10.0}
    assert evaluate("=SUM((A1,A2,C1))", values) == 13.0
    assert evaluate("=INDEX((A1:A2,C1),1,1,2)", values) == 10.0
    f = Formula.parse('=CELL("address",(A1,A2))')
    assert f.unsupported == frozenset()  # found while evaluating, as for LINEST
    with pytest.raises(UnsupportedError, match=re.escape("a union of ranges in CELL")):
        f.evaluate(values)


@pytest.mark.parametrize(
    ("text", "label"),
    [
        ('=CELL("format",A1)', 'CELL("format")'),
        ('=CELL("Width",A1)', 'CELL("width")'),
        ('=CELL("color",A1)', 'CELL("color")'),
        ("=Table1[Price]", "structured reference"),
        ("=Sheet1:Sheet2!A1", "external or 3-D reference"),
        ("=[1]Sheet1!A1", "external or 3-D reference"),
    ],
)
def test_an_unsupported_form_is_named_and_raises(text: str, label: str) -> None:
    """What the workbook's report names as not computed is named beforehand, and raises."""
    f = Formula.parse(text)
    assert label in f.unsupported
    with pytest.raises(UnsupportedError, match=re.escape(label)):
        f.evaluate({"A1": 1, "A2": 2})


@pytest.mark.parametrize(
    ("text", "label"),
    [
        ("=CELL(B1,A1)", 'CELL("format")'),
        ("=LINEST(A1:A3,B1:C3)", "LINEST with several regressors"),
    ],
)
def test_an_unsupported_case_found_only_by_its_values_raises(text: str, label: str) -> None:
    """``CELL`` given its kind in a cell is not known beforehand; evaluating raises, not #NAME?."""
    f = Formula.parse(text)
    assert not f.unsupported
    with pytest.raises(UnsupportedError, match=re.escape(label)):
        f.evaluate({"A1": 1, "A2": 2, "B1": "format"})


def test_computed_cases_beside_the_unsupported_ones_are_values() -> None:
    assert evaluate('=CELL("row",B7)') == 7.0
    assert evaluate("=ROW((A1,A2))") == ErrorValue.REF  # of several areas
    assert evaluate("=NOSUCHFN(1)") == ErrorValue.NAME


def test_text_the_formula_bar_refuses() -> None:
    for text in ("=SUM(", "=1+", "=)"):
        with pytest.raises(FormulaError):
            evaluate(text)


def test_properties() -> None:
    f = Formula.parse("=SUM(Sheet2!A1:B2)+RAND()*Rate", origin="C3")
    assert f.references == (Reference("Sheet2", 1, 1, 2, 2),)
    assert f.names == ("Rate",)
    assert {"SUM", "RAND"} <= f.functions
    assert f.volatile
    assert Reference("Sheet2", 1, 1, 2, 2).a1 == "'Sheet2'!A1:B2"


def test_excel_names_are_derived_from_the_capture_list() -> None:
    """``tools/excel_names.py`` writes the module from ``capture/functions_list.py``."""
    from pathlib import Path

    module = Path(xlrecalc.__file__).parent / "_functions" / "_excel_names.py"
    assert module.read_text(encoding="utf-8") == module_text()
    assert {"SUM", "BESSELJ", "XLOOKUP", "LAMBDA"} <= EXCEL_NAMES


def _agrees(name: str, every: int) -> list[str]:
    """Rows of a suite where the ad hoc result differs from the engine's in the sheet."""
    suite = load_suite(name)
    if suite.fixed or any(isinstance(v, CaptureFormula) for r in suite.rows for v in r.inputs):
        return []  # inputs computed by formulas of their own: not a single formula
    engine, s = run(suite)
    sheet = engine.sheets[s]

    def resolver(ref: Reference) -> object:
        if ref.sheet is not None:
            return None
        return [
            [
                None if (v := sheet.values.get((r - 1) * COLS + c - 1)) is None or c == 1 else v
                for c in range(ref.first_col, min(ref.last_col, 40) + 1)
            ]
            for r in range(ref.first_row, min(ref.last_row, len(suite.rows) + 2) + 1)
        ]

    # A function that raises on a captured row reaches its value by accident (#VALUE!).
    differ = [f"{name} {c}: {why}" for c, why in engine.report.exceptions.items()]
    for k, row in enumerate(suite.rows, 1):
        if k % every:
            continue
        template = engine.templates[sheet.templates[(k - 1) * COLS]]
        if template.unsupported or template.volatile:
            continue
        text = formula_text(row, k)
        try:
            f = Formula.parse(text, origin=f"A{k}", legacy=True)
        except FormulaError as exc:
            differ.append(f"{name} {k} {text}: {exc}")
            continue
        got = f.evaluate(resolver=resolver, date1904=suite.date1904)
        want = engine.get(s, k - 1, 0)
        if not same(got, want):
            differ.append(f"{name} {k} {text}: {got!r} != {want!r}")
    return differ


@pytest.mark.parametrize("name", SUITES)
def test_captured_formulas_alone_give_the_engines_value(name: str) -> None:
    """A sample of each suite (every 25th row); the slow tier runs every row."""
    differ = _agrees(name, 25)
    assert not differ, "\n".join(differ[:40])


@pytest.mark.slow
@pytest.mark.parametrize("name", SUITES)
def test_every_captured_formula_alone_gives_the_engines_value(name: str) -> None:
    differ = _agrees(name, 1)
    assert not differ, "\n".join(differ[:40])


def test_a_formula_evaluates_again_with_other_values() -> None:
    """Nothing from one evaluation (values, names, resolved cells) reaches the next."""
    f = Formula.parse("=A1+Rate")
    assert f.evaluate({"A1": 1}, names={"Rate": 10}) == 11.0
    assert f.evaluate() == ErrorValue("#NAME?")
    assert f.evaluate(resolver=lambda r: 5) == ErrorValue("#NAME?")
    assert f.evaluate(resolver=lambda r: 5, names={"Rate": 1}) == 6.0
    assert f.evaluate(names={"Rate": 1}) == 1.0
