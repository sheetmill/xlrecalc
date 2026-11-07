"""The evaluator (ARCHITECTURE.md §7): loading, the trust rule, scheduling, the report."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from tests.test_xlsx import package
from xlrecalc import recalculate
from xlrecalc._engine import Engine
from xlrecalc._functions import REGISTRY, function
from xlrecalc._load import load, refused_in_file
from xlrecalc._values import BLANK, ErrorValue

NA = ErrorValue.NA


def cells(**sheet: object) -> Engine:
    return Engine.from_cells({"S": dict(sheet)})


def value(engine: Engine, ref: str) -> object:
    from xlrecalc._formula._template import cell_axes

    r, c = cell_axes(ref, 0, 0)
    return engine.get(0, r.value, c.value)


# The text "#N/A" and the error #N/A stay distinct (§4.2) -------------------------------------


def test_text_and_error_in_operators():
    e = cells(A1="#N/A", A2="=NA()", B1='=A1="#N/A"', B2='=A2="#N/A"', B3="=A1=A2")
    e.recalculate()
    assert value(e, "B1") is True
    assert value(e, "B2") == NA
    assert isinstance(value(e, "B2"), ErrorValue)
    assert value(e, "B3") == NA
    assert isinstance(value(e, "B3"), ErrorValue)


def test_text_and_error_in_match_and_countif():
    e = cells(
        A1="#N/A",
        A2="=NA()",
        A3=1,
        B1='=MATCH("#N/A",A1:A3,0)',
        B2='=COUNTIF(A1:A3,"#N/A")',
        B3="=COUNTIF(A1:A3,A1)",
        B4='=COUNTIF(A1:A3,"*")',
        B5="=MATCH(NA(),A1:A3,0)",
    )
    e.recalculate()
    assert value(e, "B1") == 1.0  # MATCH finds the text only
    assert value(e, "B2") == 1.0  # the criterion "#N/A" counts the error, not the text
    assert value(e, "B3") == 1.0
    assert value(e, "B4") == 1.0  # the text is text, the error is not
    assert value(e, "B5") == NA


# Loading and the trust rule (§7.4) -------------------------------------------------------------

ROWS = (
    '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v>999</v></c>'
    '<c r="C1"><f>B1+1</f><v>777</v></c><c r="D1"><f>A1+1</f></c>'
    '<c r="E1"><f>D1*10</f><v>555</v></c></row>'
)


def test_cached_values_are_trusted():
    e = load(io.BytesIO(package(ROWS)))
    e.recalculate()
    assert value(e, "B1") == 999.0  # trusted: the file's cached value
    assert value(e, "C1") == 777.0
    assert value(e, "D1") == 3.0  # no cached value: computed
    assert value(e, "E1") == 30.0  # depends on a stale cell: computed


def test_full_calculation_on_load_or_on_request():
    e = load(io.BytesIO(package(ROWS)), trust_cached=False)
    e.recalculate()
    assert (value(e, "B1"), value(e, "C1")) == (6.0, 7.0)
    flagged = package(
        ROWS,
        sheets='<sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
        '<calcPr calcId="191029" fullCalcOnLoad="1"/><sheets>',
    )
    e = load(io.BytesIO(flagged))
    e.recalculate()
    assert (value(e, "B1"), value(e, "C1")) == (6.0, 7.0)


def test_volatile_formulas_are_always_stale():
    e = load(io.BytesIO(package('<row r="1"><c r="A1"><f>TODAY()</f><v>1</v></c></row>')))
    e.recalculate()
    assert value(e, "A1") > 45000


def test_shared_formulas_take_the_masters_template():
    rows = (
        '<row r="1"><c r="A1"><v>1</v></c>'
        '<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f></c></row>'
        '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/></c></row>'
        '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/></c></row>'
    )
    e = load(io.BytesIO(package(rows)))
    e.recalculate()
    assert [value(e, f"B{r}") for r in (1, 2, 3)] == [2.0, 4.0, 6.0]
    assert len(e.templates) == 1


def test_refused_formulas_keep_their_cached_value():
    rows = (
        '<row r="1"><c r="A1"><f>#SPILL!</f><v>102</v></c><c r="B1"><f>A1+1</f></c>'
        '<c r="C1"><f>-\t1</f><v>5</v></c></row>'
    )
    e = load(io.BytesIO(package(rows)))
    e.recalculate()
    assert value(e, "A1") == 102.0
    assert value(e, "B1") == 103.0
    assert value(e, "C1") == 5.0
    assert set(e.report.refused) == {("Data", 1, 1), ("Data", 1, 3)}


def test_a_refused_formula_without_a_cached_value_is_blank():
    """As Excel's repair leaves it (capture repair_uncached)."""
    rows = '<row r="1"><c r="A1"><f>#SPILL!</f><v/></c><c r="B1"><f>ISBLANK(A1)</f></c></row>'
    e = load(io.BytesIO(package(rows)))
    e.recalculate()
    assert value(e, "A1") is BLANK
    assert value(e, "B1") is True


@pytest.mark.parametrize(
    ("text", "refused"),
    [
        ("1+1", False),
        ("#N/A", False),
        ("#SPILL!", True),
        ("#CALC!", True),
        ("A1#", True),
        ("-\t1", True),
        ("SUM( )", True),
        ("SUM()", False),
        ("SUM( A1 )", False),
        ("AREAS(1)", True),
        ('_xlfn.ISFORMULA("a")', True),
        ("_xlfn.FORMULATEXT({1,2})", True),
        ("AREAS(#N/A)", True),
        ("AREAS(A1)", False),
        ("_xlfn.ISFORMULA(#REF!)", False),
        ("_xlfn.FORMULATEXT((A1,B1))", False),
    ],
)
def test_refused_in_file(text, refused):
    assert (refused_in_file(text) is not None) is refused


def test_legacy_array_formulas_are_evaluated_and_fill_their_range():
    rows = (
        '<row r="1"><c r="A1"><f t="array" ref="A1:A3">B1:B2*2</f></c><c r="B1"><v>1</v></c>'
        '</row><row r="2"><c r="B2"><v>3</v></c></row>'
        '<row r="4"><c r="A4"><f t="array" ref="A4">SUM(B1:B2*2)</f></c></row>'
    )
    e = load(io.BytesIO(package(rows)))
    e.recalculate()
    assert [value(e, f"A{r}") for r in (1, 2, 3, 4)] == [2.0, 6.0, ErrorValue.NA, 8.0]
    assert not e.report.unsupported


# Scheduling and the report (§7.1, §7.2, §7.6) -------------------------------------------------


def test_long_chain_is_iterative():
    sheet = {"A1": 1}
    for r in range(2, 20001):
        sheet[f"A{r}"] = f"=A{r - 1}+1"
    e = Engine.from_cells({"S": sheet})
    e.recalculate()
    assert value(e, "A20000") == 20000.0


def test_circular_references_and_their_dependents_are_empty_text():
    e = cells(A1="=B1+1", B1="=C1+1", C1="=A1+1", D1="=A1*2", E1="=D1+F1", F1=1, G1="=F1+1")
    report = e.recalculate()
    on_or_after = {("S", 1, c) for c in (1, 2, 3, 4, 5)}
    assert report.circular == on_or_after
    for ref in ("A1", "B1", "C1", "D1", "E1"):
        assert value(e, ref) == ""
    assert value(e, "G1") == 2.0


def test_reading_a_cycle_first_settles_its_dependents():
    e = cells(A1="=A1+1", B1="=A1*2")
    assert value(e, "A1") == ""  # evaluates only what A1 needs, and finds the cycle
    assert value(e, "B1") == ""


def test_unsupported_functions_and_their_dependents():
    e = cells(A1="=NOSUCHFUNCTION(1)", B1="=A1+1", C1="=2")
    report = e.recalculate()
    assert value(e, "A1") == ErrorValue.NAME
    assert report.unsupported["NOSUCHFUNCTION"] >= 1
    assert {("S", 1, 1), ("S", 1, 2)} <= report.uncomputed
    assert ("S", 1, 3) not in report.uncomputed


def test_an_internal_error_is_contained():
    @function("_TESTBOOM", "number")
    def _boom(ctx, x):
        raise RuntimeError("boom")

    try:
        e = cells(A1="=_TESTBOOM(1)", B1="=1+1")
        report = e.recalculate()
        assert value(e, "A1") == ErrorValue.VALUE
        assert value(e, "B1") == 2.0
        assert "RuntimeError: boom" in report.exceptions[("S", 1, 1)]
    finally:
        REGISTRY.pop("_TESTBOOM")


def test_changing_a_value_recomputes_its_dependents():
    e = cells(A1=1, B1="=A1*2", C1="=B1+A1", D1="=7")
    e.recalculate()
    assert value(e, "C1") == 3.0
    e.set_value(0, 0, 0, 10.0)
    assert value(e, "C1") == 30.0
    assert value(e, "B1") == 20.0


def test_defined_names():
    e = cells(A1=5, A2=6, B1="=SUM(Data)", B2="=Rate*2")
    e.define_name("Data", "S!$A$1:$A$2")
    e.define_name("Rate", "0.5")
    e.recalculate()
    assert value(e, "B1") == 11.0
    assert value(e, "B2") == 1.0


def test_implicit_intersection_and_top_level_snapping():
    e = cells(A1=1, A2=2, A3=3, B2="=A:A*10", B3="=0.1+0.2-0.3", B4="=(0.1+0.2-0.3)")
    e.recalculate()
    assert value(e, "B2") == 20.0
    assert value(e, "B3") == 0.0
    assert value(e, "B4") == pytest.approx(5.55e-17)


def test_blank_results_are_zero():
    e = cells(B1="=A1", B2="=IF(TRUE,A1)")
    e.recalculate()
    assert value(e, "B1") == 0.0
    assert value(e, "A1") is BLANK


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=SUMPRODUCT(--(A1:A3>0))", 3.0),  # an array argument: the range is an array
        ("=SUM((A1:A3)*2)", 2.0),  # elsewhere an operator takes the implicit intersection
        ("=IF(A:A>0,1)", 1.0),
        ("=SUM({1,2}*2)", 6.0),  # an array constant is an array anywhere
        ("=SUM(A1:A3)*2", 12.0),
        ("=INDEX(A1:A3*2,2)+1", 5.0),
        ("=SUM(A1:A3,-A1)", 5.0),
        ("=SUM(ABS({-1,-2}))", 3.0),  # a one-value function over an array, element by element
    ],
)
def test_legacy_formulas_over_ranges_and_arrays(formula, expected):
    """As Excel computes a legacy formula (capture arrays); the formula is in B1, row 1."""
    e = cells(A1=1, A2=2, A3=3, B1=formula)
    e.recalculate()
    assert value(e, "B1") == expected


def test_a_changed_formula_makes_its_dependents_stale_for_reads():
    e = cells(A1=2, B1="=A1*3", C1="=B1+1")
    assert value(e, "C1") == 7.0
    e.set_formula(0, 0, 1, "A1*10")
    assert value(e, "C1") == 21.0  # read directly, without recalculate()


# What is left gives an error by design: no condition met, no case matched, an array in one cell.
_BASELINE_IS_AN_ERROR = ("_xlfn.IFS(FALSE", "_xlfn.SWITCH(2,1", "TRANSPOSE(")


def test_refused_in_file_agrees_with_excel_on_empty_arguments():
    """Every formula with an empty argument Excel refused is refused, and none it computed."""
    captures = Path(__file__).parent / "captures"
    refused = (captures / "empty_args_refused.tsv").read_text(encoding="utf-8").split("\n")[1:]
    assert [f for f in refused if f and refused_in_file(f) is None] == []
    rows = (captures / "empty_args.tsv").read_text(encoding="utf-8").split("\n")
    computed = [r.split("\t")[3] for r in rows if r.startswith("empty_args\t")]
    assert len(computed) > 500
    assert [f for f in computed if refused_in_file(f) is not None] == []


def test_a_lone_surrogate_is_written_as_excel_stores_it(tmp_path):
    """``MID`` can cut a character in two; the file holds ``_xDE00_``, as Excel saves it."""
    openpyxl = pytest.importorskip("openpyxl")
    book = openpyxl.Workbook()
    book.active["B1"] = "\U0001f600ab"
    book.active["A1"] = "=MID(B1,2,2)"
    source, out = tmp_path / "in.xlsx", tmp_path / "out.xlsx"
    book.save(source)
    recalculate(source, out, full=True)
    with zipfile.ZipFile(out) as written:
        assert "<v>_xDE00_a</v>" in written.read("xl/worksheets/sheet1.xml").decode("utf-8")


def test_every_empty_argument_row_has_a_baseline_that_is_no_error():
    """A row can only show what its empty argument does if the rest of the call is valid."""
    rows = (Path(__file__).parent / "captures" / "empty_args.tsv").read_text(encoding="utf-8")
    baselines = {}
    for row in rows.split("\n"):
        if row.startswith("empty_args\t"):
            fields = row.split("\t")
            if ",," not in fields[3] and "(," not in fields[3] and ",)" not in fields[3]:
                baselines.setdefault(fields[1], []).append((fields[3], fields[5], fields[6]))
    assert len(baselines) > 300
    errors = [
        f for rows in baselines.values() if all(r[1] == "e" for r in rows) for f, _, _ in rows
    ]
    assert [f for f in errors if not f.startswith(_BASELINE_IS_AN_ERROR)] == []
