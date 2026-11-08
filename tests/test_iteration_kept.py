"""Iterative calculation through files, ``load`` and edits (ARCHITECTURE.md §7.2).

The workbook: a loop of two cells (``A1 = B1+F1``, ``B1 = A1/2``), a dependent (``C1``), an
unrelated formula (``D1``). With iteration on, a recalculation that reaches the loop makes one run
from the values the cells hold, as Excel does on opening (captures ``iterate_open``,
``iterate_calc``, checked cell by cell in ``tests/test_iteration.py``). A run that trusts the
file's values evaluates nothing, changes nothing and reports nothing.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest

import xlrecalc
from tests.test_xlsx import package
from xlrecalc._engine import ITERATION_LIMIT
from xlrecalc._report import CellRef, column_letters
from xlrecalc._values import ErrorValue
from xlrecalc._xlsx import Book

FLAG = '<calcPr iterate="1" fullCalcOnLoad="1"/>'
TRUST = '<calcPr iterate="1"/>'
OFF = '<calcPr fullCalcOnLoad="1"/>'
LOOP = {"A1": 5.0, "B1": 7.0, "C1": 999.0}
LABEL = "iterative calculation"


def cell(address: str, formula: str | None, value: str | None) -> str:
    f = "" if formula is None else f"<f>{formula}</f>"
    v = "<v/>" if value is None and formula is not None else f"<v>{value}</v>"
    return f'<c r="{address}">{f}{v}</c>'


def book(tmp_path: Path, calc: str, cells: str | None = None) -> Path:
    """A two-cell loop (``A1``, ``B1``), a dependent (``C1``), an unrelated formula (``D1``)."""
    if cells is None:
        cells = (
            cell("A1", "B1+F1", "5")
            + cell("B1", "A1/2", "7")
            + cell("C1", "A1+100", "999")
            + cell("D1", "F1+1", "77")
            + cell("F1", None, "0")
        )
    path = tmp_path / "book.xlsx"
    path.write_bytes(package(f'<row r="1">{cells}</row>', workbook_extra=calc))
    return path


def written(path: Path) -> dict[str, object]:
    """The address of every cell of the file, with the value its XML holds."""
    with Book(path) as b:
        return {f"{column_letters(c.column)}{c.row}": c.value for c in b.cells("Data")}


def has_cached(path: Path, address: str) -> bool:
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    found = re.search(rf'<c r="{address}"[^>]*>(.*?)</c>', sheet)
    assert found is not None
    return re.search(r"<v>[^<]+</v>", found[1]) is not None


def run(a: float, b: float, c: float, f: float = 0.0, runs: int = 1) -> dict[str, float]:
    """The captured rule, restated: in place, row order, until no change reaches 0.001."""
    for _ in range(runs):
        for _ in range(100):
            na = b + f
            nb = na / 2
            nc = na + 100
            change = max(abs(na - a), abs(nb - b), abs(nc - c))
            a, b, c = na, nb, nc
            if change < 0.001:
                break
    return {"A1": a, "B1": b, "C1": c}


ONE_RUN = run(5.0, 7.0, 999.0)


def assert_computed(report: xlrecalc.Report) -> None:
    assert not report.uncomputed
    assert not report.circular
    assert LABEL not in report.unsupported


def test_a_file_asking_for_a_full_calculation(tmp_path: Path) -> None:
    source = book(tmp_path, FLAG)
    out = tmp_path / "out.xlsx"
    assert_computed(xlrecalc.recalculate(source, out))
    assert written(out) == {**ONE_RUN, "D1": 1.0, "F1": 0.0}


def test_full_recalculation_runs_once_from_the_cached_values(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    out = tmp_path / "out.xlsx"
    assert_computed(xlrecalc.recalculate(source, out, full=True))
    assert written(out) == {**ONE_RUN, "D1": 1.0, "F1": 0.0}


def test_a_trusting_run_changes_nothing_and_claims_nothing(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    out = tmp_path / "out.xlsx"
    assert_computed(xlrecalc.recalculate(source, out))
    assert written(out) == {**LOOP, "D1": 77.0, "F1": 0.0}


def test_each_run_of_a_written_file_starts_from_what_it_holds(tmp_path: Path) -> None:
    """The first run writes ``fullCalcOnLoad``, so each later one evaluates every formula."""
    path = book(tmp_path, TRUST)
    xlrecalc.recalculate(path)
    assert written(path)["A1"] == 5.0
    assert_computed(xlrecalc.recalculate(path))
    assert written(path) == {**ONE_RUN, "D1": 1.0, "F1": 0.0}
    xlrecalc.recalculate(path)
    assert written(path) == {**run(5.0, 7.0, 999.0, runs=2), "D1": 1.0, "F1": 0.0}


@pytest.mark.parametrize(
    ("calc", "trust"),
    [(TRUST, False), (FLAG, True), (FLAG, False)],
    ids=["no_trust", "flag", "both"],
)
def test_value_gives_one_run(tmp_path: Path, calc: str, trust: bool) -> None:
    wb = xlrecalc.load(book(tmp_path, calc), trust_cached=trust)
    wb.recalculate()
    assert {a: wb.value(a) for a in ONE_RUN} == ONE_RUN
    assert wb.value("D1") == 1.0
    assert_computed(wb.report)


@pytest.mark.parametrize("order", [("C1", "A1", "B1"), ("B1", "C1", "A1"), ("A1", "C1", "B1")])
def test_value_on_demand_in_any_order(tmp_path: Path, order: tuple[str, str, str]) -> None:
    wb = xlrecalc.load(book(tmp_path, FLAG))
    assert {a: wb.value(a) for a in order} == ONE_RUN
    assert_computed(wb.report)


def test_a_trusting_load_gives_the_cached_values_without_evaluating(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, TRUST))
    assert {a: wb.value(a) for a in LOOP} == LOOP
    assert not wb.report.uncomputed


def test_each_recalculation_is_a_run(tmp_path: Path) -> None:
    """As Excel's ``calculate``: the loop's cells are calculated on every recalculation."""
    wb = xlrecalc.load(book(tmp_path, FLAG))
    wb.recalculate()
    wb.recalculate()
    assert {a: wb.value(a) for a in LOOP} == run(5.0, 7.0, 999.0, runs=2)
    wb.recalculate(full=True)
    assert {a: wb.value(a) for a in LOOP} == run(5.0, 7.0, 999.0, runs=3)
    assert_computed(wb.report)


def test_an_edit_runs_once_from_the_cached_values(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    wb = xlrecalc.load(source)
    wb.set("F1", 1.0)
    expected = run(5.0, 7.0, 999.0, f=1.0)
    assert {a: wb.value(a) for a in LOOP} == expected
    assert wb.value("D1") == 2.0
    assert_computed(wb.report)
    out = tmp_path / "out.xlsx"
    wb.save(out)  # saving recalculates, so the loop runs once more, as Excel's saving does
    assert written(out) == {**run(5.0, 7.0, 999.0, f=1.0, runs=2), "D1": 2.0, "F1": 1.0}


def test_a_formula_set_into_the_loop(tmp_path: Path) -> None:
    """It reads ``B1``'s value, not its own, so its missing cached value is never read."""
    wb = xlrecalc.load(book(tmp_path, TRUST))
    wb.set_formula("A1", "=B1+F1+1")
    a, b = wb.value("A1"), wb.value("B1")
    assert isinstance(a, float)
    assert isinstance(b, float)
    assert abs(a - 2.0) < 0.002
    assert b == a / 2
    assert_computed(wb.report)


def test_an_edit_breaking_the_loop_computes_its_cells(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, FLAG))
    wb.recalculate()
    wb.set("A1", 10.0)
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == [10.0, 5.0, 110.0]
    assert_computed(wb.report)
    out = tmp_path / "out.xlsx"
    wb.save(out)
    assert written(out) == {"A1": 10.0, "B1": 5.0, "C1": 110.0, "D1": 1.0, "F1": 0.0}


def test_a_loop_without_cached_values_ends_at_value_error(tmp_path: Path) -> None:
    """As in Excel (capture ``iterate_open``): what nothing calculated yet reads is ``#VALUE!``."""
    cells = cell("A1", "B1+0", None) + cell("B1", "A1/2", None) + cell("C1", "A1+100", None)
    source = book(tmp_path, FLAG, cells)
    wb = xlrecalc.load(source)
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == [ErrorValue.VALUE] * 3
    assert_computed(wb.report)


def test_a_volatile_cell_in_a_loop_makes_a_trusting_run_reach_it(tmp_path: Path) -> None:
    cells = cell("A1", "B1+RAND()*0", "5") + cell("B1", "A1/2", "7") + cell("C1", "A1+100", "999")
    source = book(tmp_path, TRUST, cells)
    out = tmp_path / "out.xlsx"
    assert_computed(xlrecalc.recalculate(source, out))
    assert written(out) == ONE_RUN


def test_an_array_formula_depending_on_a_loop_is_iterated_with_it(tmp_path: Path) -> None:
    cells = (
        cell("A1", "B1+0", "5")
        + cell("B1", "A1/2", "7")
        + '<c r="C1"><f t="array" ref="C1:E1">A1*{1,2,3}</f><v>11</v></c>'
        + '<c r="D1"><v>22</v></c><c r="E1"><v>33</v></c>'
        + cell("F1", "D1+1", "44")
    )
    wb = xlrecalc.load(book(tmp_path, FLAG, cells))
    wb.recalculate()
    a = wb.value("A1")
    assert isinstance(a, float)
    assert [wb.value(x) for x in ("C1", "D1", "E1", "F1")] == [a, 2 * a, 3 * a, 2 * a + 1]
    assert_computed(wb.report)


def test_a_loop_through_a_branch_not_taken_is_computed(tmp_path: Path) -> None:
    cells = cell("A1", "IF(TRUE,1,B1)", "5") + cell("B1", "A1+1", "7")
    source = book(tmp_path, FLAG, cells)
    wb = xlrecalc.load(source)
    assert [wb.value("A1"), wb.value("B1")] == [1.0, 2.0]
    assert_computed(wb.report)


def test_a_file_written_by_openpyxl_with_cached_values_added(tmp_path: Path) -> None:
    """openpyxl sets ``fullCalcOnLoad`` on every file: one run from the values added."""
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["B1"], ws["C1"], ws["D1"], ws["F1"] = "=B1+F1", "=A1/2", "=A1+100", "=F1+1", 1
    wb.calculation.iterate = True
    plain = tmp_path / "plain.xlsx"
    wb.save(plain)
    source = tmp_path / "cached.xlsx"
    with zipfile.ZipFile(plain) as zin, zipfile.ZipFile(source, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                for address, value in (("A1", "5"), ("B1", "7"), ("C1", "999"), ("D1", "77")):
                    text, n = re.subn(
                        rf'(<c r="{address}"[^>]*>\s*<f>[^<]*</f>)\s*<v\s*/>',
                        rf"\g<1><v>{value}</v>",
                        text,
                    )
                    assert n == 1
                data = text.encode()
            zout.writestr(item, data)
    assert 'fullCalcOnLoad="1"' in zipfile.ZipFile(source).read("xl/workbook.xml").decode()
    in_place = tmp_path / "in_place.xlsx"
    shutil.copy(source, in_place)
    assert_computed(xlrecalc.recalculate(in_place))
    assert written(in_place) == {**run(5.0, 7.0, 999.0, f=1.0), "D1": 2.0, "F1": 1.0}


def loop(tmp_path: Path, a1: str, extra: str = "") -> Path:
    cells = cell("A1", a1, "5") + cell("B1", "A1/2", "7") + cell("C1", "A1+100", "999") + extra
    return book(tmp_path, FLAG, cells)


def test_an_absent_function_in_a_loop_keeps_its_value_and_the_rest_rests_on_it(
    tmp_path: Path,
) -> None:
    wb = xlrecalc.load(loop(tmp_path, "B1+0*BESSELJ(1,1)"))
    wb.recalculate()
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == [5.0, 2.5, 105.0]
    assert wb.report.from_kept == {CellRef("Data", 1, c) for c in (1, 2, 3)}
    assert not wb.report.uncomputed


def test_a_refused_formula_in_a_loop_is_a_constant_and_breaks_it(tmp_path: Path) -> None:
    wb = xlrecalc.load(loop(tmp_path, "B1+ABS()"))
    wb.recalculate()
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == [5.0, 2.5, 105.0]
    assert set(wb.report.refused) == {CellRef("Data", 1, 1)}


def test_an_error_entering_a_loop_fills_it(tmp_path: Path) -> None:
    wb = xlrecalc.load(loop(tmp_path, "B1+F1", '<c r="F1" t="e"><v>#N/A</v></c>'))
    wb.recalculate()
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == [ErrorValue.NA] * 3
    assert_computed(wb.report)


def test_two_loops_sharing_a_cell_are_one_run(tmp_path: Path) -> None:
    """``A1`` is on both; every cell is computed in order each iteration, until none changes."""
    cells = (
        cell("A1", "B1/2+D1/4+1", "0")
        + cell("B1", "A1", "0")
        + cell("C1", "A1+100", "0")
        + cell("D1", "A1", "0")
    )
    wb = xlrecalc.load(book(tmp_path, FLAG, cells))
    wb.recalculate()
    a = b = c = d = 0.0
    for _ in range(100):
        na = b / 2 + d / 4 + 1
        nb, nc, nd = na, na + 100, na
        change = max(abs(na - a), abs(nb - b), abs(nc - c), abs(nd - d))
        a, b, c, d = na, nb, nc, nd
        if change < 0.001:
            break
    assert [wb.value(x) for x in ("A1", "B1", "C1", "D1")] == [a, b, c, d]
    assert_computed(wb.report)


def test_a_run_past_the_limit_keeps_the_values_and_names_the_limit(tmp_path: Path) -> None:
    """Three cells, 100 iterations at most: 300 evaluations would pass a limit of 200."""
    source = book(tmp_path, '<calcPr iterate="1" iterateDelta="0" fullCalcOnLoad="1"/>')
    limits = xlrecalc.ReadLimits(max_iteration_evaluations=200)
    wb = xlrecalc.load(source, limits=limits)
    wb.recalculate()
    assert {a: wb.value(a) for a in LOOP} == LOOP
    assert wb.report.uncomputed == {CellRef("Data", 1, c) for c in (1, 2, 3)}
    assert wb.report.unsupported[ITERATION_LIMIT] == 3
    out = tmp_path / "out.xlsx"
    xlrecalc.recalculate(source, out, limits=limits)
    assert written(out) == {**LOOP, "D1": 1.0, "F1": 0.0}
    wb = xlrecalc.load(source, limits=xlrecalc.ReadLimits(max_iteration_evaluations=300))
    wb.recalculate()
    assert_computed(wb.report)


def _iterate_open() -> dict[tuple[str, str], tuple[str, str]]:
    text = (Path(__file__).with_name("captures") / "iterate_open.tsv").read_text(encoding="utf-8")
    rows = (line.split("\t") for line in text.splitlines() if not line.startswith("#"))
    return {(r[0], r[3]): (r[7], r[9]) for r in rows}


@pytest.mark.excel
def test_excel_recalculates_on_opening_whatever_the_flag_says() -> None:
    """Capture ``iterate_open``: why the writer's ``fullCalcOnLoad`` changes nothing here."""
    seen = _iterate_open()
    counter = "1"
    assert seen["flag", counter] == seen["no_flag", counter] == ("110.0", "210.0")
    assert seen["flag_count_1", counter] == ("11.0", "12.0")
    assert seen["flag_count_3_delta_0", counter] == ("13.0", "16.0")
    assert seen["no_flag_all_cached_off", counter] == ("10.0", "10.0")
    assert seen["excel_copy", counter] == ("310.0", "410.0")
    for row in "12345":
        assert seen["excel_copy", row] == seen["excel_copy_written", row]
        assert seen["written_by_recalculate", row] == seen["no_flag", row]


@pytest.mark.excel
def test_excel_leaves_a_loop_without_cached_values_at_value_error() -> None:
    seen = _iterate_open()
    assert seen["flag", "6"] == ("missing value", "#VALUE!")
