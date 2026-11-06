"""With iterative calculation on, a cycle keeps its values and is reported (ARCHITECTURE.md §7.2).

Iterative calculation is not implemented. Whenever evaluation reaches a cycle in a workbook with
``<calcPr iterate="1">``, by any path, its cells and their dependents keep their cached values, are
uncomputed under the label ``iterative calculation``, and are not ``circular``. A run that trusts
the file's values evaluates nothing, changes nothing and reports nothing.

The cached values here differ from what one evaluation would give (``B1`` holds 7 where ``A1/2``
is 2.5, ``C1`` holds 999 where ``A1+100`` is 105), so a value computed and not restored shows.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest

import xlrecalc
from tests.test_xlsx import package
from xlrecalc._report import CellRef, column_letters
from xlrecalc._xlsx import Book

FLAG = '<calcPr iterate="1" fullCalcOnLoad="1"/>'
TRUST = '<calcPr iterate="1"/>'
OFF = '<calcPr fullCalcOnLoad="1"/>'
LOOP = {"A1": 5.0, "B1": 7.0, "C1": 999.0}
LABEL = "iterative calculation"
REPORTED = {CellRef("Data", 1, 1), CellRef("Data", 1, 2), CellRef("Data", 1, 3)}


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


def assert_kept(report: xlrecalc.Report) -> None:
    assert report.uncomputed == REPORTED
    assert report.unsupported[LABEL] == 3
    assert not report.circular


def test_a_file_asking_for_a_full_calculation(tmp_path: Path) -> None:
    source = book(tmp_path, FLAG)
    out = tmp_path / "out.xlsx"
    assert_kept(xlrecalc.recalculate(source, out))
    assert written(out) == {**LOOP, "D1": 1.0, "F1": 0.0}


def test_full_recalculation(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    out = tmp_path / "out.xlsx"
    assert_kept(xlrecalc.recalculate(source, out, full=True))
    assert written(out) == {**LOOP, "D1": 1.0, "F1": 0.0}


def test_a_trusting_run_changes_nothing_and_claims_nothing(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    out = tmp_path / "out.xlsx"
    report = xlrecalc.recalculate(source, out)
    assert not report.uncomputed
    assert not report.circular
    assert LABEL not in report.unsupported
    assert written(out) == {**LOOP, "D1": 77.0, "F1": 0.0}


def test_the_second_run_of_a_written_file(tmp_path: Path) -> None:
    """The first run writes ``fullCalcOnLoad``, so the second evaluates every formula."""
    path = book(tmp_path, TRUST)
    xlrecalc.recalculate(path)
    assert_kept(xlrecalc.recalculate(path))
    assert written(path) == {**LOOP, "D1": 1.0, "F1": 0.0}
    assert_kept(xlrecalc.recalculate(path))
    assert written(path) == {**LOOP, "D1": 1.0, "F1": 0.0}


@pytest.mark.parametrize(
    ("calc", "trust"),
    [(TRUST, False), (FLAG, True), (FLAG, False)],
    ids=["no_trust", "flag", "both"],
)
def test_value_gives_the_cached_values(tmp_path: Path, calc: str, trust: bool) -> None:
    """Evaluation replaced them before the cycle closed: they must come back."""
    wb = xlrecalc.load(book(tmp_path, calc), trust_cached=trust)
    wb.recalculate()
    assert [wb.value(a) for a in ("A1", "B1", "C1", "D1")] == [5.0, 7.0, 999.0, 1.0]
    assert_kept(wb.report)


@pytest.mark.parametrize("order", [("C1", "A1", "B1"), ("B1", "C1", "A1"), ("A1", "C1", "B1")])
def test_value_on_demand_in_any_order(tmp_path: Path, order: tuple[str, str, str]) -> None:
    wb = xlrecalc.load(book(tmp_path, FLAG))
    assert {a: wb.value(a) for a in order} == LOOP
    assert_kept(wb.report)


def test_a_trusting_load_gives_the_cached_values_without_evaluating(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, TRUST))
    assert {a: wb.value(a) for a in LOOP} == LOOP
    assert not wb.report.uncomputed


def test_recalculating_again_counts_the_cells_once(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, FLAG))
    wb.recalculate()
    wb.recalculate()
    wb.recalculate(full=True)
    assert_kept(wb.report)
    assert {a: wb.value(a) for a in LOOP} == LOOP


def test_an_edit_reaching_the_loop(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    wb = xlrecalc.load(source)
    wb.set("F1", 1.0)
    assert [wb.value(a) for a in ("A1", "B1", "C1", "D1")] == [5.0, 7.0, 999.0, 2.0]
    assert_kept(wb.report)
    out = tmp_path / "out.xlsx"
    wb.save(out)
    assert written(out) == {**LOOP, "D1": 2.0, "F1": 1.0}


def test_a_formula_set_into_the_loop_has_no_cached_value(tmp_path: Path) -> None:
    source = book(tmp_path, TRUST)
    wb = xlrecalc.load(source)
    wb.set_formula("A1", "=B1+F1+1")
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == ["", 7.0, 999.0]
    assert_kept(wb.report)
    out = tmp_path / "out.xlsx"
    wb.save(out)
    assert not has_cached(out, "A1")
    assert written(out)["B1"] == 7.0
    assert written(out)["C1"] == 999.0


def test_an_edit_breaking_the_loop_computes_its_cells(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, FLAG))
    wb.recalculate()
    assert_kept(wb.report)
    wb.set("A1", 10.0)
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == [10.0, 5.0, 110.0]
    assert not wb.report.uncomputed
    assert LABEL not in wb.report.unsupported
    out = tmp_path / "out.xlsx"
    wb.save(out)
    assert written(out) == {"A1": 10.0, "B1": 5.0, "C1": 110.0, "D1": 1.0, "F1": 0.0}


def test_a_cycle_without_cached_values_claims_nothing(tmp_path: Path) -> None:
    cells = cell("A1", "B1+0", None) + cell("B1", "A1/2", None) + cell("C1", "A1+100", None)
    source = book(tmp_path, FLAG, cells)
    wb = xlrecalc.load(source)
    assert [wb.value(a) for a in ("A1", "B1", "C1")] == ["", "", ""]
    assert_kept(wb.report)
    out = tmp_path / "out.xlsx"
    xlrecalc.recalculate(source, out)
    assert not any(has_cached(out, a) for a in ("A1", "B1", "C1"))


def test_a_volatile_cell_in_a_loop_is_reached_by_a_trusting_run(tmp_path: Path) -> None:
    cells = cell("A1", "B1+RAND()*0", "5") + cell("B1", "A1/2", "7") + cell("C1", "A1+100", "999")
    source = book(tmp_path, TRUST, cells)
    out = tmp_path / "out.xlsx"
    assert_kept(xlrecalc.recalculate(source, out))
    assert written(out) == LOOP


def test_an_array_formula_depending_on_a_loop_keeps_its_range(tmp_path: Path) -> None:
    cells = (
        cell("A1", "B1+0", "5")
        + cell("B1", "A1/2", "7")
        + '<c r="C1"><f t="array" ref="C1:E1">A1*{1,2,3}</f><v>11</v></c>'
        + '<c r="D1"><v>22</v></c><c r="E1"><v>33</v></c>'
        + cell("F1", "D1+1", "44")
    )
    source = book(tmp_path, FLAG, cells)
    wb = xlrecalc.load(source)
    wb.recalculate()
    assert [wb.value(a) for a in ("A1", "B1", "C1", "D1", "E1", "F1")] == [5, 7, 11, 22, 33, 44]
    assert wb.report.uncomputed == {
        CellRef("Data", 1, 1),
        CellRef("Data", 1, 2),
        CellRef("Data", 1, 3),
        CellRef("Data", 1, 6),
    }
    out = tmp_path / "out.xlsx"
    xlrecalc.recalculate(source, out)
    assert written(out) == {"A1": 5, "B1": 7, "C1": 11, "D1": 22, "E1": 33, "F1": 44}


def test_a_loop_through_a_branch_not_taken_is_computed(tmp_path: Path) -> None:
    cells = cell("A1", "IF(TRUE,1,B1)", "5") + cell("B1", "A1+1", "7")
    source = book(tmp_path, FLAG, cells)
    wb = xlrecalc.load(source)
    assert [wb.value("A1"), wb.value("B1")] == [1.0, 2.0]
    assert not wb.report.uncomputed
    assert LABEL not in wb.report.unsupported


def test_a_file_written_by_openpyxl_with_cached_values_added(tmp_path: Path) -> None:
    """openpyxl sets ``fullCalcOnLoad`` on every file, which is what made the default path wipe."""
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["B1"], ws["C1"], ws["D1"], ws["F1"] = "=B1+0", "=A1/2", "=A1+100", "=F1+1", 1
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
    assert_kept(xlrecalc.recalculate(in_place))
    assert written(in_place) == {**LOOP, "D1": 2.0, "F1": 1.0}


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
