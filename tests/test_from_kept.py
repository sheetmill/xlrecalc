"""A value resting on a formula the engine cannot evaluate is named (ARCHITECTURE.md §7.6).

A formula naming an absent function keeps the value the file holds, and what depends on it is
computed from that. ``Report.from_kept`` names them all. After a change upstream of the formula its
value is out of date: they move to ``uncomputed`` and keep their old values. A change that removes
the reliance takes them out. The two sets never overlap, and a trusting run that evaluates nothing
claims nothing.
"""

from __future__ import annotations

from pathlib import Path
from xml.sax.saxutils import escape

import pytest

import xlrecalc
from tests.test_xlsx import package
from xlrecalc._report import CellRef
from xlrecalc._values import ErrorValue
from xlrecalc._xlsx import Book

FLAG = '<calcPr fullCalcOnLoad="1"/>'
ABSENT = "NOSUCHFN(F1)"


def cell(address: str, formula: str | None, value: str | None) -> str:
    f = "" if formula is None else f"<f>{escape(formula)}</f>"
    v = "" if value is None else f"<v>{value}</v>"
    return f'<c r="{address}">{f}{v}</c>'


def chain(a1: str = ABSENT, cached: str | None = "10") -> str:
    """A1 keeps a value; B1 and C1 follow it; D1 reads B1 and a constant; E1 is unrelated."""
    return (
        cell("A1", a1, cached)
        + cell("B1", "A1*2", "999")
        + cell("C1", "B1+1", "999")
        + cell("D1", "B1+G1", "999")
        + cell("E1", "G1+1", "999")
        + cell("F1", None, "1")
        + cell("G1", None, "3")
    )


def book(tmp_path: Path, cells: str, calc: str = FLAG) -> Path:
    path = tmp_path / "book.xlsx"
    path.write_bytes(package(f'<row r="1">{cells}</row>', workbook_extra=calc))
    return path


def refs(*addresses: str) -> set[CellRef]:
    return {CellRef("Data", 1, ord(a[0]) - 64) for a in addresses}


def values(wb: xlrecalc.Workbook, *addresses: str) -> list[object]:
    out = [wb.value(a) for a in addresses]
    wb.recalculate()
    assert not wb.report.from_kept & wb.report.uncomputed
    return out


def in_file(path: Path) -> dict[str, object]:
    with Book(path) as b:
        return {f"{chr(64 + c.column)}{c.row}": c.value for c in b.cells("Data")}


def test_the_formula_and_what_is_computed_from_it_are_named(tmp_path: Path) -> None:
    source = book(tmp_path, chain())
    out = tmp_path / "out.xlsx"
    report = xlrecalc.recalculate(source, out)
    assert report.from_kept == refs("A1", "B1", "C1", "D1")
    assert not report.uncomputed
    assert report.unsupported["NOSUCHFN"] == 1
    assert in_file(out) == {
        "A1": 10.0,
        "B1": 20.0,
        "C1": 21.0,
        "D1": 23.0,
        "E1": 4.0,
        "F1": 1.0,
        "G1": 3.0,
    }


@pytest.mark.parametrize("full", [True, False])
def test_value_gives_the_kept_value(tmp_path: Path, full: bool) -> None:
    wb = xlrecalc.load(book(tmp_path, chain()), trust_cached=not full)
    assert values(wb, "A1", "B1", "C1", "D1", "E1") == [10.0, 20.0, 21.0, 23.0, 4.0]
    assert wb.report.from_kept == refs("A1", "B1", "C1", "D1")


def test_with_no_cached_value_it_behaves_as_before(tmp_path: Path) -> None:
    source = book(tmp_path, chain(cached=None))
    wb = xlrecalc.load(source)
    assert values(wb, "A1", "B1", "C1", "D1") == [ErrorValue.NAME] * 4
    assert wb.report.uncomputed == refs("A1", "B1", "C1", "D1")
    assert not wb.report.from_kept


def test_a_change_upstream_moves_them_to_uncomputed(tmp_path: Path) -> None:
    source = book(tmp_path, chain())
    wb = xlrecalc.load(source)
    wb.recalculate()
    wb.set("F1", 2.0)
    assert values(wb, "A1", "B1", "C1", "D1", "E1") == [10.0, 20.0, 21.0, 23.0, 4.0]
    assert wb.report.uncomputed == refs("A1", "B1", "C1", "D1")
    assert not wb.report.from_kept
    out = tmp_path / "out.xlsx"
    wb.save(out)
    assert in_file(out) == {
        "A1": 10.0,
        "B1": 999.0,
        "C1": 999.0,
        "D1": 999.0,
        "E1": 4.0,
        "F1": 2.0,
        "G1": 3.0,
    }


def test_a_change_upstream_on_a_trusting_load(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, chain(), calc=""))
    wb.set("F1", 2.0)
    assert values(wb, "A1", "B1", "C1") == [10.0, 999.0, 999.0]
    assert wb.report.uncomputed == refs("A1", "B1", "C1", "D1")


def test_a_change_to_another_precedent_recomputes_from_the_kept_value(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, chain()))
    wb.recalculate()
    wb.set("G1", 7.0)
    assert values(wb, "A1", "B1", "D1", "E1") == [10.0, 20.0, 27.0, 8.0]
    assert wb.report.from_kept == refs("A1", "B1", "C1", "D1")
    assert not wb.report.uncomputed


def test_replacing_the_formula_with_a_constant_removes_the_reliance(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, chain()))
    wb.recalculate()
    wb.set("A1", 100.0)
    assert values(wb, "A1", "B1", "C1", "D1") == [100.0, 200.0, 201.0, 203.0]
    assert not wb.report.from_kept
    assert not wb.report.uncomputed


def test_replacing_it_after_a_change_upstream_removes_the_reliance_too(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, chain()))
    wb.set("F1", 2.0)
    wb.recalculate()
    assert wb.report.uncomputed == refs("A1", "B1", "C1", "D1")
    wb.set_formula("A1", "=F1*5")
    assert values(wb, "A1", "B1", "C1") == [10.0, 20.0, 21.0]
    assert not wb.report.from_kept
    assert not wb.report.uncomputed
    out = tmp_path / "out.xlsx"
    wb.save(out)
    assert in_file(out)["C1"] == 21.0


def test_a_dependent_given_a_constant_or_another_formula_leaves(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, chain()))
    wb.recalculate()
    wb.set("C1", 5.0)
    wb.set_formula("D1", "=G1+1")
    assert values(wb, "C1", "D1") == [5.0, 4.0]
    assert wb.report.from_kept == refs("A1", "B1")


def test_an_absent_function_reading_through_indirect(tmp_path: Path) -> None:
    """It is never evaluated, so what it reads is not known: any change may reach it."""
    cells = chain(a1='NOSUCHFN(INDIRECT("F1"))')
    wb = xlrecalc.load(book(tmp_path, cells))
    wb.recalculate()
    assert wb.report.from_kept == refs("A1", "B1", "C1", "D1")
    wb.set("G1", 7.0)  # not a static precedent of A1
    assert values(wb, "A1", "B1", "E1") == [10.0, 20.0, 8.0]
    assert wb.report.uncomputed == refs("A1", "B1", "C1", "D1")
    assert not wb.report.from_kept


@pytest.mark.parametrize(
    "read", ['INDIRECT("A1")*2', "OFFSET(F1,0,-5)*2", 'SUM(INDIRECT("A1:A1"))*2']
)
def test_a_dependent_reading_it_through_a_dynamic_reference(tmp_path: Path, read: str) -> None:
    """No static edge shows this one: its read is recorded when it is evaluated."""
    cells = (
        cell("A1", ABSENT, "10")
        + cell("B1", read, "999")
        + cell("C1", "B1+1", "999")
        + cell("F1", None, "1")
        + cell("G1", None, "3")
    )
    source = book(tmp_path, cells)
    wb = xlrecalc.load(source)
    assert values(wb, "A1", "B1", "C1") == [10.0, 20.0, 21.0]
    assert wb.report.from_kept == refs("A1", "B1", "C1")
    wb.set("F1", 2.0)
    wb.recalculate()
    assert not wb.report.from_kept
    assert wb.report.uncomputed == refs("A1", "B1", "C1")


def test_a_long_chain_after_a_change_upstream(tmp_path: Path) -> None:
    rows = "".join(
        f'<row r="{r}">{cell(f"B{r}", "NOSUCHFN(F1)" if r == 1 else f"B{r - 1}+1", "7")}</row>'
        for r in range(1, 101)
    )
    path = tmp_path / "book.xlsx"
    path.write_bytes(
        package(rows + '<row r="101"><c r="F101"><v>1</v></c></row>', workbook_extra=FLAG)
    )
    wb = xlrecalc.load(path)
    assert wb.value("B100") == 106.0
    assert len(wb.report.from_kept) == 100
    wb.set("F1", 5.0)
    wb.recalculate()
    assert wb.value("B100") == 106.0
    assert len(wb.report.uncomputed) == 100
    assert not wb.report.from_kept


def test_a_trusting_run_that_evaluates_nothing_claims_nothing(tmp_path: Path) -> None:
    source = book(tmp_path, chain(), calc="")
    out = tmp_path / "out.xlsx"
    report = xlrecalc.recalculate(source, out)
    assert not report.from_kept
    assert not report.uncomputed
    assert in_file(out)["B1"] == 999.0


def test_a_trusting_load_names_what_an_edit_makes_it_evaluate(tmp_path: Path) -> None:
    wb = xlrecalc.load(book(tmp_path, chain(), calc=""))
    wb.set("G1", 7.0)
    assert values(wb, "D1", "E1") == [1006.0, 8.0]  # B1's cached 999 is trusted, and rests on A1
    assert wb.report.from_kept == refs("A1", "D1")
    wb2 = xlrecalc.load(book(tmp_path, chain(), calc=""))
    wb2.set_formula("B1", "=A1*2+G1")
    assert wb2.value("B1") == 23.0
    wb2.recalculate()
    assert CellRef("Data", 1, 1) in wb2.report.from_kept
    assert CellRef("Data", 1, 2) in wb2.report.from_kept


def test_recalculating_a_written_file_again_is_stable(tmp_path: Path) -> None:
    path = book(tmp_path, chain())
    first = xlrecalc.recalculate(path)
    held = in_file(path)
    second = xlrecalc.recalculate(path)
    assert in_file(path) == held
    assert second.from_kept == first.from_kept
