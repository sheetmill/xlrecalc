"""``ReadLimits.max_function_steps``: the loops inside functions, an allowance for a recalculation.

A crafted workbook of many cells each asking for the largest Bessel order Excel computes would
take about a second a cell; past the allowance, the cells keep the values the file holds and are
reported.
"""

from __future__ import annotations

import dataclasses
from pathlib import Path

import pytest
from capture._file import Book

import xlrecalc
from tests.xlsx_files import package
from xlrecalc._engine import STEP_LIMIT

CACHED = 99.0
STEPS = 10_020_000
"""What one ``BESSELJ(2.5,1E7)`` charges: the downward recurrence's 2 * ((n + isqrt(40 n)) // 2)."""


def crafted(tmp_path: Path, count: int) -> Path:
    """``count`` cells of ``BESSELJ(2.5,1E7)``, each caching 99, recalculated on load."""
    cells = "".join(
        f'<row r="{r}"><c r="A{r}"><f>BESSELJ(2.5,10000000)</f><v>{CACHED}</v></c></row>'
        for r in range(1, count + 1)
    )
    path = tmp_path / "steps.xlsx"
    path.write_bytes(package(cells, workbook_extra='<calcPr fullCalcOnLoad="1"/>'))
    return path


def test_cells_past_the_allowance_keep_their_values_and_are_reported(tmp_path: Path) -> None:
    source = crafted(tmp_path, 20)
    limits = xlrecalc.ReadLimits(max_function_steps=2 * STEPS + 1)
    wb = xlrecalc.load(source, limits=limits)
    wb.recalculate()
    values = [wb.value(f"A{r}") for r in range(1, 21)]
    assert values.count(0.0) == 2  # Excel's BESSELJ(2.5,1E7) underflows to 0
    assert values.count(CACHED) == 18
    assert len(wb.report.uncomputed) == 18
    assert wb.report.unsupported[STEP_LIMIT] == 18
    out = tmp_path / "out.xlsx"
    xlrecalc.recalculate(source, out, limits=limits)
    with Book(out) as b:
        written = [c.value for c in b.cells("Data")]
    assert written.count(CACHED) == 18


def test_each_recalculation_has_its_own_allowance(tmp_path: Path) -> None:
    source = crafted(tmp_path, 3)
    limits = xlrecalc.ReadLimits(max_function_steps=STEPS)
    wb = xlrecalc.load(source, limits=limits)
    for runs in (1, 2):  # without a fresh allowance the second run would refuse all three
        report = wb.recalculate(full=True)
        assert [wb.value(f"A{r}") for r in (1, 2, 3)] == [0.0, CACHED, CACHED]
        assert len(report.uncomputed) == 2
        assert report.unsupported[STEP_LIMIT] == 2 * runs  # the report's counts add up


def test_a_cell_within_the_allowance_is_computed(tmp_path: Path) -> None:
    wb = xlrecalc.load(crafted(tmp_path, 1), limits=xlrecalc.ReadLimits(max_function_steps=STEPS))
    wb.recalculate()
    assert wb.value("A1") == 0.0
    assert not wb.report.uncomputed


@pytest.mark.slow
def test_the_default_allowance_computes_a_workbook_of_such_cells(tmp_path: Path) -> None:
    """Three cells charge 30 million steps, more than the untrusted preset allows; about 2 s."""
    assert xlrecalc.ReadLimits.UNTRUSTED.max_function_steps < 3 * STEPS
    wb = xlrecalc.load(crafted(tmp_path, 3))
    wb.recalculate()
    assert {wb.value(f"A{r}") for r in range(1, 4)} == {0.0}
    assert not wb.report.uncomputed


def test_the_untrusted_preset_is_stricter_than_the_default_on_every_field() -> None:
    default, untrusted = xlrecalc.ReadLimits(), xlrecalc.ReadLimits.UNTRUSTED
    for field in dataclasses.fields(xlrecalc.ReadLimits):
        assert getattr(untrusted, field.name) < getattr(default, field.name), field.name
    assert untrusted.max_function_steps < STEPS * 2  # one cell at the largest order, not two


def test_the_untrusted_preset_stops_a_crafted_file_within_its_allowance(tmp_path: Path) -> None:
    wb = xlrecalc.load(crafted(tmp_path, 5), limits=xlrecalc.ReadLimits.UNTRUSTED)
    wb.recalculate()
    assert [wb.value(f"A{r}") for r in range(1, 6)].count(CACHED) == 4
    assert wb.report.unsupported[STEP_LIMIT] == 4
