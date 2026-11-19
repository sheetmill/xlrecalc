"""Reading cells one by one costs no more than one recalculation.

``Workbook.value`` on a cell that needs computing evaluates what that cell needs. The report's
whole-workbook passes (what is uncomputed, what rests on a kept value) run once, when the report
is read, not after every cell: before, a workbook whose formulas rest on a kept value (a function
the package does not compute) was quadratic to read cell by cell, over an hour for one corpus file.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

import xlrecalc
from tests.watchdog import bounded
from tests.xlsx_files import package
from xlrecalc._engine import Engine
from xlrecalc._report import CellRef


def crafted(tmp_path: Path, count: int) -> Path:
    """``A1`` an absent function (kept, 1); ``B1:B<count>`` one formula on it, filled down."""
    rows = ['<row r="1"><c r="A1"><f>RTD("a","","b")</f><v>1</v></c>'
            '<c r="B1"><f>$A$1+ROW()</f><v>0</v></c></row>']  # fmt: skip
    rows += [
        f'<row r="{r}"><c r="B{r}"><f>$A$1+ROW()</f><v>0</v></c></row>' for r in range(2, count + 1)
    ]
    path = tmp_path / "reads.xlsx"
    path.write_bytes(package("".join(rows), workbook_extra='<calcPr fullCalcOnLoad="1"/>'))
    return path


def test_the_report_passes_run_once_for_many_reads(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    calls = 0
    finish = Engine._finish_report

    def counted(self: Engine) -> None:
        nonlocal calls
        calls += 1
        finish(self)

    monkeypatch.setattr(Engine, "_finish_report", counted)
    wb = xlrecalc.load(crafted(tmp_path, 200))
    before = calls
    assert [wb.value(f"B{r}") for r in range(1, 201)] == [1.0 + r for r in range(1, 201)]
    assert calls == before  # the reads ran none of them
    report = wb.report
    assert calls == before + 1
    assert wb.report is report  # a second read of the report runs none
    assert calls == before + 1
    assert len(report.from_kept) == 201


def test_reads_one_by_one_report_what_one_recalculation_reports(tmp_path: Path) -> None:
    source = crafted(tmp_path, 50)
    read = xlrecalc.load(source)
    for r in range(1, 51):
        read.value(f"B{r}")
    whole = xlrecalc.load(source)
    whole.recalculate()
    assert read.report.from_kept == whole.report.from_kept
    assert read.report.uncomputed == whole.report.uncomputed
    assert read.report.unsupported == whole.report.unsupported
    assert CellRef("Data", 1, 1) in read.report.from_kept


@bounded(seconds=30.0)
def test_reading_a_large_workbook_cell_by_cell_is_linear(tmp_path: Path) -> None:
    """20,000 cells read one by one, then the report: about 0.1 s. Each read used to finish the
    report again over every cell resting on the kept value (8,000 cells took 21 s, 20,000 would
    take over two minutes), so this fails by its time limit under the old behavior."""
    wb = xlrecalc.load(crafted(tmp_path, 20_000))
    start = time.perf_counter()
    for r in range(1, 20_001):
        wb.value(f"B{r}")
    assert len(wb.report.from_kept) == 20_001
    assert time.perf_counter() - start < 5.0
