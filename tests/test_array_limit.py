"""``ReadLimits.max_array_cells``: every route by which a formula makes an array is bounded.

Over the limit, a formula keeps the value the file holds and is reported; within it, the array
is computed. ``SUMIF``, ``SUMIFS`` and ``COUNTIFS`` make no array: they visit non-blank cells, and
are computed at any size. Captures ``array_size`` and ``array_functions`` have what Excel does.
Each test runs under the memory watchdog: if a check regresses, the test fails instead of the
machine running out.
"""

from __future__ import annotations

import io
import re
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

import pytest

import xlrecalc
from tests.test_xlsx import DOC, MAIN, PKG, package
from tests.watchdog import GIB, bounded
from xlrecalc import ReadLimits
from xlrecalc._xlsx import Book

CAPTURES = Path(__file__).with_name("captures")
LIMIT_LABEL = "an array of more than"


def two_sheets(formula: str, cached: str | None = "7", array: str | None = None) -> bytes:
    """Sheet S holds the formula in A1 (an array formula over ``array`` if given); sheet D holds
    two cells, A1 = 1 and B2 = 2."""
    f = f'<f t="array" ref="{array}">{formula}</f>' if array else f"<f>{formula}</f>"
    v = f"<v>{cached}</v>" if cached is not None else ""
    sheet_s = f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1">{f}{v}</c></row>'
    sheet_s += "</sheetData></worksheet>"
    sheet_d = (
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
        '<row r="2"><c r="B2"><v>2</v></c></row></sheetData></worksheet>'
    )
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId3" Type="{DOC}/worksheet" Target="worksheets/sheet2.xml"/>'
        "</Relationships>"
    )
    sheets = '<sheet name="S" sheetId="1" r:id="rId1"/><sheet name="D" sheetId="2" r:id="rId3"/>'
    return package(
        sheet_xml=sheet_s,
        sheets=sheets,
        extra_parts={"xl/_rels/workbook.xml.rels": rels, "xl/worksheets/sheet2.xml": sheet_d},
    )


def recalculated(data: bytes, limits: ReadLimits | None = None) -> xlrecalc.Workbook:
    book = xlrecalc.load(io.BytesIO(data), limits=limits)
    book.recalculate(full=True)
    return book


def assert_kept(book: xlrecalc.Workbook, cached: object = 7.0) -> None:
    assert book.value("S!A1") == cached
    report = book.report
    assert any(str(c) == "S!A1" for c in report.uncomputed)
    assert any(label.startswith(LIMIT_LABEL) for label in report.unsupported)


def test_the_default_is_ten_full_columns():
    assert ReadLimits().max_array_cells == 10 * 1_048_576


@pytest.mark.parametrize("array", ["A1:XFD1048576", "A1:K1048576"])
@bounded
def test_a_declared_range_past_the_limit_keeps_its_value(array):
    assert_kept(recalculated(two_sheets("1", array=array)))


@pytest.mark.parametrize(
    "formula",
    [
        "SUMPRODUCT(D!A:XFD*1)",  # a reference made into an array
        "SUMPRODUCT(D!A:XFD)",  # SUMPRODUCT's own grids
        "SUMPRODUCT((D!A1:A1048576+0)*(D!A1:XFD1+0))",  # a column times a row: every pair
        "SUMPRODUCT(ROW(D!A1:A1048576)*COLUMN(D!A1:XFD1))",  # the same from positions
    ],
)
@bounded
def test_an_array_made_past_the_limit_keeps_the_value(formula):
    assert_kept(recalculated(two_sheets(formula)))


@bounded
def test_a_dynamic_array_past_the_limit_is_reported():
    book = xlrecalc.load(io.BytesIO(two_sheets("1")))
    book.set_formula("S!C1", "=D!A:XFD*1")
    book.recalculate()
    assert any(str(c) == "S!C1" for c in book.report.uncomputed)


def test_an_array_within_the_limit_is_computed():
    book = recalculated(two_sheets("SUMPRODUCT(D!A1:J1000*1)", cached="0"))
    assert book.value("S!A1") == 3.0
    assert not book.report.uncomputed


@pytest.mark.slow
@bounded(limit=2 * GIB)
def test_ten_full_columns_are_computed():
    book = recalculated(two_sheets("SUMPRODUCT(D!A:J*1)", cached="0"))
    assert book.value("S!A1") == 3.0
    assert not book.report.uncomputed


def test_the_limit_can_be_lowered():
    limits = ReadLimits(max_array_cells=4)
    assert_kept(recalculated(two_sheets("1", array="A1:A5"), limits))
    book = recalculated(two_sheets("D!A1:A4*1", array="A1:A4"), limits)
    assert [book.value(f"S!A{r}") for r in range(1, 5)] == [1.0, 0.0, 0.0, 0.0]


def excel_results() -> dict[str, float]:
    """The ``SUMIF`` family in capture ``array_functions``: each formula's value after Excel's
    calculation, for the formulas Excel calculated (not one whose calculation timed out)."""
    results = {}
    lines = (CAPTURES / "array_functions.tsv").read_text(encoding="utf-8").splitlines()
    for line in lines:
        reads = [f for f in line.split("\t") if f.startswith("read: ") and " | " in f]
        calculated = "\tcalc: calculated " in line
        if calculated and reads and line.startswith(("SUMIF", "COUNTIF")):
            results[line.split("\t")[0]] = float(reads[-1][6:].split(" | ")[0])
    return results


def test_architecture_states_the_sumif_family_capture_counts():
    """ARCHITECTURE.md's counts are the capture's: formulas with Excel's value, and those over a
    whole sheet (every column, or all but one beside another range of all but one)."""
    results = excel_results()
    whole = [f for f in results if re.search(r"A:XF[CD]|B:XFD", f)]
    text = (Path(__file__).parents[1] / "ARCHITECTURE.md").read_text(encoding="utf-8")
    found = re.search(
        r"holds Excel's value for (\d+) formulas of `SUMIF`.*?(\d+) of them over a whole sheet",
        text,
        re.S,
    )
    assert found, "ARCHITECTURE.md: the array_functions counts are missing"
    assert (int(found[1]), int(found[2])) == (len(results), len(whole))


@bounded
def test_the_sumif_family_gives_excels_value_at_any_size():
    got = {}
    for formula in excel_results():
        book = recalculated(two_sheets(xml_escape(formula)))
        assert not book.report.uncomputed, formula
        got[formula] = book.value("S!A1")
    assert got == excel_results()


@bounded
def test_the_sumif_family_makes_no_array():
    for formula in (
        'SUMIFS(D!A:XFD,D!A:XFD,"")',
        'COUNTIFS(D!A:XFD,"")',
        'SUMIF(D!A:XFD,"")',
        'AVERAGEIF(D!A:XFD,">0")',
        'AVERAGEIFS(D!B:XFD,D!A:XFC,"")',
        '_xlfn.MAXIFS(D!B:XFD,D!A:XFC,"")',
        '_xlfn.MINIFS(D!A:XFD,D!A:XFD,">0")',
    ):
        book = recalculated(two_sheets(formula), ReadLimits(max_array_cells=4))
        assert not book.report.uncomputed, formula


@bounded
def test_the_writer_keeps_the_cached_value():
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "book.xlsx"
        path.write_bytes(two_sheets("SUMPRODUCT(D!A:XFD*1)", cached="3"))
        report = xlrecalc.recalculate(path, full=True)
        assert any(str(c) == "S!A1" for c in report.uncomputed)
        with Book(path) as book:  # the file's own cached value, untouched
            (cell,) = list(book.cells("S"))
        assert (cell.value, cell.formula.text) == (3.0, "SUMPRODUCT(D!A:XFD*1)")


@pytest.mark.parametrize(
    "formula",
    ["DSUM(D!A:XFD,1,D!A:XFD)", "DCOUNT(D!A:XFD,,D!A:XFD)", "DCOUNTA(D!A:XFD,1,D!A:XFD)"],
)
@bounded
def test_the_database_functions_read_only_the_cells_a_sheet_holds(formula):
    """Over a whole sheet of two cells they are computed, not refused, in little memory."""
    book = recalculated(two_sheets(formula))
    assert book.value("S!A1") != 7.0
    assert not book.report.unsupported
