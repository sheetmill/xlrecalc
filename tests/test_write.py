"""The surgical writer (ARCHITECTURE.md §9.2): what it changes, and that it changes nothing else."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from capture._file import Book

from tests.xlsx_files import package
from xlrecalc import CellRef, recalculate
from xlrecalc._report import column_letters, sheet_prefix


def raw_entries(path: Path) -> dict[str, bytes]:
    """Each entry's local header, compressed data and descriptor, as stored."""
    out = {}
    with zipfile.ZipFile(path) as z, open(path, "rb") as f:
        infos = z.infolist()
        offsets = [*sorted(i.header_offset for i in infos), z.start_dir]
        for info in infos:
            end = offsets[offsets.index(info.header_offset) + 1]
            f.seek(info.header_offset)
            out[info.filename] = f.read(end - info.header_offset)
    return out


def cells_of(path: Path, sheet: str = "Data") -> dict[str, tuple]:
    with Book(path) as book:
        return {
            f"{column_letters(c.column)}{c.row}": (c.type, c.value, c.formula is not None)
            for c in book.cells(sheet)
        }


def sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode()


def workbook_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/workbook.xml").decode()


def save(tmp_path: Path, data: bytes, name: str = "book.xlsx") -> Path:
    path = tmp_path / name
    path.write_bytes(data)
    return path


def fixed(**values: object):
    """A lookup giving these values (file terms) by A1 address."""

    def lookup(row: int, col: int):
        return values.get(f"{column_letters(col)}{row}")

    return lookup


ROWS = (
    '<row r="1"><c r="A1"><v>2</v></c><c r="B1" s="3"><f>A1*3</f><v/></c>'
    '<c r="C1" t="str"><f>"x"&amp;A1</f><v>old</v></c><c r="D1"><f>A1&gt;1</f></c>'
    '<c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>'
)


def test_recalculate_end_to_end(tmp_path):
    rows = (
        '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v/></c>'
        '<c r="C1"><f>NOSUCH(1)</f><v/></c><c r="D1"><f>C1+1</f><v/></c>'
        '<c r="E1"><f>#SPILL!</f><v>5</v></c><c r="F1"><f>E1+1</f><v/></c>'
        '<c r="G1"><f>G1+1</f><v/></c></row>'
    )
    src = save(tmp_path, package(rows, workbook_extra=""))
    report = recalculate(src)
    got = cells_of(src)
    assert got["B1"] == ("n", 6.0, True)
    assert got["C1"] == ("none", None, True)  # uncomputed: left as it was
    assert got["D1"] == ("none", None, True)
    assert got["E1"] == ("n", 5.0, True)  # refused: left as written, reported
    assert got["F1"] == ("n", 6.0, True)  # computed from the refused cell's cached value
    assert got["G1"] == ("str", "", True)  # circular, as Excel saves it
    assert report.uncomputed == {CellRef("Data", 1, 3), CellRef("Data", 1, 4)}
    assert set(report.refused) == {CellRef("Data", 1, 5)}
    assert report.circular == {CellRef("Data", 1, 7)}
    assert "fullCalcOnLoad" in workbook_xml(src)
    second = recalculate(src, tmp_path / "again.xlsx")
    assert second.circular == report.circular
    assert raw_entries(tmp_path / "again.xlsx") == raw_entries(src)


def test_recalculate_on_an_openpyxl_file(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = 2
    ws["B1"] = "=A1*3"
    ws["C1"] = '="x"&A1'
    ws["D1"] = "=A1>1"
    ws["E1"] = "=1/0"
    ws["F1"] = "=SUM(B1,A1)/2"
    src = tmp_path / "book.xlsx"
    wb.save(src)
    recalculate(src)
    values = [c.value for c in openpyxl.load_workbook(src, data_only=True)["Data"][1]]
    assert values == [2, 18 / 3, "x2", True, "#DIV/0!", 4]
    formulas = [c.value for c in openpyxl.load_workbook(src)["Data"][1]]
    assert formulas == [2, "=A1*3", '="x"&A1', "=A1>1", "=1/0", "=SUM(B1,A1)/2"]


@pytest.mark.parametrize(
    ("sheet", "text"),
    [
        ("Sheet1", "Sheet1!B7"),
        ("Sheet 1", "'Sheet 1'!B7"),
        ("O'Brien", "'O''Brien'!B7"),
        ("A1", "'A1'!B7"),
        ("R1C1", "'R1C1'!B7"),
        ("R", "'R'!B7"),
        ("1st", "'1st'!B7"),
        ("Résumé", "Résumé!B7"),
        ("my_sheet_2", "my_sheet_2!B7"),
        ("v1.2", "'v1.2'!B7"),
    ],
)
def test_cell_ref_text(sheet, text):
    assert str(CellRef(sheet, 7, 2)) == text
    assert sheet_prefix(sheet) + "B7" == text


def test_cell_ref_is_a_tuple():
    assert CellRef("S", 1, 2) == ("S", 1, 2)
    assert column_letters(1) == "A"
    assert column_letters(28) == "AB"
    assert column_letters(16384) == "XFD"
