"""Excel's newer errors read from a file, and ``ERROR.TYPE`` over them.

Excel 16.115 (a capture of each error written into a cell, alone, opened and saved by Excel)
caches a newer error as ``#VALUE!`` with a rich value holding its type, and gives
``ERROR.TYPE`` 9 ``#SPILL!``, 10 ``#CONNECT!``, 11 ``#BLOCKED!``, 12 ``#UNKNOWN!``, 13 ``#FIELD!``,
14 ``#CALC!``, 16 ``#BUSY!``, 20 ``#TIMEOUT!``; it writes ``#GETTING_DATA`` back as ``#N/A``.
"""

from __future__ import annotations

import io
import zipfile

import pytest

import xlrecalc
from tests.test_arrays import DOC, _metadata
from tests.test_xlsx import package
from xlrecalc._load import load
from xlrecalc._values import ErrorValue

RICH = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RELS = "http://schemas.microsoft.com/office/2017/06/relationships"

CAPTURED = [  # rich value errorType, the error Excel shows, its ERROR.TYPE
    ("8", "#SPILL!", 9.0),
    ("9", "#CONNECT!", 10.0),
    ("10", "#BLOCKED!", 11.0),
    ("11", "#UNKNOWN!", 12.0),
    ("12", "#FIELD!", 13.0),
    ("13", "#CALC!", 14.0),
    ("15", "#BUSY!", 16.0),
    ("19", "#TIMEOUT!", 20.0),
]


def _book(error_type: str) -> bytes:
    """A1 a rich-value error of this type, as Excel saves one; B1 =ERROR.TYPE(A1)."""
    row = (
        '<row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c>'
        '<c r="B1"><f>ERROR.TYPE(A1)</f></c></row>'
    )
    values = f'<rvData xmlns="{RICH}" count="1"><rv s="0"><v>{error_type}</v><v>0</v></rv></rvData>'
    structures = (
        f'<rvStructures xmlns="{RICH}" count="1"><s t="_error"><k n="errorType" t="i"/>'
        '<k n="subType" t="i"/></s></rvStructures>'
    )
    data = package(
        row,
        extra_parts={
            "xl/metadata.xml": _metadata(),
            "xl/richData/rdrichvalue.xml": values,
            "xl/richData/rdrichvaluestructure.xml": structures,
        },
    )
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        f'<Relationship Id="rId9" Type="{RELS}/rdRichValue" Target="richData/rdrichvalue.xml"/>'
        f'<Relationship Id="rId10" Type="{RELS}/rdRichValueStructure"'
        ' Target="richData/rdrichvaluestructure.xml"/></Relationships>',
    ).encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
    return buf.getvalue()


@pytest.mark.parametrize(("error_type", "error", "number"), CAPTURED)
def test_a_newer_error_is_read_with_its_number(error_type: str, error: str, number: float) -> None:
    e = load(io.BytesIO(_book(error_type)))
    e.recalculate(full=True)
    assert (e.get(0, 0, 0), e.get(0, 0, 1)) == (ErrorValue(error), number)


def test_an_uncaptured_rich_type_reads_as_the_value_error_written() -> None:
    e = load(io.BytesIO(_book("16")))
    e.recalculate(full=True)
    assert e.get(0, 0, 0) == ErrorValue.VALUE


def test_getting_data_is_written_as_excel_writes_it_back(tmp_path) -> None:  # type: ignore[no-untyped-def]
    """A formula giving ``#GETTING_DATA`` is saved as ``#N/A``, as Excel saves it."""
    path = tmp_path / "book.xlsx"
    path.write_bytes(
        package('<row r="1"><c r="A1" t="e"><v>#GETTING_DATA</v></c><c r="B1"><f>A1</f></c></row>')
    )
    xlrecalc.recalculate(path, full=True)
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<c r="B1" t="e"><f>A1</f><v>#N/A</v></c>' in sheet
