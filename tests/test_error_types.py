"""Excel's newer errors read from a file, and ``ERROR.TYPE`` over them.

Excel 16.115 (xlio's capture ``errors``, ``tests/captures/errors.tsv`` in its repository: each
error written into a cell, alone, opened and saved by Excel) caches a newer error as ``#VALUE!``
with a rich value holding its type, and gives ``ERROR.TYPE`` 9 ``#SPILL!``, 10 ``#CONNECT!``, 11
``#BLOCKED!``, 12 ``#UNKNOWN!``, 13 ``#FIELD!``, 14 ``#CALC!``, 16 ``#BUSY!``, 20 ``#TIMEOUT!``;
it writes ``#GETTING_DATA`` back as ``#N/A``.
"""

from __future__ import annotations

import io
import zipfile

import pytest

import xlrecalc
from tests.test_arrays import DOC, _metadata
from tests.xlsx_files import package
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


def _book(error_type: str, b1: str = '<c r="B1"><f>ERROR.TYPE(A1)</f></c>') -> bytes:
    """A1 a rich-value error of this type, as Excel saves one; B1 =ERROR.TYPE(A1) by default."""
    row = f'<row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c>{b1}</row>'

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


@pytest.mark.parametrize(("error_type", "error", "_number"), CAPTURED, ids=lambda x: str(x))
def test_a_new_value_drops_the_old_rich_value(  # type: ignore[no-untyped-def]
    tmp_path, error_type: str, error: str, _number: float
) -> None:
    """A cell caching a newer error (``#VALUE!`` and ``vm``) keeps ``vm`` only while it holds it.

    Excel shows a cell's rich value over its ``<v>`` (the probe ``capture/vm_probe.py``: a stale
    ``vm`` made Excel show and save the old error for 5, TRUE, ``#DIV/0!`` and a plain
    ``#VALUE!``). Excel refuses a file storing these codes as error values, so a new one is
    written as ``#VALUE!``.
    """
    path = tmp_path / "book.xlsx"
    b1 = '<c r="B1" t="e" vm="1"><f>IF(C1=1,"a"+1,IF(C1=2,5,A1))</f><v>#VALUE!</v></c>'
    path.write_bytes(_book(error_type, b1))
    book = xlrecalc.load(path, recalc_on_load="always")
    assert book.value("B1") == ErrorValue(error)
    book.set("D1", 1.0)  # an edit elsewhere: B1 holds the error its vm stands for
    book.save()
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<c r="B1" t="e" vm="1">' in sheet
    assert '<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>' in sheet  # a constant is kept as it is
    for c1, cell in ((1.0, ' t="e">'), (2.0, ">")):  # a plain #VALUE!, then 5
        book.set("C1", c1)
        book.save()
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode()
        assert f'<c r="B1"{cell}' in sheet
        assert xlrecalc.load(path, recalc_on_load="never").value("B1") == book.value("B1")


def test_a_blocked_spill_read_from_a_file_spills_once_cleared() -> None:
    """A dynamic array a file caches as blocked spills when a trusting load clears the blocker."""
    b1 = (
        '<c r="B1" t="e" cm="1" vm="1"><f t="array" ref="B1">D1:E1</f><v>#VALUE!</v></c>'
        '<c r="C1" t="inlineStr"><is><t>x</t></is></c><c r="D1"><v>1</v></c><c r="E1"><v>2</v></c>'
    )
    book = xlrecalc.load(io.BytesIO(_book("8", b1)), recalc_on_load="never")
    assert book.value("B1") == ErrorValue("#SPILL!")
    book.set("C1", None)
    assert (book.value("B1"), book.value("C1")) == (1.0, 2.0)


def test_computed_spill_and_calc_errors_are_written_as_the_value_error(  # type: ignore[no-untyped-def]
    tmp_path,
) -> None:
    """``#SPILL!`` and ``#CALC!``, which formulas compute, reach no cell as their own codes."""
    path = tmp_path / "book.xlsx"
    path.write_bytes(
        package('<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c>'
                '<c r="C2"><v>9</v></c></row>')
    )  # fmt: skip
    book = xlrecalc.load(path)
    book.set_formula("C1", "=A1:A2")  # blocked by C2
    book.set_formula("D1", "=FILTER(A1:A2,A1:A2>5)")  # nothing left
    assert (book.value("C1"), book.value("D1")) == (ErrorValue("#SPILL!"), ErrorValue("#CALC!"))
    book.save()
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    assert sheet.count("<v>#VALUE!</v>") == 2
    for _type, error, _number in CAPTURED:
        assert f"<v>{error}</v>" not in sheet


def test_getting_data_is_written_as_excel_writes_it_back(tmp_path) -> None:  # type: ignore[no-untyped-def]
    """A formula giving ``#GETTING_DATA`` is saved as ``#N/A``, as Excel saves it."""
    path = tmp_path / "book.xlsx"
    path.write_bytes(
        package('<row r="1"><c r="A1" t="e"><v>#GETTING_DATA</v></c><c r="B1"><f>A1</f></c></row>')
    )
    xlrecalc.recalculate(path, recalc_on_load="always")
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<c r="B1" t="e"><f>A1</f><v>#N/A</v></c>' in sheet
