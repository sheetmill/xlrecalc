"""The public API: ``load``, ``Workbook`` and saving changes (ARCHITECTURE.md §9.2, §11.2)."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
import xlformat

import xlrecalc
from tests.test_write import cells_of, raw_entries, sheet_xml
from tests.test_xlsx import MAIN, package
from xlrecalc import CellRef, ErrorValue, FormulaError, load
from xlrecalc._xlsx import Book

DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROWS = (
    '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v>6</v></c>'
    '<c r="C1"><f>B1+1</f><v>7</v></c></row>'
    '<row r="3" spans="1:2"><c r="A3" s="1"><v>5</v></c><c r="C3"><v>9</v></c></row>'
)


def book(tmp_path: Path, data: bytes | None = None, name: str = "book.xlsx") -> Path:
    path = tmp_path / name
    path.write_bytes(package(ROWS) if data is None else data)
    return path


def with_calc_chain(rows: str) -> bytes:
    """A package holding a calculation chain, as Excel saves one."""
    data = package(rows)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts["xl/calcChain.xml"] = f'<calcChain xmlns="{MAIN}"><c r="B1" i="1"/></calcChain>'.encode()
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rId7" Type="{DOC}/calcChain" Target="calcChain.xml"/></Relationships>',
    ).encode()
    parts["[Content_Types].xml"] = (
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-'
        b'officedocument.spreadsheetml.calcChain+xml"/></Types>'
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)
    return buf.getvalue()


def test_public_names():
    assert sorted(xlrecalc.__all__) == sorted(
        [
            "CellRef",
            "CellValue",
            "ErrorValue",
            "FormulaError",
            "LimitError",
            "PackageError",
            "Quirks",
            "ReadLimits",
            "Report",
            "Workbook",
            "__version__",
            "load",
            "recalculate",
        ]
    )


def test_values_and_formulas(tmp_path):
    wb = load(book(tmp_path))
    assert wb.sheet_names == ["Data"]
    assert wb.value("C1") == 7.0
    assert wb.value("Data!C1") == wb.value("'Data'!$C$1") == wb.value(CellRef("Data", 1, 3))
    assert wb.value("B2") is None
    assert wb.values("A1:C3") == [[2.0, 6.0, 7.0], [None, None, None], [5.0, None, 9.0]]
    assert wb.formula("B1") == "=A1*3"
    assert wb.formula("A1") is None
    with pytest.raises(KeyError):
        wb.value("Nope!A1")
    with pytest.raises(ValueError, match="reference"):
        wb.value("A")
    with pytest.raises(ValueError, match="single cell"):
        wb.value("A1:B2")


def test_a_reference_without_a_sheet_needs_a_workbook_of_one_sheet(tmp_path):
    two = package(
        ROWS,
        sheets='<sheet name="Data" sheetId="1" r:id="rId1"/>'
        '<sheet name="Other" sheetId="2" r:id="rId1"/>',
    )
    wb = load(book(tmp_path, two))
    assert wb.value("Data!C1") == 7.0
    with pytest.raises(ValueError, match=r"\['Data', 'Other'\]"):
        wb.value("C1")


def test_changes_propagate(tmp_path):
    wb = load(book(tmp_path))
    wb.set("A1", 10)
    assert wb.value("C1") == 31.0
    wb.set_formula("D1", "=C1*2")
    assert wb.value("D1") == 62.0
    wb.set_formula("B1", "=A1+A3")
    assert wb.value("C1") == 16.0
    assert wb.formula("B1") == "=A1+A3"


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (3, 3.0),
        (2.5, 2.5),
        (True, True),
        ("text", "text"),
        ("#N/A", "#N/A"),
        ("=A1", "=A1"),
        (ErrorValue.NA, ErrorValue.NA),
        (xlformat.ExcelError.DIV0, ErrorValue.DIV0),
        (None, None),
    ],
)
def test_set_values(tmp_path, value, expected):
    wb = load(book(tmp_path))
    wb.set("E1", value)
    got = wb.value("E1")
    assert got == expected
    assert type(got) is type(expected)


@pytest.mark.parametrize(
    ("value", "error"),
    [(float("nan"), ValueError), (float("inf"), ValueError), ("x" * 32768, ValueError)],
)
def test_set_refuses_what_excel_does_not_hold(tmp_path, value, error):
    with pytest.raises(error):
        load(book(tmp_path)).set("E1", value)
    with pytest.raises(TypeError):
        load(book(tmp_path)).set("E1", object())


@pytest.mark.parametrize(
    "formula", ["A1*2", "=SUM(", "=-\t1", "=SUM( )", "=#SPILL!", "=#CALC!", "=1+", "=A1 B1 ("]
)
def test_set_formula_refuses_what_the_formula_bar_refuses(tmp_path, formula):
    with pytest.raises(FormulaError):
        load(book(tmp_path)).set_formula("E1", formula)


def test_save_writes_every_kind_of_change(tmp_path):
    src = book(tmp_path)
    wb = load(src)
    wb.set("A1", 10)  # a constant replaced
    wb.set("A3", None)  # cleared, its style kept
    wb.set("C3", None)  # cleared, no style: the element goes
    wb.set("B3", "  padded ")  # a new cell between two others
    wb.set("A2", True)  # a new row between two others
    wb.set("D9", ErrorValue.NA)  # a new row at the end
    wb.set_formula("D1", '=IFNA(C1,0)&" x"')  # a new formula with a text result
    out = tmp_path / "out.xlsx"
    wb.save(out)
    got = cells_of(out)
    assert got["A1"] == ("n", 10.0, False)
    assert got["B1"] == ("n", 30.0, True)
    assert got["C1"] == ("n", 31.0, True)
    assert got["D1"] == ("str", "31 x", True)
    assert got["A2"] == ("b", True, False)
    assert got["A3"] == ("none", None, False)
    assert got["B3"] == ("inlineStr", "  padded ", False)
    assert "C3" not in got
    assert got["D9"] == ("e", "#N/A", False)
    xml = sheet_xml(out)
    assert '<c r="A3" s="1"/>' in xml
    assert '<f>_xlfn.IFNA(C1,0)&amp;" x"</f>' in xml
    assert '<row r="3">' in xml  # spans dropped from a row that gained a cell
    assert xml.index('r="A2"') < xml.index('r="A3"') < xml.index('r="D9"')
    assert raw_entries(src) == raw_entries(tmp_path / "book.xlsx")  # the source is untouched
    again = load(out)
    assert again.formula("D1") == '=IFNA(C1,0)&" x"'
    assert again.value("D1") == "31 x"


def test_save_in_place_and_from_a_file_object(tmp_path):
    src = book(tmp_path)
    wb = load(src)
    wb.set("A1", 4)
    wb.save()
    assert cells_of(src)["C1"] == ("n", 13.0, True)
    buffered = load(io.BytesIO(src.read_bytes()))
    buffered.set("A1", 5)
    with pytest.raises(ValueError, match="path"):
        buffered.save()
    buffered.save(tmp_path / "copy.xlsx")
    assert cells_of(tmp_path / "copy.xlsx")["C1"] == ("n", 16.0, True)


def test_calc_chain_is_kept_for_values_and_dropped_when_a_formula_goes(tmp_path):
    src = book(tmp_path, with_calc_chain(ROWS))
    wb = load(src)
    wb.set("A1", 4)
    wb.set_formula("D1", "=A1")
    wb.save(tmp_path / "values.xlsx")
    with zipfile.ZipFile(tmp_path / "values.xlsx") as z:
        assert "xl/calcChain.xml" in z.namelist()
    wb.set("B1", 1)  # a formula made a constant: Excel needs the chain gone
    wb.save(tmp_path / "removed.xlsx")
    with zipfile.ZipFile(tmp_path / "removed.xlsx") as z:
        assert "xl/calcChain.xml" not in z.namelist()
        assert b"calcChain" not in z.read("xl/_rels/workbook.xml.rels")
        assert b"calcChain" not in z.read("[Content_Types].xml")
        assert b"sharedStrings" in z.read("xl/_rels/workbook.xml.rels")
    assert cells_of(tmp_path / "removed.xlsx")["B1"] == ("n", 1.0, False)


def test_editing_a_shared_formula_master_gives_the_others_their_own_text(tmp_path):
    rows = (
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f>'
        '<v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/>'
        '<v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/>'
        "<v>6</v></c></row>"
    )
    src = book(tmp_path, package(rows))
    wb = load(src)
    assert wb.formula("B3") == "=A3*2"
    wb.set_formula("B1", "=A1*10")
    wb.save(tmp_path / "out.xlsx")
    xml = sheet_xml(tmp_path / "out.xlsx")
    assert '<c r="B1"><f>A1*10</f><v>10</v></c>' in xml
    assert '<c r="B2"><f>A2*2</f><v>4</v></c>' in xml
    assert '<c r="B3"><f>A3*2</f><v>6</v></c>' in xml
    assert 't="shared"' not in xml


def test_array_formula_cells_cannot_be_changed_one_by_one(tmp_path):
    rows = (
        '<row r="1"><c r="A1"><f t="array" ref="A1:A2">ROW(B1:B2)</f><v>1</v></c></row>'
        '<row r="2"><c r="A2"><v>2</v></c></row>'
    )
    wb = load(book(tmp_path, package(rows)))
    assert wb.value("A2") == 2.0
    for ref in ("A1", "A2"):
        with pytest.raises(ValueError, match="array"):
            wb.set(ref, 1)
        with pytest.raises(ValueError, match="array"):
            wb.set_formula(ref, "=1")


def test_new_cells_in_an_empty_sheet(tmp_path):
    data = package(sheet_xml=f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>')
    src = book(tmp_path, data)
    wb = load(src)
    wb.set("B2", 1)
    wb.set_formula("A1", "=B2+1")
    wb.save()
    assert sheet_xml(src) == (
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><f>B2+1</f><v>2</v></c></row>'
        '<row r="2"><c r="B2"><v>1</v></c></row></sheetData></worksheet>'
    )


def test_an_empty_row_element_gets_its_cells(tmp_path):
    data = package('<row r="1"><c r="A1"><v>1</v></c></row><row r="2" ht="20" customHeight="1"/>')
    src = book(tmp_path, data)
    wb = load(src)
    wb.set("C2", 3)
    wb.save()
    xml = sheet_xml(src)
    assert '<row r="2" ht="20" customHeight="1"><c r="C2"><v>3</v></c></row>' in xml


def test_recalculate_and_report(tmp_path):
    rows = '<row r="1"><c r="A1"><f>NOSUCH(1)</f></c><c r="B1"><f>A1+1</f></c></row>'
    wb = load(book(tmp_path, package(rows)))
    report = wb.recalculate()
    assert report is wb.report
    assert report.uncomputed == {CellRef("Data", 1, 1), CellRef("Data", 1, 2)}
    wb.save(tmp_path / "out.xlsx")
    got = cells_of(tmp_path / "out.xlsx")
    assert got["A1"] == ("none", None, True)  # uncomputed: left as it was


MARKER = [
    line.split("\t")
    for line in (Path(__file__).with_name("captures") / "marker_typed.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if not line.startswith("#")
]


def test_marker_typed_is_what_the_kept_workbooks_hold():
    """Capture marker_typed is derived from the typed workbooks kept beside it, row for row."""
    from capture.marker_probe import KEPT, typed_rows

    assert typed_rows(*KEPT) == ["\t".join(row) for row in MARKER]


@pytest.mark.excel
@pytest.mark.parametrize("row", MARKER, ids=lambda r: r[0])
def test_set_formula_writes_what_excel_writes_for_a_typed_formula(tmp_path, row):
    """Capture marker_typed: Excel's file form, dynamic-array marker and range for each."""
    import ast

    typed, form, cm, kind, ref, value = map(ast.literal_eval, row)
    data = "".join(f'<row r="{r}"><c r="Z{r}"><v>{r}</v></c></row>' for r in (1, 2, 3))
    src = book(tmp_path, package(data))
    wb = load(src)
    at = 4 * MARKER.index(row) + 1
    wb.set_formula(f"A{at}", typed)
    wb.save()
    with Book(src) as b:
        cell = next(c for c in b.cells("Data") if (c.row, c.column) == (at, 1))
    assert cell.formula is not None
    assert cell.formula.text == form
    assert (cell.cm is not None) is (cm is not None)
    assert cell.formula.kind == kind
    assert cell.formula.ref == ref
    assert cell.value == value
