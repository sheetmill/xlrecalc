"""The surgical writer (ARCHITECTURE.md §9.2): what it changes, and that it changes nothing else."""

from __future__ import annotations

import io
import math
import os
import random
import struct
import zipfile
from pathlib import Path

import pytest

from tests.test_xlsx import MAIN, package
from xlrecalc import CellRef, recalculate
from xlrecalc._report import column_letters, sheet_prefix
from xlrecalc._xlsx import Book
from xlrecalc._xlsx._write import number_text, patch_cells, text_escape, write
from xlrecalc._xlsx._zip import ZipWriter


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


def test_values_of_every_kind_are_patched(tmp_path):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    count = write(
        src,
        out,
        {"Data": fixed(B1=("n", 6.0), C1=("str", "x2"), D1=("b", True), E1=("e", "#DIV/0!"))},
    )
    assert count == 3  # E1 already holds #DIV/0!
    assert cells_of(out) == {
        "A1": ("n", 2.0, False),
        "B1": ("n", 6.0, True),
        "C1": ("str", "x2", True),
        "D1": ("b", True, True),
        "E1": ("e", "#DIV/0!", True),
    }
    xml = sheet_xml(out)
    assert '<c r="B1" s="3"><f>A1*3</f><v>6</v></c>' in xml  # style kept, <v/> replaced
    assert '<c r="D1" t="b"><f>A1&gt;1</f><v>1</v></c>' in xml  # <v> added after <f>
    assert '<c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c>' in xml


def test_a_changed_type_replaces_the_t_attribute(tmp_path):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    write(src, out, {"Data": fixed(C1=("n", 1.5), E1=("str", ""))})
    xml = sheet_xml(out)
    assert '<c r="C1"><f>"x"&amp;A1</f><v>1.5</v></c>' in xml
    assert '<c r="E1" t="str"><f>1/0</f><v></v></c>' in xml
    assert cells_of(out)["E1"] == ("str", "", True)


def test_only_formula_cells_are_asked_and_none_keeps_a_cell(tmp_path):
    asked = []

    def lookup(row, col):
        asked.append((row, col))
        return None

    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    assert write(src, out, {"Data": lookup}) == 0
    assert asked == [(1, 2), (1, 3), (1, 4), (1, 5)]


def test_unchanged_parts_are_copied_byte_for_byte(tmp_path):
    src = save(tmp_path, package(ROWS, strings=["<si><t>a</t></si>"]))
    out = tmp_path / "out.xlsx"
    write(src, out, {"Data": fixed(B1=("n", 6.0))})
    before, after = raw_entries(src), raw_entries(out)
    assert list(before) == list(after)  # same entries, same order
    changed = {name for name in before if before[name] != after[name]}
    assert changed == {"xl/worksheets/sheet1.xml", "xl/workbook.xml"}


def test_a_sheet_without_changes_is_copied_raw(tmp_path):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    assert write(src, out, {"Data": fixed(E1=("e", "#DIV/0!"))}, full_calc_on_load=False) == 0
    assert raw_entries(src) == raw_entries(out)


def test_everything_else_in_a_sheet_is_kept_as_written(tmp_path):
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<x:worksheet xmlns:x="{MAIN}" xmlns:mc="urn:mc"><x:sheetPr/><x:sheetData>'
        '<x:row r="1" spans="1:3"><x:c r="A1"><x:v>1</x:v></x:c>'
        '<x:c t="str"><x:f t="shared" ref="B1:B2" si="0">A1 + 1</x:f><x:v>9</x:v></x:c></x:row>'
        "<!-- a comment --><x:row r='2'><x:c r='B2'><x:f t='shared' si='0'/><x:v>9</x:v></x:c>"
        "</x:row></x:sheetData><x:extLst><x:ext uri='{X}'><mc:any c='1'/></x:ext></x:extLst>"
        "</x:worksheet>"
    )
    src = save(tmp_path, package(sheet_xml=sheet))
    out = tmp_path / "out.xlsx"
    write(src, out, {"Data": fixed(B1=("n", 2.0), B2=("n", 3.0))})
    expected = sheet.replace(
        '<x:c t="str"><x:f t="shared" ref="B1:B2" si="0">A1 + 1</x:f><x:v>9</x:v>',
        '<x:c><x:f t="shared" ref="B1:B2" si="0">A1 + 1</x:f><x:v>2</x:v>',
    ).replace("<x:f t='shared' si='0'/><x:v>9</x:v>", "<x:f t='shared' si='0'/><x:v>3</x:v>")
    assert sheet_xml(out) == expected


@pytest.mark.parametrize(
    "text",
    [
        " leading and trailing ",
        "line\nbreak",
        "carriage\rreturn",
        "tab\there",
        "a_x0041_b",
        "_x005F_",
        "control\x01\x1f",
        "&<>\"'",
        "Résumé 日本 🙂",
        "",
    ],
)
def test_text_reads_back_as_written(tmp_path, text):
    src = save(tmp_path, package(ROWS))
    out = tmp_path / "out.xlsx"
    write(src, out, {"Data": fixed(C1=("str", text))})
    assert cells_of(out)["C1"] == ("str", text, True)


def test_text_escape():
    assert text_escape("a\rb") == "a_x000D_b"
    assert text_escape("_x0041_") == "_x005F_x0041_"
    assert text_escape("_x41_") == "_x41_"
    assert text_escape("1<2&3>0") == "1&lt;2&amp;3&gt;0"


def test_number_text_is_exact():
    rng = random.Random(1)
    samples = [0.0, -0.0, 1.0, -7.0, 2.0**53, 2.0**53 + 2, 1e16, 1e300, 5e-324, 0.1, 1 / 3]
    samples += [struct.unpack("<d", rng.randbytes(8))[0] for _ in range(2000)]
    for x in samples:
        if math.isfinite(x):
            assert float(number_text(x)) == x
    assert number_text(6.0) == "6"
    assert number_text(-0.0) == "0"
    assert number_text(1e16) == "1E+16"
    assert number_text(1 / 3) == "0.3333333333333333"


def test_numbers_compare_by_value_not_text(tmp_path):
    rows = '<row r="1"><c r="A1"><f>1/3</f><v>0.33333333333333331</v></c></row>'
    src = save(tmp_path, package(rows))
    assert write(src, tmp_path / "out.xlsx", {"Data": fixed(A1=("n", 1 / 3))}) == 0


@pytest.mark.parametrize(
    ("workbook", "expected"),
    [
        ('<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>'),
        ('<calcPr fullCalcOnLoad="0" calcId="1"/>', '<calcPr fullCalcOnLoad="1" calcId="1"/>'),
        ("<calcPr fullCalcOnLoad='false'></calcPr>", '<calcPr fullCalcOnLoad="1"></calcPr>'),
        ("", '<calcPr fullCalcOnLoad="1"/></workbook>'),
        ("<definedNames/><extLst/>", '<definedNames/><calcPr fullCalcOnLoad="1"/><extLst/>'),
    ],
)
def test_full_calc_on_load_is_set(tmp_path, workbook, expected):
    data = package(ROWS)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    text = parts["xl/workbook.xml"].decode()
    parts["xl/workbook.xml"] = text.replace("</sheets>", "</sheets>" + workbook).encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)
    src = save(tmp_path, buf.getvalue())
    out = tmp_path / "out.xlsx"
    write(src, out, {})
    assert expected in workbook_xml(out)
    with Book(out) as book:
        assert book.calc.full_calc_on_load
    write(src, tmp_path / "kept.xlsx", {}, full_calc_on_load=False)
    assert raw_entries(src) == raw_entries(tmp_path / "kept.xlsx")


def test_prefixed_workbook_gets_a_prefixed_calc_pr(tmp_path):
    data = package(ROWS)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts["xl/workbook.xml"] = (
        f'<x:workbook xmlns:x="{MAIN}" xmlns:r="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships"><x:sheets><x:sheet name="Data" sheetId="1" '
        'r:id="rId1"/></x:sheets></x:workbook>'
    ).encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
    src = save(tmp_path, buf.getvalue())
    write(src, tmp_path / "out.xlsx", {})
    assert '<x:calcPr fullCalcOnLoad="1"/></x:workbook>' in workbook_xml(tmp_path / "out.xlsx")


def test_in_place_and_atomic(tmp_path):
    src = save(tmp_path, package(ROWS))
    os.chmod(src, 0o640)
    before = src.read_bytes()

    def broken(row, col):
        raise RuntimeError("boom")

    with pytest.raises(RuntimeError):
        write(src, src, {"Data": broken})
    assert src.read_bytes() == before
    assert sorted(p.name for p in tmp_path.iterdir()) == ["book.xlsx"]
    write(src, src, {"Data": fixed(B1=("n", 6.0))})
    assert cells_of(src)["B1"] == ("n", 6.0, True)
    assert (src.stat().st_mode & 0o777) == 0o640
    assert sorted(p.name for p in tmp_path.iterdir()) == ["book.xlsx"]


def test_entries_with_data_descriptors_are_copied(tmp_path):
    class Unseekable(io.RawIOBase):
        def __init__(self) -> None:
            self.data = bytearray()

        def writable(self) -> bool:
            return True

        def write(self, b) -> int:  # type: ignore[override]
            self.data += b
            return len(b)

    sink = Unseekable()
    with zipfile.ZipFile(io.BytesIO(package(ROWS))) as z, zipfile.ZipFile(sink, "w") as out:
        for info in z.infolist():
            with out.open(info.filename, "w") as f:
                f.write(z.read(info))
    src = save(tmp_path, bytes(sink.data))
    with zipfile.ZipFile(src) as z:
        assert all(i.flag_bits & 0x08 for i in z.infolist())
    out_path = tmp_path / "out.xlsx"
    write(src, out_path, {"Data": fixed(B1=("n", 6.0))})
    with zipfile.ZipFile(out_path) as z:
        assert z.testzip() is None
    assert cells_of(out_path)["B1"] == ("n", 6.0, True)


def test_zip64_entries_and_directory(tmp_path):
    path = tmp_path / "many.zip"
    with open(path, "w+b") as f:
        zw = ZipWriter(f, b"comment")
        info = zipfile.ZipInfo("big.xml", (2024, 1, 2, 3, 4, 6))
        zw.start(info, zip64=True)
        zw.write(b"<a/>" * 1000)
        zw.finish()
        for k in range(70_000):
            zw.start(zipfile.ZipInfo(f"p{k}"))
            zw.write(b"x")
            zw.finish()
        zw.close()
    with zipfile.ZipFile(path) as z:
        assert len(z.infolist()) == 70_001
        assert z.read("big.xml") == b"<a/>" * 1000
        assert z.getinfo("big.xml").date_time == (2024, 1, 2, 3, 4, 6)
        assert z.read("p69999") == b"x"
        assert z.comment == b"comment"


def test_patch_cells_streams_across_chunk_boundaries():
    cells = "".join(
        f'<c r="A{r}"><f>ROW()</f><v>0</v></c><c r="B{r}" t="s"><v>0</v></c>'
        for r in range(1, 3001)
    )
    xml = f'<worksheet xmlns="{MAIN}"><sheetData><row>{cells}</row></sheetData></worksheet>'
    out = io.BytesIO()
    n = patch_cells(io.BytesIO(xml.encode()), out.write, lambda r, c: ("n", float(r)))
    assert n == 3000
    expected = xml
    for r in range(1, 3001):
        expected = expected.replace(
            f'<c r="A{r}"><f>ROW()</f><v>0</v>', f'<c r="A{r}"><f>ROW()</f><v>{r}</v>'
        )
    assert out.getvalue().decode() == expected


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
