"""Crafted packages: what the reader refuses, before or while it parses, and what it still reads.

The packages are built here, small, from raw parts (no openpyxl). Excel's own behavior with such
packages (it refuses a DTD, a misstated size, other compression methods and a part named twice)
is captured in sheetpeek's ``tests/captures/package.tsv``, at
https://github.com/sheetmill/sheetpeek/blob/main/tests/captures/package.tsv. A test whose
guard could fail into a huge allocation runs under the memory watchdog (``tests.watchdog.bounded``).
"""

from __future__ import annotations

import io
import struct
import zipfile
import zlib

import pytest

import xlrecalc
from tests.test_xlsx import MAIN, package
from tests.watchdog import bounded
from xlrecalc._xlsx import Book, PackageError

SHEET = "xl/worksheets/sheet1.xml"
DTD = '<!DOCTYPE x [<!ENTITY e "boom">]>'
BODY = (
    f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t>'
    "</is></c></row></sheetData></worksheet>"
)


def replace_part(data: bytes, part: str, body: bytes, method: int = zipfile.ZIP_DEFLATED) -> bytes:
    """A package with one part replaced by ``body`` as it stands, under the given method."""
    with zipfile.ZipFile(io.BytesIO(data)) as src:
        items = [(i.filename, src.read(i.filename)) for i in src.infolist()]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, content in items:
            if name != part:
                dst.writestr(name, content)
        dst.writestr(part, body, compress_type=method)
    return buf.getvalue()


def load(data: bytes) -> xlrecalc.Workbook:
    return xlrecalc.load(io.BytesIO(data))


# --- DTDs, in any form and any part


@pytest.mark.parametrize(
    "form",
    [
        pytest.param((DTD + BODY).encode(), id="plain"),
        pytest.param(('<?xml version="1.0"?>' + DTD + BODY).encode(), id="declaration"),
        pytest.param(b"\xef\xbb\xbf" + (DTD + BODY).encode(), id="utf-8 bom"),
        pytest.param(("\n  " + DTD + BODY).encode(), id="whitespace"),
        pytest.param(("<!--" + "c" * 200_000 + "-->" + DTD + BODY).encode(), id="after a comment"),
        pytest.param(("<?p " + "c" * 900_000 + "?>" + DTD + BODY).encode(), id="after a pi"),
        pytest.param((DTD + BODY).encode("utf-16"), id="utf-16 bom"),
        pytest.param(b"\xfe\xff" + (DTD + BODY).encode("utf-16-be"), id="utf-16-be bom"),
        pytest.param(
            ('<?xml version="1.0" encoding="utf-16"?>' + DTD + BODY).encode("utf-16"),
            id="utf-16 declared",
        ),
        pytest.param(
            ('<?xml version="1.0" encoding="iso-8859-1"?>' + DTD + BODY).encode("latin-1"),
            id="latin-1 declared",
        ),
        pytest.param(
            b"<!--" + b"c" * 65_530 + b"-->" + b"<!DOCTYPE x>" + BODY.encode(), id="split"
        ),
    ],
)
@pytest.mark.parametrize("method", [zipfile.ZIP_DEFLATED, zipfile.ZIP_STORED])
def test_a_dtd_is_refused_in_any_form(form, method):
    with pytest.raises(PackageError, match="declares a DTD"):
        load(replace_part(package(), SHEET, form, method))


@pytest.mark.parametrize("part", ["_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"])
def test_a_dtd_is_refused_in_any_part_read(part):
    with pytest.raises(PackageError, match="declares a DTD"):
        load(replace_part(package(), part, b"<!DOCTYPE x><x/>"))


def test_a_dtd_is_refused_in_shared_strings():
    data = package('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', strings=["<si><t>a</t></si>"])
    sst = f'{DTD}<sst xmlns="{MAIN}"><si><t>&e;</t></si></sst>'.encode()
    with pytest.raises(PackageError, match="declares a DTD"):
        load(replace_part(data, "xl/sharedStrings.xml", sst))


def test_a_dtd_is_refused_in_a_sheet_whatever_its_name():
    # The relationships decide which part is a worksheet, so the name decides nothing.
    data = package()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        rels = z.read("xl/_rels/workbook.xml.rels")
    rels = rels.replace(b"sheet1.xml", b"data.bin")
    data = replace_part(data, "xl/_rels/workbook.xml.rels", rels)
    data = replace_part(data, "xl/worksheets/data.bin", (DTD + BODY).encode())
    with pytest.raises(PackageError, match=r"data\.bin: declares a DTD"):
        load(data)


def test_a_billion_laughs_is_refused_whatever_expat_does():
    entities = '<!ENTITY a "aaaaaaaaaa">' + "".join(
        f'<!ENTITY {chr(98 + i)} "{("&" + chr(97 + i) + ";") * 10}">' for i in range(9)
    )
    use = BODY.replace("<t>x</t>", "<t>&j;</t>")
    with pytest.raises(PackageError, match="declares a DTD"):
        load(replace_part(package(), SHEET, f"<!DOCTYPE x [{entities}]>{use}".encode()))


@pytest.mark.parametrize(
    "form",
    [
        pytest.param(('<?xml version="1.0"?>' + BODY).encode(), id="declaration"),
        pytest.param(BODY.encode("utf-16"), id="utf-16"),
        pytest.param(("<!-- <!DOCTYPE x> -->" + BODY).encode(), id="in a comment"),
        pytest.param(("<?p <!DOCTYPE x>?>" + BODY).encode(), id="in a pi"),
        pytest.param(BODY.replace("<t>x</t>", "<t>&lt;!DOCTYPE x&gt;</t>").encode(), id="in text"),
        pytest.param(
            BODY.replace("<t>x</t>", "<t><![CDATA[<!DOCTYPE x>]]></t>").encode(), id="in cdata"
        ),
        pytest.param(("<!--" + "c" * 200_000 + "-->" + BODY).encode(), id="after a long comment"),
    ],
)
def test_the_word_alone_is_not_a_dtd(form):
    book = load(replace_part(package(), SHEET, form))
    assert book.value("A1") in ("x", "<!DOCTYPE x>")


def test_a_declared_encoding_is_still_read():
    body = BODY.replace("<t>x</t>", "<t>café</t>")
    form = ('<?xml version="1.0" encoding="ISO-8859-1"?>' + body).encode("latin-1")
    assert load(replace_part(package(), SHEET, form)).value("A1") == "café"


@pytest.mark.parametrize("body", [b"\x89PNG\r\n\x1a\n" + b"\0" * 2000, b"", b"plain text"])
def test_a_part_that_is_not_xml_passes(body):
    data = package('<row r="1"><c r="A1"><v>7</v></c></row>')
    assert load(replace_part(data, "xl/media/image1.png", body)).value("A1") == 7.0


def test_a_long_prolog_is_refused():
    form = b"<!--" + b"c" * (2 << 20) + b"-->" + BODY.encode()
    with pytest.raises(PackageError, match="no root element"):
        load(replace_part(package(), SHEET, form))


@pytest.mark.parametrize("encoding", ["gb2312", "shift_jis", "utf-7", "bogus"])
def test_an_encoding_expat_cannot_read_is_refused(encoding):
    form = f'<?xml version="1.0" encoding="{encoding}"?>{BODY}'.encode()
    with pytest.raises(PackageError):
        load(replace_part(package(), SHEET, form))


# --- malformed parts


@pytest.mark.parametrize(
    "form",
    [
        pytest.param(b"<!doctype w>" + BODY.encode(), id="lower-case doctype"),
        pytest.param(BODY.encode()[:-20], id="truncated"),
        pytest.param(BODY.replace("<t>x</t>", "<t>&e;</t>").encode(), id="undefined entity"),
        pytest.param((BODY + DTD).encode(), id="a dtd after the root"),
    ],
)
def test_a_malformed_sheet_is_a_package_error(form):
    with pytest.raises(PackageError):
        load(replace_part(package(), SHEET, form))


def test_malformed_shared_strings_are_a_package_error():
    data = package('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', strings=["<si><t>a</t></si>"])
    with pytest.raises(PackageError, match="sharedStrings"):
        load(replace_part(data, "xl/sharedStrings.xml", b"<sst><si><t>a</si></sst>"))


def test_the_reader_raises_the_same():
    with (
        pytest.raises(PackageError, match="declares a DTD"),
        Book(io.BytesIO(replace_part(package(), SHEET, (DTD + BODY).encode()))) as book,
    ):
        list(book.cells("Data"))


# --- the archive: sizes, methods, encryption, damage


def state_size(data: bytes, part: str, size: int, crc: int | None = None) -> bytes:
    """A package whose central directory states another size (and CRC) for one part."""
    buf = bytearray(data)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        position = archive.start_dir
    while buf[position : position + 4] == b"PK\x01\x02":
        name_length, extra_length, comment_length = struct.unpack_from("<HHH", buf, position + 28)
        if bytes(buf[position + 46 : position + 46 + name_length]) == part.encode():
            struct.pack_into("<I", buf, position + 24, size)
            if crc is not None:
                struct.pack_into("<I", buf, position + 16, crc)
            return bytes(buf)
        position += 46 + name_length + extra_length + comment_length
    raise KeyError(part)


def sheet_bytes(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return z.read(SHEET)


@pytest.mark.parametrize("prefix_crc", [False, True], ids=["same crc", "crc of the prefix"])
def test_a_part_smaller_than_its_bytes_is_refused(prefix_crc):
    # zipfile stops at the stated size: with the prefix's CRC, a cut at the right place would
    # have been read as a complete, shorter sheet.
    data = package('<row r="1"><c r="A1"><v>2</v></c></row>')
    xml = sheet_bytes(data)
    crc = zlib.crc32(xml[:-20]) if prefix_crc else None
    with pytest.raises(PackageError, match="the archive states"):
        load(state_size(data, SHEET, len(xml) - 20, crc))


def test_a_part_larger_than_stated_is_refused():
    data = package()
    with pytest.raises(PackageError, match="the archive states"):
        load(state_size(data, SHEET, len(sheet_bytes(data)) + 1000))


@bounded
def test_bytes_hidden_past_the_stated_size_are_refused():
    data = package('<row r="1"><c r="A1"><v>2</v></c></row>')
    xml = sheet_bytes(data)
    data = replace_part(data, SHEET, xml + b" " * (64 << 20))
    with pytest.raises(PackageError, match="the archive states"):
        load(state_size(data, SHEET, len(xml), zlib.crc32(xml)))


@pytest.mark.parametrize("method", [zipfile.ZIP_BZIP2, zipfile.ZIP_LZMA])
def test_other_compression_methods_are_refused(method):
    data = package()
    with pytest.raises(PackageError, match="compression method"):
        load(replace_part(data, SHEET, sheet_bytes(data), method))


def test_an_encrypted_part_is_refused():
    data = bytearray(package())
    with zipfile.ZipFile(io.BytesIO(bytes(data))) as archive:
        data[archive.start_dir + 8] |= 1
    with pytest.raises(PackageError, match="encrypted"):
        load(bytes(data))


def test_damaged_data_is_refused():
    data = package()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        info = archive.getinfo(SHEET)
    buf = bytearray(data)
    start = info.header_offset + 30 + len(info.filename)
    buf[start : start + 8] = b"\xff" * 8
    with pytest.raises(PackageError, match="damaged"):
        load(bytes(buf))


def test_a_part_read_twice_at_once_reads_the_same():
    data = package('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', strings=["<si><t>a</t></si>"])
    with Book(io.BytesIO(data)) as book:
        whole = sheet_bytes(data)
        with book._package.part(SHEET) as first, book._package.part(SHEET) as second:
            head = first.read(10)
            assert second.read() == whole
            assert head + first.read() == whole
        assert [c.value for c in book.cells("Data")] == ["a"]


# --- the grid: cells far away, huge ranges, rows and columns off the sheet


@bounded
def test_a_stray_cell_at_the_last_position_costs_nothing():
    data = package(
        '<row r="1"><c r="A1"><v>2</v></c></row>'
        '<row r="1048576"><c r="XFD1048576"><f>SUM(A1:XFC1048575)</f></c></row>'
    )
    book = load(data)
    assert book.value("XFD1048576") == 2.0


@pytest.mark.parametrize(
    "extra",
    [
        pytest.param(
            '<row r="2"><c r="A2"><f t="dataTable" ref="A2:XFD1048576" dt2D="1" r1="B1" r2="C1"/>'
            "</c></row>",
            id="data table over the sheet",
        ),
        pytest.param(
            '<row r="2"><c r="A2"><f t="shared" ref="A2:XFD1048576" si="0">A1</f></c></row>',
            id="shared formula range over the sheet",
        ),
    ],
)
@bounded
def test_a_huge_range_in_a_cell_costs_nothing(extra):
    book = load(package('<row r="1"><c r="A1"><v>2</v></c></row>' + extra))
    assert book.value("A1") == 2.0


@bounded
def test_a_huge_merge_costs_nothing():
    sheet = (
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>2</v></c></row>'
        '</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>'
        "</worksheet>"
    )
    assert load(package(sheet_xml=sheet)).value("A1") == 2.0


@pytest.mark.parametrize(
    "rows",
    [
        '<row r="0"><c r="A0"><v>1</v></c></row>',
        '<row r="1048577"><c r="A1048577"><v>1</v></c></row>',
        '<row r="99999999999"><c r="A99999999999"><v>1</v></c></row>',
        '<row r="99999999999"><c><v>1</v></c></row>',
        '<row r="-5"><c><v>1</v></c></row>',
        '<row r="x"><c><v>1</v></c></row>',
        '<row r="1"><c r="A99999999999"><v>1</v></c></row>',
    ],
)
@bounded
def test_a_row_off_the_sheet_is_refused(rows):
    with pytest.raises(PackageError, match="row"):
        load(package(rows))


def test_a_column_off_the_sheet_is_placed_by_position():
    # Excel opens it and reads the cell where it stands (capture value_forms, "column XFE").
    book = load(package('<row r="1"><c r="XFE1"><v>7</v></c><c r="ZZZZZZZZZZ1"><v>8</v></c></row>'))
    assert (book.value("A1"), book.value("B1")) == (7.0, 8.0)


@bounded
def test_more_cells_than_columns_is_refused():
    cells = "<c><v>1</v></c>" * 16_385
    with pytest.raises(PackageError, match="more cells than columns"):
        load(package(f'<row r="1">{cells}</row>'))


# --- numbers and indexes in the attributes and values the reader uses


@pytest.mark.parametrize(
    "text", ["nan", "NaN", "inf", "-inf", "Infinity", "1_000", "0x10", "+5", "1e"]
)
def test_a_number_excel_refuses_is_refused(text):
    with pytest.raises(PackageError, match="number"):
        load(package(f'<row r="1"><c r="A1"><v>{text}</v></c></row>'))


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("1e400", xlrecalc.ErrorValue.NULL),
        ("-1e400", xlrecalc.ErrorValue.NULL),
        ("1e308", 1e308),
        ("4.9e-324", 0.0),
        (" 5", 5.0),
        ("5.", 5.0),
        (".5", 0.5),
        ("1E+2", 100.0),
        ("", None),
    ],
)
def test_numbers_as_excel_reads_them(text, expected):
    assert load(package(f'<row r="1"><c r="A1"><v>{text}</v></c></row>')).value("A1") == expected


@pytest.mark.parametrize("index", ["-1", "1", "99999999999", "nan"])
def test_a_shared_string_that_does_not_exist_is_refused(index):
    data = package(
        f'<row r="1"><c r="A1" t="s"><v>{index}</v></c></row>', strings=["<si><t>a</t></si>"]
    )
    with pytest.raises(PackageError, match="no shared string"):
        load(data)


def test_a_shared_formula_index_that_is_not_a_number_is_refused():
    rows = '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="nan">1</f></c></row>'
    with pytest.raises(PackageError, match="shared formula index"):
        load(package(rows))


def test_a_sheet_scope_that_is_not_a_number_is_refused():
    names = '<definedNames><definedName name="x" localSheetId="nan">Data!$A$1</definedName>'
    sheets = f'<sheet name="Data" sheetId="1" r:id="rId1"/></sheets>{names}</definedNames><sheets>'
    data = package(sheets=sheets)
    with pytest.raises(PackageError, match="localSheetId"):
        load(data)


@pytest.mark.parametrize(
    "attribute",
    [
        'calcId="nan"', 'iterateCount="nan"', 'iterateCount="99999999999"',
        'iterateDelta="nan"', 'iterateDelta="inf"', 'calcId="-1"',
    ],
)  # fmt: skip
def test_calculation_settings_that_are_not_numbers_are_ignored(attribute):
    calc = f'<sheet name="Data" sheetId="1" r:id="rId1"/></sheets><calcPr {attribute}/><sheets>'
    book = load(package('<row r="1"><c r="A1"><v>2</v></c></row>', sheets=calc))
    assert book.value("A1") == 2.0


@pytest.mark.parametrize(
    "cell",
    [
        '<c r="A1" s="nan"><v>2</v></c>',
        '<c r="A1" s="-1"><v>2</v></c>',
        '<c r="A1" cm="99999999999"><v>2</v></c>',
        '<c r="A1" vm="nan"><v>2</v></c>',
    ],
)
def test_style_and_metadata_indexes_that_lead_nowhere_are_harmless(cell):
    assert load(package(f'<row r="1">{cell}</row>')).value("A1") == 2.0


@pytest.mark.parametrize("tab", ["99999999999", "-1", "nan"])
def test_an_active_tab_that_leads_nowhere_is_harmless(tab):
    views = f'<bookViews><workbookView activeTab="{tab}"/></bookViews>'
    assert (
        load(package('<row r="1"><c r="A1"><v>2</v></c></row>', workbook_extra=views)).value("A1")
        == 2.0
    )
