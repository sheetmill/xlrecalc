"""Crafted packages: what the reader refuses, before or while it parses, and what it still reads.

The packages are built here, small, from raw parts (no openpyxl). Excel's own behavior with such
packages (it refuses a DTD, a misstated size, other compression methods and a part named twice)
is captured in sheetpeek's ``tests/captures/package.tsv``. A test whose guard could fail into a
huge allocation runs under the memory watchdog (``tests.watchdog.bounded``).
"""

from __future__ import annotations

import io
import zipfile

import pytest

import xlrecalc
from tests.test_xlsx import MAIN, package
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
