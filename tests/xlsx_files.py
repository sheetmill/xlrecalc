"""Minimal xlsx packages built from raw parts, for tests that feed the reader crafted files."""

from __future__ import annotations

import io
import zipfile

import xlio

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def package(
    sheet_data: str = "",
    *,
    strings: list[str] | None = None,
    workbook_extra: str = "",
    sheets: str = '<sheet name="Data" sheetId="1" r:id="rId1"/>',
    extra_parts: dict[str, str | bytes] | None = None,
    sheet_xml: str | None = None,
    sheet_rels: str | None = None,
) -> bytes:
    """A minimal xlsx package built from raw parts."""
    parts: dict[str, str | bytes] = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        "_rels/.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
        f'Type="{DOC}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}">{workbook_extra}'
        f"<sheets>{sheets}</sheets></workbook>",
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{DOC}/chartsheet" Target="chartsheets/sheet1.xml"/>'
        f'<Relationship Id="rId9" Type="{DOC}/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>",
        "xl/worksheets/sheet1.xml": sheet_xml
        or f'<worksheet xmlns="{MAIN}"><sheetData>{sheet_data}</sheetData></worksheet>',
    }
    if strings is not None:
        parts["xl/sharedStrings.xml"] = f'<sst xmlns="{MAIN}">' + "".join(strings) + "</sst>"
    if sheet_rels is not None:
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = sheet_rels
    parts.update(extra_parts or {})
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return buf.getvalue()


def cells(data: bytes) -> list[xlio.Cell]:
    """The cells of the package's sheet ``Data``, as ``xlio`` reads them."""
    with xlio.open(io.BytesIO(data)) as book:
        return list(book.sheet("Data").cells())
