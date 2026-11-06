"""The xlsx reader (ARCHITECTURE.md §9.1): what it reads, its limits, and its boundary."""

from __future__ import annotations

import ast
import io
import sys
import zipfile
from pathlib import Path

import pytest

from xlrecalc._xlsx import Book, LimitError, PackageError, ReadLimits, unescape

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main"
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


def cells(data: bytes, limits: ReadLimits | None = None) -> list:
    with Book(io.BytesIO(data), limits) as book:
        return list(book.cells("Data"))


def values(data: bytes) -> dict[str, tuple]:
    out = {}
    for c in cells(data):
        letters = "ABCDEFGHIJ"[c.column - 1]
        out[f"{letters}{c.row}"] = (c.value, c.type, c.formula.text if c.formula else None)
    return out


# What it reads ------------------------------------------------------------------------------


def test_value_types():
    data = package(
        '<row r="1">'
        '<c r="A1"><v>1.5</v></c>'
        '<c r="B1" t="n"><v>-0</v></c>'
        '<c r="C1" t="s"><v>0</v></c>'
        '<c r="D1" t="b"><v>1</v></c>'
        '<c r="E1" t="e"><v>#N/A</v></c>'
        '<c r="F1" t="inlineStr"><is><t>inline</t></is></c>'
        '<c r="G1" t="d"><v>2025-03-04T10:30:00</v></c>'
        '<c r="H1" s="3"/>'
        "</row>",
        strings=["<si><t>shared</t></si>"],
    )
    assert values(data) == {
        "A1": (1.5, "n", None),
        "B1": (-0.0, "n", None),
        "C1": ("shared", "s", None),
        "D1": (True, "b", None),
        "E1": ("#N/A", "e", None),
        "F1": ("inline", "inlineStr", None),
        "G1": ("2025-03-04T10:30:00", "d", None),
        "H1": (None, "none", None),
    }


def test_formulas_and_cached_values():
    data = package(
        '<row r="1">'
        '<c r="A1"><f>1+1</f><v>2</v></c>'
        '<c r="B1"><f>A1*3</f><v /></c>'  # openpyxl's uncached form
        '<c r="C1" t="str"><f>""</f><v/></c>'  # Excel's empty text result
        '<c r="D1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>'
        '<c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
        "</row>"
    )
    v = values(data)
    assert v["A1"] == (2.0, "n", "1+1")
    assert v["B1"] == (None, "none", "A1*3")
    assert v["C1"] == ("", "str", '""')
    assert v["D1"] == ("ab", "str", '"a"&"b"')
    assert v["E1"] == ("#DIV/0!", "e", "1/0")


def test_shared_and_array_formulas():
    data = package(
        '<row r="1"><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>'
        '<row r="2"><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>'
        '<row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c>'
        '<c r="C3"><f t="array" ref="C3:C4" aca="false">A1:A2*2</f><v>2</v></c></row>'
    )
    by = {(c.row, c.column): c for c in cells(data)}
    master, dependent, array = by[(1, 2)].formula, by[(2, 2)].formula, by[(3, 3)].formula
    assert (master.text, master.kind, master.ref, master.si) == ("A1*2", "shared", "B1:B3", 0)
    assert (dependent.text, dependent.kind, dependent.si) == (None, "shared", 0)
    assert (array.text, array.kind, array.ref) == ("A1:A2*2", "array", "C3:C4")
    assert array.attributes["aca"] == "false"


def test_rich_text_phonetic_runs_and_escapes():
    data = package(
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>'
        '<c r="C1" t="inlineStr"><is><r><t>in</t></r><r><t xml:space="preserve"> line</t></r>'
        "<rPh><t>phonetic</t></rPh></is></c>"
        '<c r="D1" t="str"><f>CHAR(13)&amp;"x"</f><v>_x000D_x</v></c>'
        '<c r="E1"><f>LEN("_x0041_")</f><v>1</v></c></row>',
        strings=[
            "<si><r><t>bo</t></r><r><rPr><b/></rPr><t>ld</t></r><rPh><t>ふ</t></rPh></si>",
            "<si><t>a_x000D_b _x005F_x0041_ c_x000d_</t></si>",
        ],
    )
    v = values(data)
    assert v["A1"][0] == "bold"
    assert v["B1"][0] == "a\rb _x0041_ c\r"
    assert v["C1"][0] == "in line"
    assert v["D1"][0] == "\rx"
    assert v["E1"][2] == 'LEN("A")'  # Excel decodes escapes in formula text (escape probe)


def test_unescape():
    assert unescape("a_x000D_b") == "a\rb"
    assert unescape("_x005F_x0041_") == "_x0041_"
    assert unescape("a_x000Db") == "a_x000Db"
    assert unescape("plain") == "plain"


def test_positions_without_r_attributes():
    data = package(
        '<row><c><v>1</v></c><c><v>2</v></c></row><row r="5"><c><v>3</v></c><c r="D5"><v>4</v></c>'
        "<c><v>5</v></c></row><row><c><v>6</v></c></row>"
    )
    assert [(c.row, c.column) for c in cells(data)] == [
        (1, 1),
        (1, 2),
        (5, 1),
        (5, 4),
        (5, 5),
        (6, 1),
    ]


def test_prefixed_and_strict_namespaces():
    prefixed = (
        f'<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1">'
        '<x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:f>A1</x:f><x:v>1</x:v></x:c>'
        "</x:row></x:sheetData></x:worksheet>"
    )
    data = package(sheet_xml=prefixed, strings=["<si><t>p</t></si>"])
    assert values(data) == {"A1": ("p", "s", None), "B1": (1.0, "n", "A1")}
    strict = (
        f'<worksheet xmlns="{STRICT}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row>'
        "</sheetData></worksheet>"
    )
    assert values(package(sheet_xml=strict)) == {"A1": (7.0, "n", None)}


def test_workbook_properties():
    # Calculation settings and names come after <sheets> in the schema.
    data = package(
        workbook_extra='<workbookPr date1904="1"/>',
        sheets='<sheet name="Data" sheetId="1" r:id="rId1" state="hidden"/>'
        '<sheet name="Chart" sheetId="2" r:id="rId2"/></sheets><definedNames>'
        '<definedName name="Rate">Data!$A$1</definedName>'
        '<definedName name="Local" localSheetId="0" hidden="1">Data!$B$2</definedName>'
        '</definedNames><calcPr calcId="191029" fullCalcOnLoad="1" iterate="1" '
        'iterateCount="50" iterateDelta="0.01"/><sheets>',
        extra_parts={"xl/chartsheets/sheet1.xml": "<chartsheet/>"},
    )
    with Book(io.BytesIO(data)) as book:
        assert book.date1904
        assert [(s.name, s.state, s.kind) for s in book.sheets] == [
            ("Data", "hidden", "worksheet"),
            ("Chart", "visible", "chartsheet"),
        ]
        assert book.sheet("data").name == "Data"  # without case, as Excel
        with pytest.raises(ValueError, match="chartsheet"):
            list(book.cells("Chart"))
        assert [(n.name, n.text, n.sheet, n.hidden) for n in book.defined_names] == [
            ("Rate", "Data!$A$1", None, False),
            ("Local", "Data!$B$2", 0, True),
        ]
        calc = book.calc
        assert (calc.calc_id, calc.full_calc_on_load, calc.iterate) == (191029, True, True)
        assert (calc.iterate_count, calc.iterate_delta) == (50, 0.01)


def test_tables():
    table = (
        f'<table xmlns="{MAIN}" id="1" name="Table1" displayName="Sales" ref="A1:B3" '
        'totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Price"/>'
        '<tableColumn id="2" name="Unit Price"/></tableColumns></table>'
    )
    rels = (
        f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{DOC}/table" '
        'Target="../tables/table1.xml"/></Relationships>'
    )
    data = package(sheet_rels=rels, extra_parts={"xl/tables/table1.xml": table})
    with Book(io.BytesIO(data)) as book:
        (t,) = book.tables("Data")
        assert (t.name, t.display_name, t.ref, t.columns) == (
            "Table1",
            "Sales",
            "A1:B3",
            ("Price", "Unit Price"),
        )
        assert (t.header_rows, t.totals_rows) == (1, 1)


def test_agrees_with_openpyxl_on_a_generated_workbook(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    rows = [[1, 2.5, "text", True, "=A1+B1"], [None, -3, "ünï", False, '=C1&"x"']]
    for row in rows:
        ws.append(row)
    wb.save(path)
    with Book(path) as book:
        got = {(c.row, c.column): c for c in book.cells("Data")}
    for r, row in enumerate(rows, 1):
        for k, expected in enumerate(row, 1):
            cell = got.get((r, k))
            if expected is None:
                assert cell is None
            elif isinstance(expected, str) and expected.startswith("="):
                assert cell.formula.text == expected[1:]
                assert cell.type == "none"  # openpyxl stores no cached value
            elif isinstance(expected, bool):
                assert cell.value is expected
            else:
                assert cell.value == expected


# Limits and untrusted input -------------------------------------------------------------------


def test_not_a_package():
    with pytest.raises(PackageError, match="not a zip"):
        Book(io.BytesIO(b"not a zip at all"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("hello.txt", "x")
    with pytest.raises(PackageError):
        Book(io.BytesIO(buf.getvalue()))


def test_part_count_limit():
    data = package(extra_parts={f"extra/{k}.bin": b"" for k in range(20)})
    with pytest.raises(LimitError, match="parts"):
        Book(io.BytesIO(data), ReadLimits(max_parts=10))
    Book(io.BytesIO(data), ReadLimits(max_parts=100)).close()


def test_part_size_limit_counts_decompressed_bytes():
    big = "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>' for r in range(1, 5001))
    data = package(big)
    assert len(data) < 40_000  # compresses well
    with pytest.raises(LimitError, match="more than"):
        cells(data, ReadLimits(max_part_bytes=50_000))
    assert len(cells(data, ReadLimits(max_part_bytes=1_000_000))) == 5000


def test_total_size_limit():
    big = "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>' for r in range(1, 5001))
    with pytest.raises(LimitError, match="in all"):
        cells(package(big), ReadLimits(max_total_bytes=60_000))


def test_duplicate_part_names():
    buf = io.BytesIO(package())
    with pytest.warns(UserWarning, match="Duplicate name"), zipfile.ZipFile(buf, "a") as z:
        z.writestr("xl/workbook.xml", "<workbook/>")
    with pytest.raises(PackageError, match="twice"):
        Book(io.BytesIO(buf.getvalue()))


LAUGHS = (
    '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol">'
    + "".join(
        f'<!ENTITY lol{i} "' + (f"&lol{i - 1};" if i else "&lol;") * 10 + '">' for i in range(10)
    )
    + "]>"
)
DTDS = {
    "billion laughs": LAUGHS + f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>',
    "quadratic blowup": '<!DOCTYPE q [<!ENTITY a "' + "x" * 5000 + '">]>'
    f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>',
    "external entity": '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>'
    f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&e;</t>'
    "</is></c></row></sheetData></worksheet>",
    "external parameter entity": '<!DOCTYPE x [<!ENTITY % p SYSTEM "http://127.0.0.1:9/x.dtd">'
    " %p;]>"
    f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>',
    "after a comment and a BOM": "﻿<?xml version='1.0'?>\n<!-- note --><!DOCTYPE w>"
    f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>',
}


@pytest.mark.parametrize("name", list(DTDS))
def test_a_dtd_is_refused_in_any_part(name):
    with pytest.raises(PackageError, match="DTD"):
        cells(package(sheet_xml=DTDS[name]))
    with pytest.raises(PackageError, match="DTD"):
        Book(io.BytesIO(package(extra_parts={"xl/workbook.xml": DTDS[name]})))
    sst = DTDS[name].replace("worksheet", "sst").replace("<sheetData/>", "")
    with pytest.raises(PackageError, match="DTD"):
        cells(_with_strings(sst))


def _with_strings(sst: str) -> bytes:
    data = package('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', strings=["<si><t>a</t></si>"])
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(buf, "w") as dst:
        for info in src.infolist():
            dst.writestr(info, sst if info.filename == "xl/sharedStrings.xml" else src.read(info))
    return buf.getvalue()


def test_declaration_and_comments_without_a_dtd_are_fine():
    sheet = (
        "﻿<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<!-- a comment -->"
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
        "</sheetData></worksheet>"
    )
    assert values(package(sheet_xml=sheet)) == {"A1": (1.0, "n", None)}


# Boundary ------------------------------------------------------------------------------------

XLSX = Path(__file__).resolve().parents[1] / "src" / "xlrecalc" / "_xlsx"


@pytest.mark.parametrize("module", sorted(XLSX.glob("*.py")), ids=lambda p: p.name)
def test_standard_library_only(module):
    """_xlsx imports nothing but the standard library and itself, so it can be extracted."""
    tree = ast.parse(module.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names = [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom):
            names = [node.module or ""]
        else:
            continue
        for name in names:
            top = name.split(".")[0]
            if name.startswith("xlrecalc."):
                assert name.startswith("xlrecalc._xlsx"), f"{module.name} imports {name}"
            else:
                assert top in sys.stdlib_module_names or top == "__future__", (
                    f"{module.name} imports {name}"
                )


def test_a_text_result_loses_unpreserved_outer_whitespace():
    """As Excel reads it (tests/test_excel_live.py): ``<v> a </v>`` is ``a``."""
    rows = (
        '<row r="1"><c r="A1" t="str"><f>" a "</f><v> a\n b </v></c>'
        '<c r="B1" t="str"><f>" a "</f><v xml:space="preserve"> a </v></c></row>'
    )
    got = {(c.row, c.column): c.value for c in cells(package(rows))}
    assert got == {(1, 1): "a\n b", (1, 2): " a "}
