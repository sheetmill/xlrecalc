"""OFFSET, INDIRECT, CELL and INDEX as a reference, against Excel (capture ``dynrefs``, §6.3).

The probe's workbook is rebuilt from the capture (data, three sheets, the name ``Rate``, every
formula in its cell), loaded with the path ``<dir>/<book>`` (the capture writes the file's
directory and name so), recalculated, and every captured formula cell must match.
"""

from __future__ import annotations

import ast
import io
import re
import zipfile
from pathlib import Path

import pytest

from tests.test_captures import same
from xlrecalc._engine import Engine
from xlrecalc._formula._tokens import column_number
from xlrecalc._load import load, refused_in_file
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._xlsx._write import text_escape

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
ROWS = [
    line.rstrip("\n").split("\t")
    for line in (Path(__file__).with_name("captures") / "dynrefs.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if not line.startswith("#")
]


def _cell(ref: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    assert m is not None
    return int(m[2]), column_number(m[1])


def _sheet(cells: dict[tuple[int, int], str]) -> str:
    rows: dict[int, list[str]] = {}
    for (r, _c), xml in sorted(cells.items()):
        rows.setdefault(r, []).append(xml)
    data = "".join(f'<row r="{r}">{"".join(x)}</row>' for r, x in sorted(rows.items()))
    return f'<worksheet xmlns="{MAIN}"><sheetData>{data}</sheetData></worksheet>'


def workbook() -> bytes:
    cells: dict[tuple[int, int], str] = {}
    for r in range(1, 6):
        cells[r, 26] = f'<c r="Z{r}"><v>{r}</v></c>'
        cells[r, 27] = f'<c r="AA{r}"><v>{10 * r}</v></c>'
    cells[1, 28] = '<c r="AB1" t="inlineStr"><is><t>text</t></is></c>'
    for where, formula, *_ in ROWS:
        r, c = _cell(where)
        cells[r, c] = f'<c r="{where}"><f>{text_escape(ast.literal_eval(formula))}</f></c>'
    sheets = [("S", _sheet(cells)), ("Sheet2", _sheet({(1, 1): '<c r="A1"><v>100</v></c>'}))]
    sheets.append(("My Sheet", _sheet({(1, 1): '<c r="A1"><v>200</v></c>'})))
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types"/>',
        "_rels/.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
        f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>'
        + "".join(
            f'<sheet name="{n}" sheetId="{k}" r:id="rId{k}"/>' for k, (n, _) in enumerate(sheets, 1)
        )
        + '</sheets><definedNames><definedName name="Rate">S!$Z$2</definedName></definedNames>'
        "</workbook>",
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">'
        + "".join(
            f'<Relationship Id="rId{k}" Type="{REL}/worksheet" Target="worksheets/sheet{k}.xml"/>'
            for k in range(1, len(sheets) + 1)
        )
        + "</Relationships>",
    }
    for k, (_, xml) in enumerate(sheets, 1):
        parts[f"xl/worksheets/sheet{k}.xml"] = xml
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return buf.getvalue()


@pytest.fixture(scope="module")
def engine() -> Engine:
    e = load(io.BytesIO(workbook()))
    e.path = "<dir>/<book>"
    e.recalculate()
    return e


@pytest.mark.excel
@pytest.mark.parametrize("row", ROWS, ids=lambda r: f"{r[0]}:{ast.literal_eval(r[1])[:40]}")
def test_matches_excel(engine, row):
    where, _formula, kind, text = row
    r, c = _cell(where)
    tid = engine.sheets[0].templates.get((r - 1) * 16384 + c - 1)
    if tid is not None and engine.templates[tid].unsupported:
        pytest.skip("not implemented: " + ", ".join(sorted(engine.templates[tid].unsupported)))
    formatting = re.search(r'CELL\("(format|protect|width|prefix|color|parentheses)"', _formula)
    if formatting:  # counted, not evaluated: the engine reads no styles yet
        assert f'CELL("{formatting[1]}")' in engine.report.unsupported
        pytest.skip(f'not implemented: CELL("{formatting[1]}"), which depends on formatting')
    value = ast.literal_eval(text)
    expected = ErrorValue(value) if kind == "e" else value
    got = engine.get(0, r - 1, c - 1)
    if kind == "none":
        assert got is BLANK
    else:
        assert same(got, expected), (where, got, expected)


def test_offset_with_a_constant_reference_is_refused_in_a_file():
    assert refused_in_file("OFFSET(5,0,0)") is not None
    assert refused_in_file('OFFSET("a",0,0)') is not None
    assert refused_in_file("OFFSET(A1,0,0)") is None


def test_volatile_references_follow_a_change():
    e = Engine.from_cells({"S": {"A1": 1, "A2": 2, "B1": '=INDIRECT("A"&C1)', "C1": 1}})
    e.recalculate()
    assert e.get(0, 0, 1) == 1.0
    e.set_value(0, 0, 2, 2.0)  # C1 = 2: B1 now reads A2, which no static reference shows
    assert e.get(0, 0, 1) == 2.0
    e.set_value(0, 1, 0, 20.0)  # A2 = 20
    assert e.get(0, 0, 1) == 20.0


def test_a_dynamic_precedent_is_evaluated_first():
    e = Engine.from_cells(
        {"S": {"A1": '=INDIRECT("A3")+OFFSET(A1,3,0)', "A3": "=A4*2", "A4": "=1+1"}}
    )
    e.recalculate()
    assert e.get(0, 0, 0) == 6.0  # A3 = 4 and A4 = 2, evaluated before A1 read them


def test_a_cycle_through_a_dynamic_reference_is_circular():
    e = Engine.from_cells({"S": {"A1": '=INDIRECT("B1")', "B1": "=A1+1", "C1": "=B1"}})
    report = e.recalculate()
    assert e.get(0, 0, 0) == e.get(0, 0, 1) == e.get(0, 0, 2) == ""
    assert {(r.row, r.col) for r in report.circular} == {(1, 1), (1, 2), (1, 3)}


NOREF = [
    line.rstrip("\n").split("\t")
    for line in (Path(__file__).with_name("captures") / "cell_noref.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if not line.startswith("#")
]


def _noref_workbook(active: int, cells: list[str]) -> bytes:
    formulas = ['CELL("address")', 'CELL("filename")', 'CELL("row")', 'CELL("col")']
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types"/>',
        "_rels/.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
        f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews>'
        f'<workbookView activeTab="{active}"/></bookViews><sheets>'
        + "".join(f'<sheet name="S{k}" sheetId="{k}" r:id="rId{k}"/>' for k in (1, 2, 3))
        + "</sheets></workbook>",
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">'
        + "".join(
            f'<Relationship Id="rId{k}" Type="{REL}/worksheet" Target="worksheets/sheet{k}.xml"/>'
            for k in (1, 2, 3)
        )
        + "</Relationships>",
    }
    for k, cell in enumerate(cells, 1):
        rows = "".join(
            f'<row r="{r}"><c r="A{r}"><f>{text_escape(f)}</f></c></row>'
            for r, f in enumerate(formulas, 1)
        )
        parts[f"xl/worksheets/sheet{k}.xml"] = (
            f'<worksheet xmlns="{MAIN}"><sheetViews><sheetView workbookViewId="0">'
            f'<selection activeCell="{cell}" sqref="{cell}"/></sheetView></sheetViews>'
            f"<sheetData>{rows}</sheetData></worksheet>"
        )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return buf.getvalue()


@pytest.mark.excel
@pytest.mark.parametrize("variant", ["0", "1", "2"])
def test_cell_without_a_reference_reads_the_active_cell(variant):
    """Capture cell_noref: the active sheet's active cell, from every sheet."""
    rows = [r for r in NOREF if r[0] == variant]
    active, cells = int(rows[0][1]), rows[0][2].split(",")
    e = load(io.BytesIO(_noref_workbook(active, cells)))
    e.path = "<dir>/<book>"
    e.recalculate()
    for _, _, _, sheet, formula, value in rows:
        k = int(sheet[1]) - 1
        r = ['CELL("address")', 'CELL("filename")', 'CELL("row")', 'CELL("col")'].index(
            ast.literal_eval(formula)
        )
        assert e.get(k, r, 0) == ast.literal_eval(value), (variant, sheet, formula)
