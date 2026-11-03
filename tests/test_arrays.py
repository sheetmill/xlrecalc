"""Arrays, spill and implicit intersection against Excel (capture ``arrays``; §7.7).

The workbook ``capture/array_probe.py`` gave Excel is rebuilt from the capture's input rows
(legacy formulas, legacy array formulas over their declared ranges, dynamic-array formulas
with the dynamic-array metadata), loaded, recalculated, and every cell Excel saved must hold
the same value here.
"""

from __future__ import annotations

import ast
import io
import re
import zipfile
from pathlib import Path

import pytest

from tests.test_captures import same
from tests.test_xlsx import MAIN, package
from xlrecalc._engine import node
from xlrecalc._formula._tokens import column_number
from xlrecalc._load import load
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._xlsx._write import text_escape

CAPTURES = Path(__file__).with_name("captures")
ROWS = [
    line.rstrip("\n").split("\t")
    for line in (CAPTURES / "arrays.tsv").read_text(encoding="utf-8").splitlines()
    if not line.startswith("#")
]
INPUTS = [r[1:] for r in ROWS if r[0] == "input"]
CELLS = {r[1]: r[2:] for r in ROWS if r[0] == "cell"}
DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _cell(ref: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    assert m is not None
    return int(m[2]), column_number(m[1])


def workbook() -> bytes:
    """The probe's workbook: data, the in-the-way constant, and every input formula."""
    cells: dict[tuple[int, int], str] = {}
    for r in range(1, 6):
        cells[r, 26] = f'<c r="Z{r}"><v>{r}</v></c>'
        cells[r, 27] = f'<c r="AA{r}"><v>{10 * r}</v></c>'
        cells[r, 28] = f'<c r="AB{r}" t="inlineStr"><is><t>{"abcde"[r - 1]}</t></is></c>'
    cells[3, 19] = '<c r="S3" t="inlineStr"><is><t>in the way</t></is></c>'
    for kind, where, formula in INPUTS:
        text = text_escape(ast.literal_eval(formula))
        first = where.split(":")[0]
        r, c = _cell(first)
        if kind == "cse":
            cells[r, c] = f'<c r="{first}"><f t="array" ref="{where}">{text}</f></c>'
        elif kind == "dynamic":
            cells[r, c] = f'<c r="{first}" cm="1"><f t="array" ref="{first}">{text}</f></c>'
        else:
            cells[r, c] = f'<c r="{first}"><f>{text}</f></c>'
    rows: dict[int, list[str]] = {}
    for (r, _c), xml in sorted(cells.items()):
        rows.setdefault(r, []).append(xml)
    sheet_data = "".join(f'<row r="{r}">{"".join(x)}</row>' for r, x in sorted(rows.items()))
    data = package(sheet_data, extra_parts={"xl/metadata.xml": _metadata()})
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    rels = (
        parts["xl/_rels/workbook.xml.rels"]
        .decode()
        .replace(
            "</Relationships>",
            f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
            "</Relationships>",
        )
    )
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)
    return buf.getvalue()


def _metadata() -> str:
    """The dynamic-array part of the metadata Excel saved (cell metadata 1 is XLDAPR)."""
    text = (CAPTURES / "arrays_metadata.xml").read_text(encoding="utf-8")
    assert f'xmlns="{MAIN}"' in text
    return text


@pytest.fixture(scope="module")
def engine():
    e = load(io.BytesIO(workbook()))
    e.recalculate()
    return e


def expected(cell: list[str]) -> object:
    kind, text, _formula, _t, _ref, _cm, vm = cell
    value = ast.literal_eval(text)
    if kind == "e":
        return ErrorValue("#SPILL!") if vm else ErrorValue(value)  # vm: a rich-value error
    if kind == "n":
        return float(value)
    return value


@pytest.mark.excel
@pytest.mark.parametrize("address", sorted(CELLS, key=lambda a: _cell(a)), ids=str)
def test_cell_matches_excel(engine, address):
    row, col = _cell(address)
    tid = engine.sheets[0].templates.get((row - 1) * 16384 + col - 1)
    if tid is not None and engine.templates[tid].unsupported:
        names = ", ".join(sorted(engine.templates[tid].unsupported))
        pytest.skip(f"not implemented: {names}")  # counted, as in tests/test_captures.py
    got = engine.get(0, row - 1, col - 1)
    want = expected(CELLS[address])
    if want is None:
        assert got is BLANK
        return
    assert same(got, want), (address, got, want)


@pytest.mark.excel
def test_spill_ranges_match_excel(engine):
    """Each dynamic array spills over the range Excel recorded in its ``ref``."""
    for kind, anchor, _ in INPUTS:
        if kind != "dynamic":
            continue
        r, c = _cell(anchor)
        ref = CELLS[anchor][4]
        first, _, last = ref.partition(":")
        r2, c2 = _cell(last or first)
        assert engine.anchors[node(0, r - 1, c - 1)].box == (r - 1, c - 1, r2 - 1, c2 - 1), anchor
