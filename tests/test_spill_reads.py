"""A dynamic array an edit makes stale: a read finds the cells it now spills into (§7.7).

After such an edit Excel shows a blank cell the array now spills into, and a formula reading one,
with the spilled value (capture ``spill_edit``, ``tests/test_spill_edit.py``). Read lazily, they
must show it too, however the cells are read: one by one, in one pass, or through a lazy
``Loader``.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest
import xlio

import xlrecalc
from tests.xlsx_files import DOC, PKG, package

FULL = '<calcPr fullCalcOnLoad="1"/>'
ROWS = (
    '<row r="1"><c r="D1" cm="1"><f t="array" ref="D1:D2">_xlfn.SEQUENCE(E1)</f></c>'
    '<c r="E1"><v>2</v></c><c r="F1"><f>D3+1</f><v>0</v></c></row>'
)


def _data() -> bytes:
    metadata = (Path(__file__).parent / "captures" / "arrays_metadata.xml").read_text("utf-8")
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>"
    )
    parts = {"xl/metadata.xml": metadata, "xl/_rels/workbook.xml.rels": rels}
    return package(ROWS, workbook_extra=FULL, extra_parts=parts)


@pytest.mark.parametrize("way", ["load", "loader"])
@pytest.mark.parametrize("batch", [False, True])
@pytest.mark.parametrize("first", ["D3", "F1"])
def test_a_grown_spill_is_read_before_its_anchor(way: str, batch: bool, first: str) -> None:
    with xlio.open(io.BytesIO(_data())) as book:
        wb = (
            xlrecalc.load(io.BytesIO(_data()))
            if way == "load"
            else xlrecalc.Loader(book).workbook()
        )
        assert wb.value("D3") is None
        assert wb.value("F1") == 1.0
        wb.set("E1", 3.0)
        refs = [first, "D3", "F1"]
        values = wb.values_at(refs) if batch else [wb.value(r) for r in refs]
        assert values == ([3.0, 3.0, 4.0] if first == "D3" else [4.0, 3.0, 4.0])
        assert wb.status("F1").state == "computed"
