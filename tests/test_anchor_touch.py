"""A cell set by hand or read from a file wakes the dynamic arrays whose range holds it (§7.7).

Only dynamic arrays spill, so only they are looked at; a file of many legacy array formulas,
each set as it is read, does not go through every anchor for each.
"""

from __future__ import annotations

import io
import random

import pytest

import xlrecalc
from tests.xlsx_files import package
from xlrecalc._engine import Engine, node


def _woken_by_any_anchor(engine: Engine, sheet: int, row: int, col: int) -> set[int]:
    """The anchors a set cell wakes, looking at every anchor: the reference."""
    n = node(sheet, row, col)
    out = set()
    for a, x in engine.anchors.items():
        if not x.dynamic or a == n or a >> 34 != sheet:
            continue
        for box in (x.box, x.wanted):
            if box is not None and box[0] <= row <= box[2] and box[1] <= col <= box[3]:
                out.add(a)
    return out


@pytest.mark.parametrize("seed", range(30))
def test_a_set_cell_wakes_the_dynamic_arrays_holding_it(seed: int) -> None:
    """Legacy and dynamic anchors on two sheets, ranges and wanted ranges overlapping, some
    anchors dropped: every cell set wakes the same anchors as looking at every anchor does."""
    rng = random.Random(seed)
    engine = Engine.from_cells({"S": {"A1": 1.0}, "T": {"A1": 1.0}})
    placed = []
    for _ in range(rng.randint(3, 25)):
        sheet, row, col = rng.randint(0, 1), rng.randint(0, 30), rng.randint(1, 8)
        box = (row, col, row + rng.randint(0, 4), col + rng.randint(0, 2))
        engine.set_formula(sheet, row, col, "A1*2", array=(rng.random() < 0.5, box))
        placed.append(node(sheet, row, col))
    for a in rng.sample(placed, len(placed) // 4):  # some replaced by a value: dropped
        if a in engine.anchors:
            sheet, row, col = a >> 34, (a & ((1 << 34) - 1)) // 16384, a % 16384
            engine.set_value(sheet, row, col, 0.0)
    for x in engine.anchors.values():  # blocked arrays want a larger range
        if x.dynamic and rng.random() < 0.4:
            r1, c1, r2, c2 = x.box
            x.wanted = (r1, c1, r2 + rng.randint(1, 5), c2 + rng.randint(0, 2))
    for _ in range(60):
        sheet, row, col = rng.randint(0, 1), rng.randint(0, 40), rng.randint(0, 12)
        expected = _woken_by_any_anchor(engine, sheet, row, col)
        engine.stale.clear()
        engine._touch_spill(sheet, row, col)
        assert engine.stale == expected


def test_reading_many_legacy_array_formulas_does_not_rebuild_the_anchor_ranges() -> None:
    """2,000 single-cell legacy array formulas: loading them, a read and an edit rebuild the
    anchors' ranges a few times at most, where loading rebuilt them for every formula read."""
    rows = "".join(
        f'<row r="{r}"><c r="A{r}"><v>{r}</v></c>'
        f'<c r="B{r}"><f t="array" ref="B{r}">A{r}*2</f><v>{2 * r}</v></c></row>'
        for r in range(1, 2001)
    )
    data = package(rows, workbook_extra='<calcPr fullCalcOnLoad="1"/>')
    count = [0]
    rebuild = Engine._rebuild_anchors

    def counted(self: Engine, sheet: int) -> None:
        count[0] += 1
        rebuild(self, sheet)

    Engine._rebuild_anchors = counted  # type: ignore[method-assign]
    try:
        wb = xlrecalc.load(io.BytesIO(data))
        assert wb.value("B2000") == 4000.0
        wb.set("A1", 5.0)
        assert wb.value("B1") == 10.0
    finally:
        Engine._rebuild_anchors = rebuild  # type: ignore[method-assign]
    assert count[0] <= 5
