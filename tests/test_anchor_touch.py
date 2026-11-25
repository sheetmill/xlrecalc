"""A cell set by hand or read from a file wakes the dynamic arrays whose range holds it (§7.7).

Only dynamic arrays spill, so only they are looked at; a file of many legacy array formulas,
each set as it is read, does not go through every anchor for each. The index of the ranges
anchors fill is kept up to date anchor by anchor, so arrays whose ranges change as they are
evaluated do not go through every anchor for each either.
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


def _assert_index_is_rebuilt_as_is(engine: Engine) -> None:
    """The anchor index kept up to date equals one built from every anchor: the reference."""
    boxes: dict[int, dict[int, tuple[tuple[int, ...], ...]]] = {}
    columns: dict[int, dict[int, list[tuple[int, int, int]]]] = {}
    for a, x in engine.anchors.items():
        sheet = a >> 34
        entry = ((*x.box, a),) if x.wanted is None else ((*x.box, a), (*x.wanted, a))
        boxes.setdefault(sheet, {})[a] = entry
        for c in range(x.box[1], x.box[3] + 1):
            columns.setdefault(sheet, {}).setdefault(c, []).append((x.box[0], x.box[2], a))
    kept = {sheet: d for sheet, d in engine._anchor_boxes.items() if d}
    assert kept == boxes
    assert {sheet: list(d) for sheet, d in kept.items()} == {  # the order of each sheet's
        sheet: list(d) for sheet, d in boxes.items()
    }
    for sheet in set(engine._member_spans) | set(columns):
        spans = {c: v for c, v in engine._member_spans.get(sheet, {}).items() if v[1]}
        expected = {c: sorted(v) for c, v in columns.get(sheet, {}).items()}
        assert {c: found for c, (_, found) in spans.items()} == expected
        assert all(starts == [f[0] for f in found] for starts, found in spans.values())


SIZES = 30  # column AE: the cells holding each array's height and width


@pytest.mark.parametrize("seed", range(40))
def test_a_set_cell_wakes_the_dynamic_arrays_holding_it(seed: int) -> None:
    """Legacy arrays and dynamic arrays (``SEQUENCE`` of a height and a width held in cells) on
    two sheets, some wider than the index's column limit; their sizes edited and constants put in
    their way and taken out, so that spills grow, shrink, are blocked and free again, through
    evaluation; some arrays replaced by values. After each round, every cell set wakes the same
    arrays as looking at every array does."""
    rng = random.Random(seed)
    engine = Engine.from_cells({"S": {"A1": 1.0}, "T": {"A1": 1.0}})
    placed = []
    for k in range(rng.randint(3, 20)):
        sheet, row, col = rng.randint(0, 1), rng.randint(0, 30), rng.randint(1, 8)
        if rng.random() < 0.3:
            box = (row, col, row + rng.randint(0, 3), col + rng.randint(0, 2))
            engine.set_formula(sheet, row, col, "A1*2", array=(False, box))
        else:
            wide = rng.random() < 0.15
            engine.set_value(sheet, k, SIZES, float(rng.randint(1, 4)))
            engine.set_value(sheet, k, SIZES + 1, float(rng.randint(70, 80) if wide else 2))
            size = f"_xlfn.SEQUENCE(AE{k + 1},AF{k + 1})"
            engine.set_formula(sheet, row, col, size, array=(True, (row, col, row, col)))
        placed.append(node(sheet, row, col))
    for _ in range(4):
        engine.recalculate()
        _assert_index_is_rebuilt_as_is(engine)
        for _ in range(60):
            sheet, row, col = rng.randint(0, 1), rng.randint(0, 40), rng.randint(0, 90)
            expected = _woken_by_any_anchor(engine, sheet, row, col)
            stale, unclosed = set(engine.stale), list(engine._unclosed)
            engine.stale.clear()
            engine._touch_spill(sheet, row, col)
            assert engine.stale == expected, (sheet, row, col)
            engine.stale, engine._unclosed = stale, unclosed  # the round's state, untouched
        for _ in range(rng.randint(1, 6)):  # sizes edited, constants put in the way and taken out
            sheet = rng.randint(0, 1)
            if rng.random() < 0.5:
                engine.set_value(sheet, rng.randint(0, 19), SIZES, float(rng.randint(1, 6)))
            else:
                row, col = rng.randint(0, 35), rng.randint(1, 12)
                if node(sheet, row, col) not in engine.anchors:
                    value = 7.0 if rng.random() < 0.6 else None
                    engine.set_value(sheet, row, col, value)
        if placed and rng.random() < 0.3:  # an array replaced by a value
            a = placed.pop(rng.randrange(len(placed)))
            engine.set_value(a >> 34, (a & ((1 << 34) - 1)) // 16384, a % 16384, 0.0)


class _CountedAnchors(dict):  # type: ignore[type-arg]
    """``Engine.anchors``, counting the walks over every anchor."""

    walks = 0

    def __iter__(self):  # type: ignore[no-untyped-def]
        _CountedAnchors.walks += 1
        return super().__iter__()

    def items(self):  # type: ignore[no-untyped-def]
        _CountedAnchors.walks += 1
        return super().items()

    def values(self):  # type: ignore[no-untyped-def]
        _CountedAnchors.walks += 1
        return super().values()

    # keys() is not counted: Engine.dependents intersects it with a set, which walks the set.


def _counting_anchors() -> object:
    """Make every new engine count the walks over its anchors; returns the original __init__."""
    init = Engine.__init__

    def counted(self: Engine, *args: object, **kwargs: object) -> None:
        init(self, *args, **kwargs)  # type: ignore[arg-type]
        self.anchors = _CountedAnchors()

    _CountedAnchors.walks = 0
    Engine.__init__ = counted  # type: ignore[method-assign]
    return init


def test_reading_many_legacy_array_formulas_does_not_go_through_every_anchor() -> None:
    """2,000 single-cell legacy array formulas: loading them, a read and an edit walk the
    anchors a few times at most, where loading rebuilt their ranges for every formula read."""
    rows = "".join(
        f'<row r="{r}"><c r="A{r}"><v>{r}</v></c>'
        f'<c r="B{r}"><f t="array" ref="B{r}">A{r}*2</f><v>{2 * r}</v></c></row>'
        for r in range(1, 2001)
    )
    data = package(rows, workbook_extra='<calcPr fullCalcOnLoad="1"/>')
    init = _counting_anchors()
    try:
        wb = xlrecalc.load(io.BytesIO(data))
        assert wb.value("B2000") == 4000.0
        wb.set("A1", 5.0)
        assert wb.value("B1") == 10.0
    finally:
        Engine.__init__ = init  # type: ignore[method-assign]
    assert len(wb._engine.anchors) == 2000
    assert _CountedAnchors.walks <= 5


def test_dynamic_arrays_whose_ranges_change_do_not_go_through_every_anchor() -> None:
    """2,000 dynamic arrays set with their anchor cell as their range, spilling over two rows
    when evaluated, then all resized by an edit: each range change updates the index for its
    own anchor, where each rebuilt it from every anchor on the sheet."""
    init = _counting_anchors()
    try:
        engine = Engine.from_cells({"S": {"C1": 2.0}})
        for k in range(2000):
            engine.set_formula(
                0, 3 * k, 0, "_xlfn.SEQUENCE($C$1)", array=(True, (3 * k, 0, 3 * k, 0))
            )
        engine.recalculate()
        assert engine.sheets[0].values[(3 * 1999 + 1) * 16384] == 2.0
        engine.set_value(0, 0, 2, 3.0)  # three rows: each array now blocked by the next
        engine.recalculate()
        engine.set_value(0, 0, 2, 1.0)
        engine.recalculate()
    finally:
        Engine.__init__ = init  # type: ignore[method-assign]
    _assert_index_is_rebuilt_as_is(engine)
    assert _CountedAnchors.walks <= 5


class _CountedWalks(dict):  # type: ignore[type-arg]
    """A dict counting the walks over it."""

    walks = 0

    def __iter__(self):  # type: ignore[no-untyped-def]
        _CountedWalks.walks += 1
        return super().__iter__()


def test_reading_many_dynamic_arrays_does_not_go_through_all_of_them_for_each() -> None:
    """2,000 dynamic arrays (``SEQUENCE(2)`` every third row): loading them, a read and an edit
    walk the list of dynamic arrays a few times at most, where loading went through it for every
    formula read."""
    from pathlib import Path

    from tests.xlsx_files import DOC, PKG

    metadata = (Path(__file__).parent / "captures" / "arrays_metadata.xml").read_text("utf-8")
    rows = "".join(
        f'<row r="{3 * k + 1}"><c r="A{3 * k + 1}" cm="1">'
        f'<f t="array" ref="A{3 * k + 1}:A{3 * k + 2}">_xlfn.SEQUENCE(2)</f><v>1</v></c></row>'
        f'<row r="{3 * k + 2}"><c r="A{3 * k + 2}"><v>2</v></c></row>'
        for k in range(2000)
    )
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>"
    )
    data = package(
        rows,
        workbook_extra='<calcPr fullCalcOnLoad="1"/>',
        extra_parts={"xl/metadata.xml": metadata, "xl/_rels/workbook.xml.rels": rels},
    )
    init = Engine.__init__

    def counted(self: Engine, *args: object, **kwargs: object) -> None:
        init(self, *args, **kwargs)  # type: ignore[arg-type]
        self._dynamic = _CountedWalks()

    _CountedWalks.walks = 0
    Engine.__init__ = counted  # type: ignore[method-assign]
    try:
        wb = xlrecalc.load(io.BytesIO(data))
        assert wb.value("A5999") == 2.0
        wb.set("C1", 5.0)
        assert wb.value("A2") == 2.0
    finally:
        Engine.__init__ = init  # type: ignore[method-assign]
    assert len(wb._engine._dynamic) == 2000
    assert _CountedWalks.walks <= 5
