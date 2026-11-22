"""``Workbook.values`` evaluates a range in one pass and gives what reading each cell gives.

Random workbooks from ``tests/test_loader_orders.py`` (references across sheets, defined names,
``INDIRECT`` and ``OFFSET``, shared and array formulas, cycles, cells marked ``ca="1"``, volatile
functions) are read twice under every calculation case and ``recalc_on_load``: by
``Workbook.values`` over each sheet's range, and cell by cell with ``Workbook.value`` in the same
order. Each workbook is also read after an edit, through a lazy ``Loader``, and with a step
allowance so small that some cells stop. Every value, every cell's status, the report and the
values after a recalculation must be the same, bit for bit and in order.
"""

from __future__ import annotations

import io
import random

import pytest
import xlio

import xlrecalc
from tests.test_loader_orders import CALCS, MODES, _book, _key
from tests.xlsx_files import DOC, MAIN, PKG, package

COLUMNS = "ABCDEFG"
WAYS = ["load", "edited", "loader", "steps"]


def _workbook(data: bytes, mode: str, way: str, book: xlio.Workbook | None) -> xlrecalc.Workbook:
    if way == "loader":
        assert book is not None
        return xlrecalc.Loader(book, recalc_on_load=mode).workbook()  # type: ignore[arg-type]
    limits = xlrecalc.ReadLimits(max_function_steps=3) if way == "steps" else None
    wb = xlrecalc.load(io.BytesIO(data), recalc_on_load=mode, limits=limits)  # type: ignore[arg-type]
    if way == "edited":
        wb.set("s1!A3", 7.0)
    return wb


def _read(wb: xlrecalc.Workbook, batch: bool) -> tuple[object, ...]:
    names = wb.sheet_names
    values: list[object] = []
    for name in names:
        if batch:
            rows = wb.values(f"{name}!A1:G5")
            values += [_key(v) for row in rows for v in row]
        else:
            values += [_key(wb.value(f"{name}!{c}{r}")) for r in range(1, 6) for c in COLUMNS]
    statuses = [wb.status(f"{n}!{c}{r}") for n in names for r in range(1, 6) for c in COLUMNS]
    report = wb.report
    reported = (
        list(report.unsupported.items()),
        list(report.refused.items()),
        list(report.exceptions.items()),
        sorted(report.uncomputed),
        sorted(report.from_kept),
        sorted(report.circular),
    )
    wb.recalculate()
    after = [_key(wb.value(f"{n}!{c}{r}")) for n in names for r in range(1, 6) for c in COLUMNS]
    return values, statuses, reported, after


def _both(data: bytes, mode: str, way: str) -> tuple[tuple[object, ...], tuple[object, ...]]:
    random.seed(0)
    with xlio.open(io.BytesIO(data)) as book:
        one = _read(_workbook(data, mode, way, book), batch=False)
    random.seed(0)
    with xlio.open(io.BytesIO(data)) as book:
        batch = _read(_workbook(data, mode, way, book), batch=True)
    return one, batch


@pytest.mark.parametrize("way", WAYS)
@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("calc", list(CALCS))
@pytest.mark.parametrize("seed", range(8))
def test_values_in_one_pass_are_the_values_read_one_by_one(
    seed: int, calc: str, mode: str, way: str
) -> None:
    data, _ = _book(random.Random(seed), CALCS[calc])
    one, batch = _both(data, mode, way)
    assert batch == one


def _read_at(wb: xlrecalc.Workbook, refs: list[str], batch: bool) -> tuple[object, ...]:
    values = wb.values_at(refs) if batch else [wb.value(r) for r in refs]
    names = wb.sheet_names
    statuses = [wb.status(f"{n}!{c}{r}") for n in names for r in range(1, 6) for c in COLUMNS]
    report = wb.report
    reported = (
        list(report.unsupported.items()),
        list(report.exceptions.items()),
        sorted(report.uncomputed),
        sorted(report.from_kept),
        sorted(report.circular),
    )
    return [_key(v) for v in values], statuses, reported


@pytest.mark.parametrize("way", WAYS)
@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("calc", list(CALCS))
@pytest.mark.parametrize("seed", range(8))
def test_values_at_are_the_values_read_one_by_one(
    seed: int, calc: str, mode: str, way: str
) -> None:
    """Cells on any sheets, in any order, some twice: the values, statuses and report that
    reading them one by one in that order gives, volatile draws included."""
    rng = random.Random(seed)
    data, every = _book(rng, CALCS[calc])
    refs = rng.choices(every, k=len(every))  # any order, repeats
    got = []
    for batch in (True, False):
        random.seed(0)
        with xlio.open(io.BytesIO(data)) as book:
            got.append(_read_at(_workbook(data, mode, way, book), refs, batch))
    assert got[0] == got[1]


@pytest.mark.parametrize(
    ("ref", "error"),
    [("s1!A1:A2", ValueError), ("nowhere!A1", KeyError), ("A1", ValueError), ("A0", ValueError)],
)
def test_values_at_raises_what_value_raises(ref: str, error: type[Exception]) -> None:
    data, _ = _book(random.Random(3), CALCS["current"])
    data2, _ = _book(random.Random(5), CALCS["current"])
    for d in (data, data2):
        wb = xlrecalc.load(io.BytesIO(d))
        if len(wb.sheet_names) == 1 and ref == "A1":
            continue  # one sheet: a reference without a sheet is that sheet's
        with pytest.raises(error):
            wb.value(ref)
        with pytest.raises(error):
            wb.values_at(["s1!A1", ref])


FULL = '<calcPr fullCalcOnLoad="1"/>'
BESSEL = "BESSELJ(2.5,100)"
"""A cell that charges about 160 steps to the allowance (``ReadLimits.max_function_steps``)."""


def _pair(data: bytes, ref: str, limits: xlrecalc.ReadLimits | None = None) -> list[object]:
    """The values of ``ref`` read in one pass and cell by cell, each from a fresh load."""
    batch = xlrecalc.load(io.BytesIO(data), limits=limits).values(ref)
    wb = xlrecalc.load(io.BytesIO(data), limits=limits)
    one = [[wb.value(c) for c in row] for row in _cells(ref)]
    return [batch, one]


def _cells(ref: str) -> list[list[str]]:
    first, last = ref.split(":")
    c1, r1, c2, r2 = first[0], int(first[1:]), last[0], int(last[1:])
    return [[f"{chr(c)}{r}" for c in range(ord(c1), ord(c2) + 1)] for r in range(r1, r2 + 1)]


def test_each_read_has_its_own_step_allowance() -> None:
    """Two cells and an array's anchor, each within the allowance, together past it: read one
    by one each is computed, and so in one pass (the anchor through a cell it fills)."""
    cells = (
        f'<row r="1"><c r="B1"><f t="array" ref="B1:B2">{BESSEL}*{{1;1}}</f><v>9</v></c></row>'
        f'<row r="2"><c r="A2"><f>{BESSEL}</f><v>9</v></c><c r="B2"><v>9</v></c></row>'
        f'<row r="3"><c r="A3"><f>{BESSEL}</f><v>9</v></c></row>'
    )
    data = package(cells, workbook_extra=FULL)
    batch, one = _pair(data, "A2:B3", xlrecalc.ReadLimits(max_function_steps=250))
    assert batch == one
    assert 9.0 not in [v for row in one for v in row if v is not None]


def test_a_cell_an_array_fills_is_read_after_its_anchor() -> None:
    """A range holding cells of an array but not its anchor: the anchor is evaluated first."""
    cells = (
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B2">A1:A2*2</f>'
        '<v>99</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>99</v></c></row>'
    )
    batch, one = _pair(package(cells, workbook_extra=FULL), "B2:B2")
    assert batch == one == [[6.0]]


def test_a_spill_that_moves_is_settled_before_the_next_read() -> None:
    """A dynamic array grows when it is evaluated: the cells reading where it now spills are
    evaluated again before the next read (C5), and a cell read before it keeps the value it had
    then (C1)."""

    def read(batch: bool) -> object:
        wb = xlrecalc.load(io.BytesIO(package("", workbook_extra=FULL)))
        wb.set("B1", 3.0)
        wb.set_formula("C1", "=A4*1")
        wb.set_formula("C5", "=A4*10")
        wb.recalculate()
        wb.set_formula("A3", "=_xlfn.SEQUENCE(B1)")
        if batch:
            return wb.values("A1:C5")
        return [[wb.value(c) for c in row] for row in _cells("A1:C5")]

    one = read(False)
    assert read(True) == one
    assert one[0][2] == 0.0 and one[4][2] == 20.0  # noqa: PT018 - the two reads' moments


def test_a_pass_leaves_the_report_due() -> None:
    """A cell computed in the pass from a value the file kept (a function not computed): its
    status, read after, is kept, which only the report's passes run again can say."""
    rows = '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c><c r="B1"><f>A1+1</f></c></row>'
    data = package(rows, workbook_extra=FULL)
    states = []
    for batch in (True, False):
        wb = xlrecalc.load(io.BytesIO(data))
        _ = wb.report  # the passes run, before anything is evaluated
        wb.values("B1:B1") if batch else wb.value("B1")
        states.append(wb.status("B1").state)
    assert states == ["kept", "kept"]


SHEETS = '<sheet name="s1" sheetId="1" r:id="rId1"/><sheet name="s2" sheetId="2" r:id="rId2"/>'


def _sheet(rows: str) -> str:
    return f'<worksheet xmlns="{MAIN}"><sheetData>{rows}</sheetData></worksheet>'


def test_a_sheet_loaded_during_the_pass() -> None:
    """A lazy ``Loader``: the second sheet is read when a formula of the first needs it."""
    sheet1 = (
        '<row r="1"><c r="A1"><f>s2!A1*2</f><v>9</v></c></row>'
        '<row r="2"><c r="A2"><f>SUM(s2!A1:B2)</f><v>9</v></c></row>'
    )
    sheet2 = (
        '<row r="1"><c r="A1"><f>B1+1</f><v>9</v></c><c r="B1"><v>4</v></c></row>'
        '<row r="2"><c r="A2"><f>A1*3</f><v>9</v></c></row>'
    )
    rels = (
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{DOC}/worksheet" Target="worksheets/sheet2.xml"/>'
    )
    data = package(
        sheets=SHEETS,
        workbook_extra=FULL,
        extra_parts={
            "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">{rels}</Relationships>',
            "xl/worksheets/sheet1.xml": _sheet(sheet1),
            "xl/worksheets/sheet2.xml": _sheet(sheet2),
        },
    )
    got = []
    for batch in (True, False):
        with xlio.open(io.BytesIO(data)) as book:
            wb = xlrecalc.Loader(book).workbook()
            got.append(
                wb.values("s1!A1:A2") if batch else [[wb.value("s1!A1")], [wb.value("s1!A2")]]
            )
    assert got[0] == got[1] == [[10.0], [24.0]]
