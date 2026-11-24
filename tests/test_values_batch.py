"""``Workbook.values`` evaluates a range in one pass and gives what reading each cell gives.

Random workbooks from ``tests/test_loader_orders.py`` (references across sheets, defined names,
``INDIRECT`` and ``OFFSET``, shared and array formulas, cycles, cells marked ``ca="1"``, volatile
functions) are read twice under every calculation case and ``recalc_on_load``: by
``Workbook.values`` over each sheet's range, and cell by cell with ``Workbook.value`` in the same
order. Each workbook is also read after an edit, through a lazy ``Loader``, and with a step
allowance so small that some cells stop. Every value, every cell's status, the report and the
values after a recalculation must be the same, bit for bit and in order, and reading the
statuses after the values must run the report's passes as many times either way.
"""

from __future__ import annotations

import contextlib
import io
import random
from collections.abc import Iterator
from pathlib import Path

import pytest
import xlio

import xlrecalc
from tests.test_loader_orders import CALCS, MODES, _book, _key
from tests.xlsx_files import DOC, MAIN, PKG, package
from xlrecalc._engine import COLS, CellRef, CellStatus, Engine, node, unpack

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


@contextlib.contextmanager
def _passes() -> Iterator[list[int]]:
    """Count the report's whole-workbook passes run inside the block (their cost, not a time)."""
    count = [0]
    run = Engine._report_passes

    def counted(self: Engine) -> None:
        count[0] += 1
        run(self)

    Engine._report_passes = counted  # type: ignore[method-assign]
    try:
        yield count
    finally:
        Engine._report_passes = run  # type: ignore[method-assign]


def _read(wb: xlrecalc.Workbook, batch: bool) -> tuple[object, ...]:
    names = wb.sheet_names
    values: list[object] = []
    for name in names:
        if batch:
            rows = wb.values(f"{name}!A1:G5")
            values += [_key(v) for row in rows for v in row]
        else:
            values += [_key(wb.value(f"{name}!{c}{r}")) for r in range(1, 6) for c in COLUMNS]
    with _passes() as passes:
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
    return values, statuses, passes[0], reported, after


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
    with _passes() as passes:
        statuses = [wb.status(r) for r in refs]
        statuses += [wb.status(f"{n}!{c}{r}") for n in names for r in range(1, 6) for c in COLUMNS]
    report = wb.report
    reported = (
        list(report.unsupported.items()),
        list(report.exceptions.items()),
        sorted(report.uncomputed),
        sorted(report.from_kept),
        sorted(report.circular),
    )
    return [_key(v) for v in values], statuses, passes[0], reported


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


@contextlib.contextmanager
def _checked_live() -> Iterator[None]:
    """Check every use of the references kept per (template, sheet, column) from one pass to
    the next against those found afresh."""
    stale_precedents = Engine._stale_precedents

    def checked(self: Engine, n: int, *args: object) -> Iterator[int]:
        sheet, row, col = unpack(n)
        tid = self.sheets[sheet].templates.at(row, col)
        kept = self._live.get((tid, sheet, col)) if tid is not None else None
        if kept is not None:
            assert kept == self._live_specs(tid, sheet, col)  # type: ignore[arg-type]
        return stale_precedents(self, n, *args)  # type: ignore[arg-type]

    Engine._stale_precedents = checked  # type: ignore[assignment,method-assign]
    try:
        yield
    finally:
        Engine._stale_precedents = stale_precedents  # type: ignore[method-assign]


@pytest.mark.parametrize("way", WAYS)
@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("calc", list(CALCS))
@pytest.mark.parametrize("seed", range(4))
def test_the_references_kept_between_passes_are_current(
    seed: int, calc: str, mode: str, way: str
) -> None:
    """Reading cells one by one runs a pass for each; the references a template can reach a
    formula through are kept from one to the next. Through reads, statuses, an edit, a formula
    in a column that held none, and a recalculation, they are the ones found afresh."""
    rng = random.Random(seed)
    data, every = _book(rng, CALCS[calc])
    half = len(every) // 2
    random.seed(0)
    with _checked_live(), xlio.open(io.BytesIO(data)) as book:
        wb = _workbook(data, mode, way, book)
        for ref in every[:half]:
            wb.value(ref)
        _ = [wb.status(r) for r in every[:half]]
        wb.set("s1!A2", 3.0)
        wb.set_formula("s1!K5", "=A2*2")
        for ref in every[half:]:
            wb.value(ref)
        wb.recalculate()
        _ = [wb.status(r) for r in every]


def test_a_formula_in_a_column_that_held_none_is_a_precedent() -> None:
    """B1 reads C1:C3, which hold no formula, and its reference is dropped; once C2 holds a
    formula, the reference is kept again and C2 is evaluated before B1."""
    rows = (
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(C1:C3)</f><v>0</v></c></row>'
        '<row r="2"><c r="C2"><v>5</v></c></row>'
    )
    wb = xlrecalc.load(io.BytesIO(package(rows, workbook_extra=FULL)))
    with _checked_live():
        assert wb.value("B1") == 5.0
        engine = wb._engine
        key = (engine.sheets[0].templates.at(0, 1), 0, 1)  # B1's template, sheet, column
        assert engine._live[key] == ()  # kept after the pass: no reference reaches a formula
        wb.set_formula("C2", "=A1*10")
        assert key not in engine._live  # dropped with the change
        assert wb.value("B1") == 10.0
        assert len(engine._live[key]) == 1  # SUM's reference now reaches one


def _classify_in_full(engine: Engine, n: int, filled: bool) -> CellStatus:
    """A cell's status worked out in full from the report, as ``Engine._classify`` did before
    it answered a report naming no cell at once: the reference the shortcut must agree with."""
    sheet, row, col = unpack(n)
    s = engine.sheets[sheet]
    where = CellRef(s.name, row + 1, col + 1)
    anchor = where if filled else None
    report = engine.report
    tid = s.templates.get(row * COLS + col)
    labels = set(engine._stopped.get(n, ()))
    if tid is not None:
        labels.update(engine.templates[tid].unsupported)
    own = tuple(sorted(labels))
    refusal = report.refused.get(where)
    if refusal is not None:
        return CellStatus("refused", refusal=refusal, anchor=anchor)
    error = report.exceptions.get(where)
    if where in report.uncomputed:
        inherited = not own and error is None
        return CellStatus("uncomputed", own, error, inherited=inherited, anchor=anchor)
    if tid is None:
        if where in report.from_kept:
            return CellStatus("kept", own, anchor=anchor)
        return CellStatus("constant")
    if where in report.circular:
        return CellStatus("circular", anchor=anchor)
    if where in report.from_kept:
        kept = n in engine._absent
        return CellStatus("kept", own if kept else (), inherited=not kept, anchor=anchor)
    return CellStatus("computed", anchor=anchor)


@pytest.mark.parametrize("way", WAYS)
@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("calc", list(CALCS))
@pytest.mark.parametrize("seed", range(4))
def test_each_status_is_the_one_worked_out_in_full(
    seed: int, calc: str, mode: str, way: str
) -> None:
    """Every cell's status, the formula cells, the cells arrays fill and some constants, after
    reading and after an edit, is the one worked out in full from the report."""
    rng = random.Random(seed)
    data, every = _book(rng, CALCS[calc])
    random.seed(0)
    with xlio.open(io.BytesIO(data)) as book:
        wb = _workbook(data, mode, way, book)
        engine = wb._engine
        for step in range(2):
            for ref in [*every, "s1!A1", "s1!A2", "s1!Z9"]:
                got = wb.status(ref)
                sheet, row, col = wb._cell(ref)
                n = node(sheet, row, col)
                anchor = engine.member_of(n)
                filled = anchor is not None
                assert got == _classify_in_full(engine, anchor if filled else n, filled), ref
            if step == 0:
                wb.set("s1!A2", 3.0)


def _kept_book(rows: str, extra: str = "", names: str = "") -> xlrecalc.Workbook:
    """A one-sheet workbook, with the metadata Excel saves for a dynamic array (``cm="1"``)."""
    metadata = (Path(__file__).parent / "captures" / "arrays_metadata.xml").read_text("utf-8")
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>"
    )
    parts = {"xl/metadata.xml": metadata, "xl/_rels/workbook.xml.rels": rels}
    return xlrecalc.load(io.BytesIO(package(rows, workbook_extra=extra + names, extra_parts=parts)))


def test_what_rests_on_a_kept_value_follows_a_reference_made_later() -> None:
    """A reference made in evaluation after the passes ran: the cell reading the kept value
    through it is kept too (the report follows the reads recorded since)."""
    rows = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c>'
        '<c r="B1"><f>INDIRECT("A1")+1</f><v>0</v></c><c r="C1"><f>A1*2</f><v>0</v></c></row>'
    )
    wb = _kept_book(rows, FULL)
    assert wb.status("C1").state == "kept"
    wb.value("B1")
    assert wb.status("B1").state == "kept"


def test_what_rests_on_a_kept_value_follows_the_kept_cells() -> None:
    """An edit puts a kept value out of date: the kept cells change, the formulas do not."""
    rows = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")+D1</f><v>4</v></c>'
        '<c r="B1"><f>CUBEVALUE("d")</f><v>3</v></c>'
        '<c r="C1"><f>A1*2</f><v>0</v></c><c r="D1"><v>1</v></c>'
        '<c r="E1"><f>B1*2</f><v>0</v></c></row>'
    )
    wb = _kept_book(rows, FULL)
    assert wb.status("C1").state == "kept"
    wb.set("D1", 2.0)
    assert wb.status("C1").state == "uncomputed"
    assert wb.status("E1").state == "kept"


def test_what_rests_on_a_kept_value_follows_a_spill() -> None:
    """A spill grows over a cell another formula reads: that formula now rests on the kept
    value the spill is computed from."""
    rows = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c>'
        '<c r="D1" cm="1"><f t="array" ref="D1:D2">_xlfn.SEQUENCE(E1)*A1</f></c>'
        '<c r="E1"><v>2</v></c><c r="F1"><f>D3+1</f><v>0</v></c></row>'
    )
    wb = _kept_book(rows, FULL)
    assert wb.status("F1").state == "computed"
    wb.set("E1", 3.0)
    assert wb.value("D1") == 4.0  # the anchor, evaluated again: it now fills D1:D3
    assert wb.value("D3") == 12.0
    assert wb.status("F1").state == "kept"


def test_what_rests_on_a_kept_value_follows_a_volatile_name_woken() -> None:
    """A cell volatile only through a name keeps its cached value on opening (dormant, its
    name's edges left out); computed by a full recalculation, it rests on the kept value."""
    rows = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c>'
        '<c r="B1"><f>vk*2</f><v>8</v></c></row>'
    )
    names = '<definedNames><definedName name="vk">Data!$A$1+NOW()*0</definedName></definedNames>'
    wb = _kept_book(rows, '<fileVersion appName="xl"/><calcPr calcId="191029"/>', names)
    assert wb._engine.dormant
    assert wb.status("B1").state == "computed"  # its edge from the name left out
    wb.recalculate(full=True)
    assert not wb._engine.dormant
    assert wb.status("B1").state == "kept"


def test_what_rests_on_a_kept_value_follows_a_formula_set() -> None:
    """A formula set after the passes ran, reading a kept value: it is kept too."""
    rows = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c>'
        '<c r="B1"><f>A1+1</f><v>0</v></c></row>'
    )
    wb = _kept_book(rows, FULL)
    assert wb.status("B1").state == "kept"
    wb.set_formula("C1", "=A1*2")
    assert wb.status("C1").state == "kept"


def test_what_rests_on_a_kept_value_follows_a_sheet_loaded_later() -> None:
    """A lazy ``Loader`` reads a second sheet after the passes ran: its formula reading the
    kept value is kept too."""
    sheet1 = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c>'
        '<c r="B1"><f>A1+1</f><v>0</v></c></row>'
    )
    sheet2 = '<row r="1"><c r="A1"><f>s1!A1*2</f><v>0</v></c></row>'
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
    with xlio.open(io.BytesIO(data)) as book:
        wb = xlrecalc.Loader(book).workbook()
        assert wb.status("s1!B1").state == "kept"
        assert wb.status("s2!A1").state == "kept"


def test_what_rests_on_a_kept_value_drops_a_reference_no_longer_made() -> None:
    """A formula that read a kept value through ``INDIRECT`` and, after an edit, no longer
    does: it is computed, no longer kept."""
    rows = (
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c>'
        '<c r="B1"><f>IF(D1=1,INDIRECT("A1"),5)</f><v>0</v></c><c r="D1"><v>1</v></c></row>'
    )
    wb = _kept_book(rows, FULL)
    assert wb.status("B1").state == "kept"
    wb.set("D1", 0.0)
    assert wb.status("B1").state == "computed"


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
    """A dynamic array grows when it is evaluated: the cells reading where it now spills see the
    spilled value, one read before the array's own cell (C1) as much as one read after (C5),
    since a read evaluates the stale dynamic arrays first, as Excel recalculates them."""

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
    assert one[0][2] == 2.0
    assert one[4][2] == 20.0


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
            refs = ["s1!A1", "s1!A2"]
            values = wb.values_at(refs) if batch else [wb.value(r) for r in refs]
            with _passes() as passes:
                statuses = [wb.status(r).state for r in refs]
            got.append((values, statuses, passes[0]))
    assert got[0] == got[1]
    assert got[0][0] == [10.0, 24.0]
