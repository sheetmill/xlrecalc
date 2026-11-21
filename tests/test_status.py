"""``Workbook.status``: each cell's entry is the report's, and only its precedents' sheets are read.

Random workbooks of up to five sheets hold every kind of cell the report tells apart: a function
not implemented with and without a cached value, an internal error, a cycle, a formula Excel
refuses, an array formula, references across sheets, through a defined name, through
``INDIRECT``, and in an ``IF`` branch not taken (a static precedent no evaluation reads). Each is
read lazily (some sheets fed, the rest left) and asked for the status of cells in a random order,
values and edits between; an eagerly loaded twin replays the same reads. Each status must equal
what the twin's report says of the cell at that moment, and at the end what the lazy workbook's
own report says; before any edit, no sheet is read that neither feeding nor the static precedents
of the cells read can reach.
"""

from __future__ import annotations

import io
import random
from collections.abc import Iterator

import pytest
import xlio

import xlrecalc
from tests.test_loader_orders import Reads
from tests.xlsx_files import DOC, PKG, package
from xlrecalc import CellRef, CellStatus, Report
from xlrecalc._functions import REGISTRY, function

ROWS = 4
COLUMNS = "ABC"
EXCEL_SAVED = '<fileVersion appName="xl"/><calcPr calcId="191029"/>'
STATES = ("refused", "uncomputed", "circular", "kept")


@pytest.fixture(autouse=True)
def _boom() -> Iterator[None]:
    """``_TESTBOOM``: a function whose implementation fails, so that its cell is an exception."""

    @function("_TESTBOOM", "number")
    def _fails(ctx: object, x: object) -> object:
        raise RuntimeError("boom")

    yield
    REGISTRY.pop("_TESTBOOM")


def _formula(rng: random.Random, sheets: int, row: int) -> tuple[str, set[str]]:
    """A formula, and the sheets it names (its static precedents' sheets)."""
    other = f"s{rng.randint(1, sheets)}"
    r = rng.randint(1, ROWS)
    pool = [
        (f"A{r}*2", set()),
        (f"{other}!A{r}+1", {other}),
        (f"SUM({other}!A1:B{ROWS})", {other}),
        (f"IF(FALSE,{other}!B{r},1)", {other}),  # a static precedent no evaluation reads
        (f'INDIRECT("{other}!A{r}")', {other}),
        (f"B{r}+1", set()),  # a chain, or a cycle within the sheet
        (f"{other}!B{r}+C{row}", {other}),  # a chain, or a cycle across sheets
        ("nm+1", {"s1"}),
        (f'CUBEVALUE("c",{other}!A{r})', {other}),  # not implemented: the cached value is kept
        (f"_TESTBOOM({other}!A{r})", {other}),  # an internal error
        ("AREAS(1)", set()),  # Excel refuses it in a file
    ]
    if rng.random() < 0.6:  # most formulas stay on their sheet, so that sheets stay unread
        pool = [
            (f"A{r}*2", set()),
            (f"B{r}+1", set()),
            (f'CUBEVALUE("c",A{r})', set()),
            (f"_TESTBOOM(A{r})", set()),
            ("AREAS(1)", set()),
        ]
    return rng.choice(pool)


def _sheet(rng: random.Random, sheets: int, excel: bool) -> tuple[str, set[str], list[str]]:
    """A sheet's XML, the sheets its formulas name, and the cells an array formula fills."""
    rows: dict[int, list[str]] = {r: [] for r in range(1, ROWS + 1)}
    named: set[str] = set()
    for r in range(1, ROWS + 1):
        for c in COLUMNS:
            if rng.random() < 0.4:
                rows[r].append(f'<c r="{c}{r}"><v>{rng.randint(1, 9)}</v></c>')
                continue
            text, names = _formula(rng, sheets, r)
            named |= names
            cached = "<v>7</v>" if excel or rng.random() < 0.6 else ""
            body = text.replace("&", "&amp;").replace('"', "&quot;")
            rows[r].append(f'<c r="{c}{r}"><f>{body}</f>{cached}</c>')
    members: list[str] = []
    if rng.random() < 0.5:  # an array formula over D1:D2 that keeps its cached values
        text, names = rng.choice([('CUBEVALUE("c")', set()), ("A1:A2*2", set())])
        named |= names
        body = text.replace('"', "&quot;")
        rows[1].append(f'<c r="D1"><f t="array" ref="D1:D2">{body}</f><v>5</v></c>')
        rows[2].append('<c r="D2"><v>6</v></c>')
        members.append("D2")
    data = "".join(f'<row r="{r}">{"".join(cells)}</row>' for r, cells in rows.items())
    return f'<worksheet xmlns="{_MAIN}"><sheetData>{data}</sheetData></worksheet>', named, members


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _book(rng: random.Random) -> tuple[bytes, list[str], dict[str, set[str]], dict[str, str]]:
    """A workbook, the cells worth asking, each sheet's named sheets, and array members' anchors."""
    sheets = rng.randint(1, 5)
    excel = rng.random() < 0.3  # Excel's saved form: every cached value trusted
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{DOC}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheets + 1)
    )
    parts: dict[str, str | bytes] = {
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">{rels}</Relationships>'
    }
    named: dict[str, set[str]] = {}
    anchors: dict[str, str] = {}
    for i in range(1, sheets + 1):
        xml, names, members = _sheet(rng, sheets, excel)
        parts[f"xl/worksheets/sheet{i}.xml"] = xml
        named[f"s{i}"] = names
        anchors.update({f"s{i}!{m}": f"s{i}!D1" for m in members})
    data = package(
        sheets="".join(
            f'<sheet name="s{i}" sheetId="{i}" r:id="rId{i}"/>' for i in range(1, sheets + 1)
        ),
        workbook_extra=(EXCEL_SAVED if excel else "")
        + '<definedNames><definedName name="nm">s1!$A$1</definedName></definedNames>',
        extra_parts=parts,
    )
    refs = [
        f"s{i}!{c}{r}" for i in range(1, sheets + 1) for c in "ABCD" for r in range(1, ROWS + 1)
    ]
    return data, refs, named, anchors


def expected(report: Report, where: CellRef, formula: bool) -> str:
    """A cell's state from the report's sets alone, in ``CellStatus``'s order."""
    if where in report.refused:
        return "refused"
    if where in report.uncomputed:
        return "uncomputed"
    if not formula:
        return "constant"
    if where in report.circular:
        return "circular"
    if where in report.from_kept:
        return "kept"
    return "computed"


def _ref(text: str) -> CellRef:
    sheet, cell = text.split("!")
    return CellRef(sheet, int(cell[1:]), ord(cell[0]) - 64)


def _check_status(
    status: CellStatus, report: Report, ref: str, wb: xlrecalc.Workbook, anchors: dict[str, str]
) -> None:
    """The status agrees with the report on the cell (or on the array formula filling it)."""
    owner = anchors.get(ref, ref)
    where = _ref(owner)
    formula = wb.formula(owner) is not None or where in report.refused
    assert status.state == expected(report, where, formula), (ref, status)
    assert status.anchor == (where if owner != ref else None), (ref, status)
    assert status.error == report.exceptions.get(where), (ref, status)
    assert status.refusal == report.refused.get(where), (ref, status)
    text = (wb.formula(owner) or "").upper()
    for label in status.labels:  # the cell's own formula's: counted or not (an edit holds it)
        assert label in report.unsupported or label.split()[0] + "(" in text, (ref, status)
    if status.inherited:
        assert status.state in ("uncomputed", "kept")
        assert not status.labels
        assert status.error is None
    elif status.state in ("uncomputed", "kept"):
        assert status.labels or status.error, (ref, status)


def _reach(named: dict[str, set[str]], start: set[str]) -> set[str]:
    """The sheets reachable from ``start`` through the sheets each sheet's formulas name."""
    seen, todo = set(start), list(start)
    while todo:
        for name in named.get(todo.pop(), ()):
            if name not in seen:
                seen.add(name)
                todo.append(name)
    return seen


TOLD: dict[str, int] = {"left unread": 0, "statuses": 0}


def _check(seed: int, monkeypatch: pytest.MonkeyPatch) -> None:
    rng = random.Random(seed)
    data, refs, named, anchors = _book(rng)
    twin = xlrecalc.load(io.BytesIO(data))
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        sheets = [s for s in book.sheets if s.kind == "worksheet"]
        loader = xlrecalc.Loader(book)
        fed = {s.name for s in sheets if rng.random() < 0.4}
        for sheet in sheets:
            if sheet.name in fed:
                for _ in loader.cells(sheet):
                    pass
        wb = loader.workbook()
        asked: list[str] = []
        touched: set[str] = set(fed)
        edited = False
        for _ in range(rng.randint(1, 12)):
            ref = rng.choice(refs)
            roll = rng.random()
            if roll < 0.1 and ref[-2] != "D":  # an edit outside the array, in both
                value = float(rng.randint(1, 9))
                wb.set(ref, value)
                twin.set(ref, value)
                edited = True
                continue
            touched.add(ref.split("!")[0])
            if roll < 0.3:  # a value read, in both
                assert repr(wb.value(ref)) == repr(twin.value(ref))
                continue
            status = wb.status(ref)
            twin.value(ref)
            _check_status(status, twin.report, ref, twin, anchors)
            asked.append(ref)
            TOLD["statuses"] += 1
            if not edited:
                read = {name for name, calls in reads.calls.items() if calls}
                assert read <= _reach(named, touched), (seed, read, touched)
                if len(read) < len(sheets):
                    TOLD["left unread"] += 1
        report = wb.report
        for ref in asked:
            _check_status(wb.status(ref), report, ref, wb, anchors)


@pytest.mark.parametrize("seed", range(150))
def test_status_is_the_reports_entry(seed: int, monkeypatch: pytest.MonkeyPatch) -> None:
    _check(seed, monkeypatch)


@pytest.mark.slow
@pytest.mark.parametrize("seed", range(150, 3150))
def test_status_is_the_reports_entry_many(seed: int, monkeypatch: pytest.MonkeyPatch) -> None:
    _check(seed, monkeypatch)


def test_the_property_test_exercises_laziness(monkeypatch: pytest.MonkeyPatch) -> None:
    """Its cases leave sheets unread and ask many statuses: a setup reading every sheet first
    would hide what it tests."""
    TOLD.update({"left unread": 0, "statuses": 0})
    for seed in range(150):
        with monkeypatch.context() as m:
            _check(seed, m)
    assert TOLD["statuses"] > 500, TOLD
    assert TOLD["left unread"] > 50, TOLD


def _statuses(first: str, second: str, refs: list[str]) -> dict[str, CellStatus]:
    """Each cell's status, from a ``Loader`` with sheet "one" fed and "two" not."""
    from tests.test_loader import _lazy_book

    with xlio.open(io.BytesIO(_lazy_book(first, second))) as book:
        loader = xlrecalc.Loader(book)
        for _ in loader.cells(book.sheets[0]):
            pass
        wb = loader.workbook()
        return {ref: wb.status(ref) for ref in refs}


def test_each_state_and_its_reason() -> None:
    one = (
        '<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>A1*2</f></c>'
        '<c r="C1"><f>AREAS(1)</f><v>4</v></c><c r="D1"><f>CUBEVALUE("c")</f></c>'
        '<c r="E1"><f>D1+1</f></c><c r="F1"><f>CUBEVALUE("c")</f><v>5</v></c>'
        '<c r="G1"><f>F1+1</f><v>0</v></c><c r="H1"><f>H1+1</f><v>2</v></c>'
        '<c r="I1"><f>_TESTBOOM(1)</f></c></row>'
    )
    got = _statuses(one, "", [f"one!{c}1" for c in "ABCDEFGHI"])
    assert got == {
        "one!A1": CellStatus("constant"),
        "one!B1": CellStatus("computed"),
        "one!C1": CellStatus("refused", refusal="an argument of AREAS that is no reference"),
        "one!D1": CellStatus("uncomputed", ("CUBEVALUE",)),
        "one!E1": CellStatus("uncomputed", inherited=True),
        "one!F1": CellStatus("kept", ("CUBEVALUE",)),
        "one!G1": CellStatus("kept", inherited=True),
        "one!H1": CellStatus("circular"),
        "one!I1": CellStatus("uncomputed", error="RuntimeError: boom"),
    }


def test_a_cell_an_array_fills_answers_with_its_formula() -> None:
    one = (
        '<row r="1"><c r="A1"><f t="array" ref="A1:A2">CUBEVALUE("c")</f><v>5</v></c></row>'
        '<row r="2"><c r="A2"><v>6</v></c></row>'
    )
    anchor = CellRef("one", 1, 1)
    got = _statuses(one, "", ["one!A2"])
    assert got["one!A2"] == CellStatus("kept", ("CUBEVALUE",), anchor=anchor)


def test_a_limit_passed_is_the_cells_label() -> None:
    data = package(
        '<row r="1"><c r="A1"><f t="array" ref="A1">SUM(B1:B100*1)</f></c></row>',
        workbook_extra='<calcPr fullCalcOnLoad="1"/>',
    )
    wb = xlrecalc.load(io.BytesIO(data), limits=xlrecalc.ReadLimits(max_array_cells=10))
    assert wb.status("Data!A1") == CellStatus("uncomputed", ("an array of more than 10 cells",))


def test_only_the_sheets_precedents_name_are_read(monkeypatch: pytest.MonkeyPatch) -> None:
    """A cell naming no other sheet reads its own; a sheet named only in an ``IF`` branch not
    taken is read (a static precedent); one no formula names is not."""
    data = _book_of(
        {
            "a": '<row r="1"><c r="A1"><f>IF(FALSE,b!A1,1)</f></c><c r="B1"><f>2*3</f></c></row>',
            "b": '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c></row>',
            "c": '<row r="1"><c r="A1"><v>1</v></c></row>',
        }
    )
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        wb = xlrecalc.Loader(book).workbook()
        assert wb.status("a!B1") == CellStatus("computed")
        assert set(reads.calls) == {"a"}
        assert wb.status("a!A1") == CellStatus("kept", inherited=True)
        assert set(reads.calls) == {"a", "b"}
        assert CellRef("a", 1, 1) in wb.report.from_kept
        assert set(reads.calls) == {"a", "b", "c"}


def _book_of(sheets: dict[str, str], defined: str = "") -> bytes:
    names = list(sheets)
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{DOC}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(names) + 1)
    )
    parts: dict[str, str | bytes] = {
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">{rels}</Relationships>'
    }
    for i, name in enumerate(names, 1):
        xml = f'<worksheet xmlns="{_MAIN}"><sheetData>{sheets[name]}</sheetData></worksheet>'
        parts[f"xl/worksheets/sheet{i}.xml"] = xml
    return package(
        sheets="".join(
            f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(names, 1)
        ),
        workbook_extra=defined + '<calcPr fullCalcOnLoad="1"/>',
        extra_parts=parts,
    )


def test_values_first_then_statuses_run_the_reports_passes_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from xlrecalc._engine import Engine

    rows = "".join(
        f'<row r="{r}"><c r="A{r}"><f>CUBEVALUE("c")</f><v>{r}</v></c>'
        f'<c r="B{r}"><f>A{r}+1</f></c></row>'
        for r in range(1, 51)
    )
    wb = xlrecalc.load(io.BytesIO(package(rows, workbook_extra='<calcPr fullCalcOnLoad="1"/>')))
    runs = []
    original = Engine._report_passes

    def counted(engine: Engine) -> None:
        runs.append(1)
        original(engine)

    monkeypatch.setattr(Engine, "_report_passes", counted)
    wb.values("Data!A1:B50")
    states = {wb.status(f"Data!B{r}").state for r in range(1, 51)}
    assert states == {"kept"}
    assert len(runs) == 1


def test_a_defined_name_s_sheets_are_read(monkeypatch: pytest.MonkeyPatch) -> None:
    """A name in a branch not taken is a static precedent too: its sheet is read."""
    data = _book_of(
        {
            "a": '<row r="1"><c r="A1"><f>IF(FALSE,far,1)</f></c></row>',
            "b": '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f><v>4</v></c></row>',
        },
        '<definedNames><definedName name="far">b!$A$1</definedName></definedNames>',
    )
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        wb = xlrecalc.Loader(book).workbook()
        assert wb.status("a!A1") == CellStatus("kept", inherited=True)
        assert set(reads.calls) == {"a", "b"}
