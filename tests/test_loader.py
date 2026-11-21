"""``Loader``: a workbook read from the cells a caller streams through ``xlio``, read once.

Fed in full it gives exactly ``load``'s workbook; a sheet not fed, or left part-way, is read again
by ``finish``; misuse raises; and its opening rule is ``RecalcCheck``'s and Excel's (capture
``volatile_names``).
"""

from __future__ import annotations

import io
import itertools
from collections.abc import Callable

import pytest
import xlio

import xlrecalc
from tests.test_loader_orders import Reads
from tests.test_volatile_names import REFS, _answer, _rows, _shown, _workbook
from tests.xlsx_files import DOC, MAIN, PKG, package
from xlrecalc import Report

FIRST = {  # sheet "one": constants, a shared formula, a legacy array formula, an absent function
    "A1": "<v>2</v>",
    "A2": "<v>3</v>",
    "A3": "<v>5</v>",
    "B1": '<f t="shared" ref="B1:B3" si="0">A1*10</f><v>0</v>',
    "B2": '<f t="shared" si="0"/><v>0</v>',
    "B3": '<f t="shared" si="0"/><v>0</v>',
    "C1": '<f t="array" ref="C1:C2">A1:A2+1</f><v>0</v>',
    "C2": "<v>0</v>",
    "D1": '<f>RTD("a","","b")</f><v>7</v>',
    "D2": "<f>D1+B3</f><v>0</v>",
    "E1": '<f ca="1">NOW()*0+1</f><v>0</v>',
}
SECOND = {  # sheet "two": reads sheet "one"
    "A1": "<f>one!B2+one!D2</f><v>0</v>",
    "A2": "<f>SUM(one!A1:A3)</f><v>0</v>",
    "A3": "<v>4</v>",
}


def _sheet(cells: dict[str, str]) -> str:
    rows: dict[int, list[str]] = {}
    for ref, body in cells.items():
        rows.setdefault(int(ref[1:]), []).append(f'<c r="{ref}">{body}</c>')
    data = "".join(f'<row r="{r}">{"".join(c)}</row>' for r, c in sorted(rows.items()))
    return f'<worksheet xmlns="{MAIN}"><sheetData>{data}</sheetData></worksheet>'


def two_sheets(calc: str = '<calcPr fullCalcOnLoad="1"/>') -> bytes:
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId3" Type="{DOC}/worksheet" Target="worksheets/sheet2.xml"/>'
        "</Relationships>"
    )
    return package(
        sheet_xml=_sheet(FIRST),
        sheets='<sheet name="one" sheetId="1" r:id="rId1"/>'
        '<sheet name="two" sheetId="2" r:id="rId3"/>',
        workbook_extra=calc,
        extra_parts={
            "xl/_rels/workbook.xml.rels": rels,
            "xl/worksheets/sheet2.xml": _sheet(SECOND),
        },
    )


def state(wb: xlrecalc.Workbook) -> tuple[object, ...]:
    """What a workbook holds: its values (before and after a recalculation) and its report."""
    refs = [f"one!{r}" for r in FIRST] + [f"two!{r}" for r in SECOND]
    report = wb.report
    before = (
        [wb.value(r) for r in refs],
        sorted(report.uncomputed),
        sorted(report.from_kept),
        dict(report.unsupported),
        dict(report.refused),
    )
    wb.recalculate()
    return (*before, [wb.value(r) for r in refs], sorted(wb.report.uncomputed))


def loaded(data: bytes, feed: object = None, **options: object) -> xlrecalc.Workbook:
    """A ``Loader``'s workbook; ``feed(loader, book)`` feeds what it likes before ``finish``."""
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book, **options)  # type: ignore[arg-type]
        if feed is None:
            for sheet in book.sheets:
                for _ in loader.cells(sheet):
                    pass
        else:
            feed(loader, book)  # type: ignore[operator]
        return loader.finish()


@pytest.mark.parametrize("mode", ["excel", "always", "never"])
@pytest.mark.parametrize("calc", ['<calcPr fullCalcOnLoad="1"/>', '<calcPr calcId="191029"/>'])
def test_fed_in_full_it_is_load(mode: str, calc: str) -> None:
    data = two_sheets(calc)
    expected = state(xlrecalc.load(io.BytesIO(data), recalc_on_load=mode))  # type: ignore[arg-type]
    assert state(loaded(data, recalc_on_load=mode)) == expected


def test_the_cells_pass_through_unchanged() -> None:
    data = two_sheets()
    with xlio.open(io.BytesIO(data)) as book:
        direct = [(c.row, c.column, c.value) for s in book.sheets for c in s.cells()]
        loader = xlrecalc.Loader(book)
        passed = [(c.row, c.column, c.value) for s in book.sheets for c in loader.cells(s)]
        loader.finish()
    assert passed == direct


@pytest.mark.parametrize("stop", [0, 1, 4, 7, 10, 11])
def test_a_sheet_left_part_way_and_one_not_fed_are_read_by_finish(stop: int) -> None:
    """The first sheet is left after ``stop`` cells (11 is all of them, without the loop's end);
    the second is not fed."""
    data = two_sheets()

    def feed(loader: xlrecalc.Loader, book: xlio.Workbook) -> None:
        received = list(itertools.islice(loader.cells(book.sheets[0]), stop))
        assert len(received) == stop

    assert state(loaded(data, feed)) == state(xlrecalc.load(io.BytesIO(data)))


def test_a_loop_still_held_at_finish_hands_on_the_rest_unrecorded() -> None:
    """The caller keeps the iterator alive past ``finish``: what it recorded counts once, and the
    rest of its cells still come, recorded by the loader's own reading only."""
    data = two_sheets()
    with xlio.open(io.BytesIO(data)) as book:
        direct = [(c.row, c.column, c.value) for c in book.sheets[0].cells()]
        loader = xlrecalc.Loader(book)
        held = loader.cells(book.sheets[0])
        first = [next(held) for _ in range(5)]
        wb = loader.finish()
        rest = list(held)
    assert [(c.row, c.column, c.value) for c in first + rest] == direct
    assert state(wb) == state(xlrecalc.load(io.BytesIO(data)))


def test_sheets_fed_out_of_order_give_the_same_workbook() -> None:
    data = two_sheets()

    def feed(loader: xlrecalc.Loader, book: xlio.Workbook) -> None:
        for sheet in reversed(book.sheets):
            for _ in loader.cells(sheet):
                pass

    assert state(loaded(data, feed)) == state(xlrecalc.load(io.BytesIO(data)))


def test_misuse_raises() -> None:
    data = two_sheets()
    with xlio.open(io.BytesIO(data)) as book, xlio.open(io.BytesIO(data)) as other:
        loader = xlrecalc.Loader(book)
        list(loader.cells(book.sheets[0]))
        with pytest.raises(ValueError, match="already been fed"):
            loader.cells(book.sheets[0])
        with pytest.raises(ValueError, match="not a sheet of this workbook"):
            loader.cells(other.sheets[0])
        wb = loader.finish()
        with pytest.raises(ValueError, match="already finished"):
            loader.finish()
        assert loader.workbook() is wb
    with pytest.raises(ValueError, match="recalc_on_load"):
        xlrecalc.Loader(book, recalc_on_load="sometimes")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="no file to copy"):
        wb.save()


@pytest.mark.parametrize("rich", [False, True])
@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_the_opening_rule_is_the_check_s_and_excel_s(row: list[str], rich: bool) -> None:
    """One stream feeds a ``RecalcCheck`` and a ``Loader``: the loader computes what Excel showed
    on opening, and the check answers what the loader computed."""
    data = _workbook(row)
    check = xlrecalc.RecalcCheck(io.BytesIO(data))
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        for sheet in book.sheets:
            for cell in loader.cells(sheet, rich=rich):
                f = cell.formula
                if f is not None:
                    ca = f.attributes.get("ca") in ("1", "true")
                    check.add(f.text, cached=cell.value is not None, ca=ca)
        wb = loader.finish()
    values = {ref: wb.value(ref) for ref in REFS}
    assert values == _shown(row), row[0]
    assert check.finish() == _answer(values), row[0]


_RICH_STRINGS = [
    # runs and a phonetic guide
    '<si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> text</t></r>'
    '<rPh sb="0" eb="4"><t>ボ</t></rPh><phoneticPr fontId="0"/></si>',
    "<si><t>plain</t></si>",
]
_RICH_CELLS = (
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>LEN(A1)</f><v>0</v></c></row>'
    '<row r="2"><c r="A2" t="s" ph="1"><v>1</v></c>'
    '<c r="B2" t="str"><f>A1&amp;A2</f><v>x</v></c></row>'
    '<row r="3"><c r="A3" t="inlineStr"><is><r><t>in</t></r><r><rPr><i/></rPr><t>line</t></r>'
    '<rPh sb="0" eb="2"><t>イン</t></rPh></is></c>'
    '<c r="B3"><f>EXACT(A3,"inline")</f><v>0</v></c></row>'
    '<row r="4"><c r="A4"><v>2.5</v></c><c r="B4"><f>A4*2</f><v>0</v></c></row>'
)


def _rich_book(calc: str) -> bytes:
    return package(_RICH_CELLS, strings=_RICH_STRINGS, workbook_extra=calc)


def _rich_state(wb: xlrecalc.Workbook) -> tuple[object, ...]:
    refs = [f"{c}{r}" for r in range(1, 5) for c in "AB"]
    report = wb.report
    before = ([wb.value(r) for r in refs], sorted(report.uncomputed), sorted(report.from_kept))
    wb.recalculate()
    return (*before, [wb.value(r) for r in refs])


def test_rich_cells_pass_through_as_xlio_gives_them() -> None:
    data = _rich_book('<calcPr fullCalcOnLoad="1"/>')
    with xlio.open(io.BytesIO(data)) as book:
        direct = list(book.sheets[0].cells(rich=True))
        loader = xlrecalc.Loader(book)
        passed = list(loader.cells(book.sheets[0], rich=True))
        loader.finish()
    assert passed == direct
    rich = passed[0].rich  # what a renderer reads
    assert rich is not None
    assert rich.phonetic


@pytest.mark.parametrize("calc", ['<calcPr fullCalcOnLoad="1"/>', '<calcPr calcId="191029"/>'])
@pytest.mark.parametrize("stop", [None, 0, 3])
def test_rich_or_plain_records_the_same_workbook(calc: str, stop: int | None) -> None:
    """Fed in full or left after ``stop`` cells, with runs or without: what ``load`` gives."""
    data = _rich_book(calc)

    def feed(rich: bool) -> object:
        def run(loader: xlrecalc.Loader, book: xlio.Workbook) -> None:
            cells = loader.cells(book.sheets[0], rich=rich)
            for _ in itertools.islice(cells, stop):
                pass

        return run

    expected = _rich_state(xlrecalc.load(io.BytesIO(data)))
    assert expected[0][:3] == ["Bold text", 9.0, "plain"]  # the value is the text, runs joined
    assert _rich_state(loaded(data, feed(True))) == expected
    assert _rich_state(loaded(data, feed(False))) == expected


@pytest.mark.parametrize("ca", [False, True])
@pytest.mark.parametrize("text", ["RAND()+", "(NOW()"])
def test_a_formula_the_parser_cannot_read_counts_by_its_text(text: str, ca: bool) -> None:
    """A volatile function in a cell's own text without ``ca="1"`` makes the load compute every
    formula (capture ``volatile_names``), also where the parser cannot read the rest of the text,
    as ``RecalcCheck`` answers."""
    mark = ' ca="1"' if ca else ""
    data = package(
        '<row r="1"><c r="A1"><v>2</v></c></row>'
        '<row r="2"><c r="A2"><f>A1*10</f><v>999</v></c></row>'
        f'<row r="3"><c r="A3"><f{mark}>{text}</f><v>0</v></c></row>',
        workbook_extra='<fileVersion appName="xl"/><calcPr calcId="191029"/>',
    )
    check = xlrecalc.RecalcCheck(io.BytesIO(data))
    check.add(text, ca=ca)
    check.add("A1*10")
    wb = xlrecalc.load(io.BytesIO(data))
    assert list(wb.report.refused) == [xlrecalc.CellRef("Data", 3, 1)]  # the parser's refusal
    assert wb.value("A2") == (999.0 if ca else 20.0)
    assert check.finish() == ("some" if ca else "all")


@pytest.mark.parametrize(
    ("calc", "first", "total"),
    [  # every formula computed: nothing read before it is needed; else the loader reads all
        ('<calcPr fullCalcOnLoad="1"/>', {}, {"one": 1, "two": 1}),
        ('<fileVersion appName="xl"/><calcPr calcId="191029"/>', {"one": 1, "two": 1}, None),
    ],
)
def test_the_workbook_before_any_feed_then_every_sheet_fed(
    monkeypatch: pytest.MonkeyPatch, calc: str, first: dict[str, int], total: object
) -> None:
    """Values asked first. Where the file alone makes every formula computed, the loader reads
    nothing until a value needs a sheet, and the caller's feeds after that are recorded; else it
    reads every sheet, and the caller's feeds pass each sheet's cells through again."""
    data = two_sheets(calc)
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        wb = loader.workbook()
        assert dict(reads.calls) == first
        assert loader.workbook() is wb
        passed = [[(c.row, c.column) for c in loader.cells(s)] for s in book.sheets]
        wb.set("one!A1", 100)
        assert loader.finish() is wb
    assert [len(p) for p in passed] == [11, 3]
    assert dict(reads.calls) == (total or {"one": 2, "two": 2})
    eager = xlrecalc.load(io.BytesIO(data))
    eager.set("one!A1", 100)
    assert state(wb) == state(eager)


def test_a_value_reads_only_the_sheets_it_needs(monkeypatch: pytest.MonkeyPatch) -> None:
    """Every formula computed (``fullCalcOnLoad``): a value of sheet "one", whose formulas read
    only their own sheet, reads "one" alone; one of "two", which reads "one", reads "two"."""
    data = two_sheets()
    eager = xlrecalc.load(io.BytesIO(data))
    expected = [eager.value("one!D2"), eager.value("two!A2")]
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        wb = xlrecalc.Loader(book).workbook()
        assert wb.value("one!D2") == expected[0]
        assert dict(reads.calls) == {"one": 1}
        assert wb.value("two!A2") == expected[1]
        assert dict(reads.calls) == {"one": 1, "two": 1}


def test_each_sheet_fed_in_full_is_read_once(monkeypatch: pytest.MonkeyPatch) -> None:
    data = two_sheets()
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        for sheet in reversed(book.sheets):
            for _ in loader.cells(sheet):
                pass
        loader.finish()
    assert dict(reads.calls) == {"one": 1, "two": 1}
    assert dict(reads.cells) == {"one": 11, "two": 3}


@pytest.mark.parametrize(
    ("calc", "final"),
    [
        ('<fileVersion appName="xl"/><calcPr calcId="152511"/>', "all"),  # an older Excel
        ('<fileVersion appName="xl"/><calcPr calcId="191029"/>', "some"),  # E1 is marked
    ],
)
def test_opening_is_known_before_any_sheet_is_read(
    monkeypatch: pytest.MonkeyPatch, calc: str, final: str
) -> None:
    """From the workbook part, or from a formulas-only scan of the sheets not fed
    (``Sheet.formulas``): no cell is read for it, and feeding afterwards changes nothing."""
    data = two_sheets(calc)
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        assert loader.opening == final
        assert dict(reads.calls) == {}
        assert dict(reads.scans) == ({} if final == "all" else {"one": 1, "two": 1})
        list(loader.cells(book.sheets[0]))
        list(loader.cells(book.sheets[1]))
        assert loader.opening == final
        loader.finish()
        assert loader.opening == final
    assert dict(reads.scans) == ({} if final == "all" else {"one": 1, "two": 1})


def test_a_sheet_needed_after_the_file_is_closed_raises() -> None:
    data = two_sheets()
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        list(loader.cells(book.sheets[0]))
    wb = loader.workbook()  # every formula computed: no sheet is read until a value needs it
    assert wb.value("one!B2") == 30.0
    with pytest.raises(ValueError, match="workbook is closed"):
        wb.value("two!A1")
    with xlio.open(
        io.BytesIO(two_sheets('<fileVersion appName="xl"/><calcPr calcId="191029"/>'))
    ) as book:
        loader = xlrecalc.Loader(book)
        list(loader.cells(book.sheets[0]))
    with pytest.raises(ValueError, match="workbook is closed"):
        loader.workbook()  # the opening rule needs every sheet's formulas
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        for sheet in book.sheets:
            list(loader.cells(sheet))
    assert state(loader.finish()) == state(xlrecalc.load(io.BytesIO(data)))


def test_a_repeated_formula_keeps_its_own_mark_and_cached_value() -> None:
    """Cells repeating one formula (one raw-text key) are recorded without their text: each still
    counts its own ``ca="1"`` and its own missing value (capture ``volatile_names``: a cell
    volatile through a name keeps its value unless marked)."""
    data = package(
        '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c>'
        '<c r="C1"><f>vn*0+1</f><v>5</v></c></row>'  # kept
        '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><f>A2*2</f></c>'  # no cached value
        '<c r="C2"><f ca="1">vn*0+1</f><v>7</v></c></row>',  # marked: computed
        workbook_extra='<fileVersion appName="xl"/><calcPr calcId="191029"/>'
        '<definedNames><definedName name="vn">NOW()</definedName></definedNames>',
    )
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        for sheet in book.sheets:
            list(loader.cells(sheet))
        wb = loader.finish()
        assert loader.opening == "some"
    assert [wb.value(r) for r in ("B1", "B2", "C1", "C2")] == [4.0, 6.0, 5.0, 1.0]
    check = xlrecalc.RecalcCheck(io.BytesIO(data))
    check.add("A2*2", cached=False)
    assert check.finish() == "some"


def test_a_repeated_formula_without_a_value_makes_the_opening_some() -> None:
    data = package(
        '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>'
        '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><f>A2*2</f></c></row>',
        workbook_extra='<fileVersion appName="xl"/><calcPr calcId="191029"/>',
    )
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        list(loader.cells(book.sheets[0]))
        assert loader.opening == "some"
        assert loader.finish().value("B2") == 6.0


def test_the_report_is_in_sheet_row_and_column_order() -> None:
    """Cells enter the report as they are compiled and evaluated; it shows them by position, and
    its labels in text order, however the file or the reading ordered them."""
    data = package(  # rows stored out of order, as a file may hold them
        '<row r="3"><c r="B3"><f>AREAS(1)</f><v>0</v></c>'
        '<c r="C3"><f t="dataTable" ref="C3:C4" dt2D="0" dtr="0" r1="A1"/><v>1</v></c></row>'
        '<row r="1"><c r="B1"><f>RAND()+</f><v>0</v></c>'
        '<c r="A1"><f>AREAS(2)</f><v>0</v></c>'
        '<c r="D1"><f t="array" ref="D1:D2"/><v>0</v></c></row>',
        workbook_extra='<calcPr fullCalcOnLoad="1"/>',
    )
    report = xlrecalc.load(io.BytesIO(data)).report
    assert [str(c) for c in report.refused] == ["Data!A1", "Data!B1", "Data!B3"]
    assert list(report.unsupported) == sorted(report.unsupported)
    assert len(report.unsupported) == 2


def _lazy_book(first: str, second: str, calc: str = '<calcPr fullCalcOnLoad="1"/>') -> bytes:
    """Two sheets, "one" and "two", from their sheet data; ``fullCalcOnLoad``: loaded lazily."""
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId3" Type="{DOC}/worksheet" Target="worksheets/sheet2.xml"/>'
        "</Relationships>"
    )
    sheets = (
        '<sheet name="one" sheetId="1" r:id="rId1"/><sheet name="two" sheetId="2" r:id="rId3"/>'
    )
    two = f'<worksheet xmlns="{MAIN}"><sheetData>{second}</sheetData></worksheet>'
    return package(
        sheet_xml=f'<worksheet xmlns="{MAIN}"><sheetData>{first}</sheetData></worksheet>',
        sheets=sheets,
        workbook_extra=calc,
        extra_parts={"xl/_rels/workbook.xml.rels": rels, "xl/worksheets/sheet2.xml": two},
    )


def _lazy(data: bytes, read: Callable[[xlrecalc.Workbook], object]) -> object:
    """What ``read`` gives from a ``Loader``'s workbook with no sheet fed."""
    with xlio.open(io.BytesIO(data)) as book:
        return read(xlrecalc.Loader(book).workbook())


def test_cell_without_a_reference_loads_the_active_sheet() -> None:
    """``CELL("contents")`` reads the active cell (sheet "one", A1) by no name."""
    data = _lazy_book(
        '<row r="1"><c r="A1"><v>42</v></c></row>',
        '<row r="1"><c r="A1"><f>CELL("contents")</f></c></row>',
    )
    assert _lazy(data, lambda wb: wb.value("two!A1")) == 42.0


def _reported(report: Report) -> tuple[object, object]:
    return dict(report.refused), dict(report.unsupported)


def test_the_report_and_a_recalculation_s_cover_sheets_not_read_yet() -> None:
    data = _lazy_book(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        '<row r="1"><c r="A1"><f>AREAS(1)</f><v>5</v></c>'
        '<c r="B1"><f t="dataTable" ref="B1:B2" dt2D="0" dtr="0" r1="A1"/><v>6</v></c></row>',
    )
    expected = _reported(xlrecalc.load(io.BytesIO(data)).report)
    assert expected == (
        {xlrecalc.CellRef("two", 1, 1): "an argument of AREAS that is no reference"},
        {"data table": 1},
    )
    assert _lazy(data, lambda wb: _reported(wb.report)) == expected
    assert _lazy(data, lambda wb: _reported(wb.recalculate())) == expected


def test_a_reference_without_a_sheet_loads_the_only_sheet() -> None:
    data = package(
        '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f></c></row>',
        workbook_extra='<calcPr fullCalcOnLoad="1"/>',
    )
    assert _lazy(data, lambda wb: wb.value("B1")) == 6.0


def test_with_iteration_every_sheet_is_read_first() -> None:
    """Loops on two sheets settle at different speeds; Excel runs them together, with one
    stopping test (capture ``iterate_reads``, ``lengths``): read first, one sheet's loop has the
    value it has when both are read."""
    calc = '<calcPr fullCalcOnLoad="1" iterate="1" iterateCount="100"/>'
    data = _lazy_book(
        '<row r="1"><c r="B1"><f>B1/2+1</f><v>0</v></c></row>',
        '<row r="1"><c r="C1"><f>E1*0.9+1</f><v>0</v></c><c r="D1"><f>C1</f><v>0</v></c>'
        '<c r="E1"><f>D1</f><v>0</v></c></row>',
        calc,
    )
    assert xlrecalc.load(io.BytesIO(data)).value("one!B1") == 2.0  # as long as the slower loop
    assert _lazy(data, lambda wb: wb.value("one!B1")) == 2.0


NONE = '<fileVersion appName="xl"/><calcPr calcId="191029"/>'  # as Excel saves: trusted
QUIET_ONE = '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*10</f><v>20</v></c></row>'
QUIET_TWO = (
    '<row r="1"><c r="A1"><f>one!B1+1</f><v>21</v></c><c r="B1"><f>SUM(A1:A1)</f><v>21</v></c>'
    "</row>"
)


def test_where_every_cached_value_stands_nothing_is_compiled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """``opening`` "none": values are the cached ones, read from the sheets fed, and from a sheet
    not fed only once a value needs it (its formulas scanned once, before, for the rule); no
    formula is compiled until one is needed, and then only its sheet's."""
    data = _lazy_book(QUIET_ONE, QUIET_TWO, NONE)
    reads = Reads(monkeypatch)
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        list(loader.cells(book.sheets[0]))
        assert loader.opening == "none"
        assert dict(reads.scans) == {"two": 1}
        wb = loader.workbook()
        engine = wb._engine
        assert (wb.value("one!B1"), len(engine.templates)) == (20.0, 0)
        assert dict(reads.calls) == {"one": 1}
        assert (wb.value("two!A1"), len(engine.templates)) == (21.0, 0)
        assert dict(reads.calls) == {"one": 1, "two": 1}
        assert wb.formula("one!B1") == "=A1*10"
        assert engine.formula_count() == 1  # sheet "one" alone
        assert wb.report.refused == {}
        assert engine.formula_count() == 3
        loader.finish()
    assert dict(reads.scans) == {"two": 1}


def test_where_every_cached_value_stands_an_edit_reaches_every_sheet() -> None:
    data = _lazy_book(QUIET_ONE, QUIET_TWO, NONE)

    def edit(wb: xlrecalc.Workbook) -> object:
        wb.set("one!A1", 5)
        return wb.value("two!A1"), wb.value("two!B1")

    assert _lazy(data, edit) == (51.0, 51.0)


def test_a_shared_formula_s_dependent_counts_for_the_rule() -> None:
    """A dependent cell of a shared formula holds no text: its missing value, or its mark, still
    decides the rule, in recording and in the scan alike."""
    for dependent, answer in (
        ('<f t="shared" si="0"/>', "some"),
        ('<f t="shared" si="0"/><v>3</v>', "none"),
    ):
        data = _lazy_book(
            '<row r="1"><c r="A1"><v>1</v></c></row>',
            '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f>'
            f'<v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2">{dependent}</c></row>',
            NONE,
        )
        with xlio.open(io.BytesIO(data)) as book:
            scanned = xlrecalc.Loader(book).opening
            loader = xlrecalc.Loader(book)
            for sheet in book.sheets:
                list(loader.cells(sheet))
            assert (scanned, loader.opening) == (answer, answer)
            assert loader.finish().value("two!B2") == (6.0 if answer == "some" else 3.0)


def test_formulas_compiled_late_keep_the_trusted_volatile_value() -> None:
    """``opening`` "none" with a volatile cell trusted (volatile through a name, unmarked): sheets
    compiled one after the other, as their formulas are asked for, keep its cached value."""
    names = '<definedNames><definedName name="vn">OFFSET(one!$A$1,0,0)</definedName></definedNames>'
    data = _lazy_book(
        '<row r="1"><c r="A1"><v>0.1</v></c><c r="B1"><f>vn</f><v>7</v></c></row>',
        '<row r="1"><c r="A1"><f>one!A1*2</f><v>0.2</v></c></row>',
        NONE + names,
    )
    eager = xlrecalc.load(io.BytesIO(data))
    expected = (eager.formula("two!A1"), eager.formula("one!B1"), eager.value("one!B1"))
    assert expected == ("=one!A1*2", "=vn", 7.0)

    def read(wb: xlrecalc.Workbook) -> object:
        return wb.formula("two!A1"), wb.formula("one!B1"), wb.value("one!B1")

    assert _lazy(data, read) == expected


def test_a_sheet_read_after_a_cycle_closed_keeps_its_dependents_of_it() -> None:
    """A cell depending on a cycle keeps its value, as Excel keeps it (capture ``circular_cached``),
    also when its sheet is read after the cycle closed: ``one!A1`` read after ``two``'s cycle
    was 8, computed from the cycle's kept value."""
    data = _lazy_book(
        '<row r="1"><c r="A1"><f>two!A1+1</f><v>7</v></c></row>',
        '<row r="1"><c r="A1"><f>A2+1</f><v>7</v></c></row>'
        '<row r="2"><c r="A2"><f>A1+1</f><v>7</v></c></row>',
    )
    eager = xlrecalc.load(io.BytesIO(data))
    assert eager.value("one!A1") == 7.0
    cells = {xlrecalc.CellRef(*c) for c in (("one", 1, 1), ("two", 1, 1), ("two", 2, 1))}
    assert eager.report.circular == cells

    def cycle_first(wb: xlrecalc.Workbook) -> object:
        return wb.value("two!A1"), wb.value("one!A1"), wb.report.circular

    assert _lazy(data, cycle_first) == (7.0, 7.0, cells)


def test_a_sheet_read_during_an_evaluation_that_closes_a_cycle_keeps_what_the_cycle_kept() -> None:
    """``two`` is read while ``one``'s cycle is evaluated (its ``A1`` reads ``two!B1``); the cycle
    keeps ``two!A2`` and the array ``two!D1:D2`` reading it, as a full load keeps them. Reading
    the sheet was taken for an edit, and the array, kept, was made stale again and computed."""
    data = _lazy_book(
        '<row r="1"><c r="A1"><f>A2+two!B1*0</f><v>7</v></c></row>'
        '<row r="2"><c r="A2"><f>A1+1</f><v>7</v></c></row>',
        '<row r="1"><c r="A1"><f>CUBEVALUE("c")</f></c><c r="B1"><v>1</v></c>'
        '<c r="D1"><f t="array" ref="D1:D2">A1:A2*2</f><v>5</v></c></row>'
        '<row r="2"><c r="A2"><f>one!A1+1</f><v>7</v></c><c r="D2"><v>6</v></c></row>',
    )

    def read(wb: xlrecalc.Workbook) -> object:
        cells = [wb.value(r) for r in ("one!A1", "two!D1", "two!D2")]
        return cells, sorted(wb.report.circular), sorted(wb.report.uncomputed)

    eager = read(xlrecalc.load(io.BytesIO(data)))
    assert eager[0] == [7.0, 5.0, 6.0]
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        for _ in loader.cells(book.sheets[0]):
            pass
        assert read(loader.workbook()) == eager
