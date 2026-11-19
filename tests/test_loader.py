"""``Loader``: a workbook read from the cells a caller streams through ``xlio``, read once.

Fed in full it gives exactly ``load``'s workbook; a sheet not fed, or left part-way, is read again
by ``finish``; misuse raises; and its opening rule is ``RecalcCheck``'s and Excel's (capture
``volatile_names``).
"""

from __future__ import annotations

import io
import itertools

import pytest
import xlio

import xlrecalc
from tests.test_volatile_names import REFS, _answer, _rows, _shown, _workbook
from tests.xlsx_files import DOC, MAIN, PKG, package

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


def test_a_loop_still_held_at_finish_records_nothing_twice() -> None:
    """The caller keeps the iterator alive past ``finish``: what it recorded counts once, and it
    records nothing after."""
    data = two_sheets()
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        held = loader.cells(book.sheets[0])
        for _ in range(5):
            next(held)
        wb = loader.finish()
        assert list(held) == []
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
        with pytest.raises(ValueError, match="after finish"):
            loader.cells(book.sheets[1])
        with pytest.raises(ValueError, match="already finished"):
            loader.finish()
    with pytest.raises(ValueError, match="recalc_on_load"):
        xlrecalc.Loader(book, recalc_on_load="sometimes")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="no file to copy"):
        wb.save()


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_the_opening_rule_is_the_check_s_and_excel_s(row: list[str]) -> None:
    """One stream feeds a ``RecalcCheck`` and a ``Loader``: the loader computes what Excel showed
    on opening, and the check answers what the loader computed."""
    data = _workbook(row)
    check = xlrecalc.RecalcCheck(io.BytesIO(data))
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book)
        for sheet in book.sheets:
            for cell in loader.cells(sheet):
                f = cell.formula
                if f is not None:
                    ca = f.attributes.get("ca") in ("1", "true")
                    check.add(f.text, cached=cell.value is not None, ca=ca)
        wb = loader.finish()
    values = {ref: wb.value(ref) for ref in REFS}
    assert values == _shown(row), row[0]
    assert check.finish() == _answer(values), row[0]
