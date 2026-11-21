"""``Loader``: any feeding order gives ``load``'s workbook, and no sheet is read more than needed.

Random workbooks of up to four sheets, whose formulas reach across sheets through references,
ranges, defined names, ``INDIRECT`` and ``OFFSET``, with shared and array formulas, cells marked
``ca="1"``, volatile functions with and without the mark, formulas the parser cannot read and
formulas Excel refuses in a file, cycles within and across sheets, and cells with and without a
cached value. Each is loaded under every calculation case and ``recalc_on_load``, by ``load`` and
by a ``Loader`` fed in a random order: sheets fed in full, left part-way, held open past the
moment values are asked, never fed, or fed after the loader has read them; values asked at
random moments. Every value (before and after a recalculation, bit for bit), the report (in its
order) and ``opening`` must be ``load``'s, and ``xlio``'s ``Sheet.cells`` is counted: a sheet the
caller fed in full before values were asked is read once.
"""

from __future__ import annotations

import io
import itertools
import random
from collections import Counter
from collections.abc import Iterator

import pytest
import xlio

import xlrecalc
from tests.xlsx_files import DOC, MAIN, PKG, package
from xlrecalc._load import refused_in_file

CALCS = {
    "full": '<calcPr fullCalcOnLoad="1"/>',
    "current": '<fileVersion appName="xl"/><calcPr calcId="191029"/>',
    "older": '<fileVersion appName="xl"/><calcPr calcId="152511"/>',
    "manual": '<fileVersion appName="xl"/><calcPr calcId="191029" calcMode="manual"/>',
    "iterate": '<fileVersion appName="xl"/><calcPr calcId="191029" iterate="1"/>',
}
MODES = ["excel", "always", "never"]
ROWS = 5
NAMES = (
    '<definedName name="nm">s1!$A$1</definedName>'
    '<definedName name="vn">OFFSET(s1!$A$2,0,0)</definedName>'
)


VOLATILE = ("INDIRECT", "OFFSET", "RAND", "NOW", "vn")


def _formula(rng: random.Random, sheets: int, row: int, tame: bool, quiet: bool) -> str:
    """A formula; ``tame``, one Excel could have saved (it reads and is not refused); ``quiet``,
    one with no volatile function or name."""
    other = f"s{rng.randint(1, sheets)}"
    r = rng.randint(1, ROWS)
    pool = [
        f"A{r}*2",
        f"SUM(A1:A{ROWS})",
        f"{other}!A{r}+1",
        f"SUM({other}!A1:A3)",
        f'INDIRECT("{other}!A{r}")',
        f"OFFSET({other}!A1,{r - 1},0)",
        f"B{r}+1",  # a chain, or a cycle within the sheet
        f"{other}!B{r}+C{row}",  # a chain, or a cycle across sheets
        f"{other}!E{r}*2",  # a shared formula's cell elsewhere
        "RAND()*0+1",
        "NOW()*0+2",
        "nm+1",
        "vn*3",
    ]
    if quiet:
        pool = [f for f in pool if not any(word in f for word in VOLATILE)]
    if not tame:
        pool += ["RAND()+", "AREAS(1)"]  # the parser cannot read it; Excel refuses it in a file
    return rng.choice(pool)


def _cell(ref: str, body: str) -> str:
    return f'<c r="{ref}">{body}</c>'


def _formula_cell(rng: random.Random, ref: str, inner: str, tame: bool) -> str:
    """A formula cell; ``tame``, with a cached value, and marked only where Excel marks a cell
    (a volatile function or name in its text; Excel marks its dependents too)."""
    cached = rng.choice(["<v>0</v>", "<v>7.5</v>", "<v>-3</v>"] + ([] if tame else [""]))
    volatile = any(word in inner for word in VOLATILE)
    ca = ' ca="1"' if (volatile if tame else rng.random() < 0.25) else ""
    return _cell(ref, inner.replace("<f", f"<f{ca}", 1) + cached)


def _sheet_xml(rng: random.Random, sheets: int, tame: bool, quiet: bool) -> str:
    rows: dict[int, list[str]] = {r: [] for r in range(1, ROWS + 1)}
    for r in range(1, ROWS + 1):
        if rng.random() < 0.8:
            rows[r].append(_cell(f"A{r}", f"<v>{rng.choice([1, 2.5, -4, 10, 0.1])}</v>"))
        for col in "BC":
            if rng.random() < 0.7:
                text = (
                    _formula(rng, sheets, r, tame, quiet)
                    .replace("&", "&amp;")
                    .replace('"', "&quot;")
                )
                rows[r].append(_formula_cell(rng, f"{col}{r}", f"<f>{text}</f>", tame))
    if rng.random() < 0.7:  # a shared formula over E1:E3, its master in E1
        rows[1].append(_formula_cell(rng, "E1", '<f t="shared" ref="E1:E3" si="0">A1*10</f>', tame))
        for r in (2, 3):
            rows[r].append(_formula_cell(rng, f"E{r}", '<f t="shared" si="0"/>', tame))
    if rng.random() < 0.5:  # a legacy array formula over F1:F2
        arrays = ["A1:A2*2", "s1!A1:A2+1"] + ([] if quiet else ["RAND()*0+A1:A2"])
        text = rng.choice(arrays + ([] if tame else ["RAND()+"]))
        rows[1].append(_formula_cell(rng, "F1", f'<f t="array" ref="F1:F2">{text}</f>', tame))
        rows[2].append(_cell("F2", rng.choice(["", "<v>4</v>"])))
    if rng.random() < 0.3:  # a data table over G1:G2, kept with its cached values and counted
        body = '<f t="dataTable" ref="G1:G2" dt2D="0" dtr="0" r1="A1"/>'
        rows[1].append(_formula_cell(rng, "G1", body, tame))
        rows[2].append(_cell("G2", "<v>5</v>"))
    data = "".join(f'<row r="{r}">{"".join(c)}</row>' for r, c in rows.items() if c)
    return f'<worksheet xmlns="{MAIN}"><sheetData>{data}</sheetData></worksheet>'


def _book(rng: random.Random, calc: str) -> tuple[bytes, list[str]]:
    """A random workbook and every cell reference worth reading in it."""
    sheets = rng.randint(1, 4)
    tame = rng.random() < 0.6
    quiet = tame and rng.random() < 0.5
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{DOC}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheets + 1)
    )
    parts: dict[str, str | bytes] = {
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">{rels}</Relationships>'
    }
    for i in range(1, sheets + 1):
        parts[f"xl/worksheets/sheet{i}.xml"] = _sheet_xml(rng, sheets, tame, quiet)
    data = package(
        sheets="".join(
            f'<sheet name="s{i}" sheetId="{i}" r:id="rId{i}"/>' for i in range(1, sheets + 1)
        ),
        workbook_extra=calc + f"<definedNames>{NAMES}</definedNames>",
        extra_parts=parts,
    )
    refs = [f"s{i}!{c}{r}" for i in range(1, sheets + 1) for c in "ABCEFG" for r in range(1, 6)]
    return data, refs


def _key(value: object) -> object:
    return value.hex() if isinstance(value, float) else repr(value)


def state(wb: xlrecalc.Workbook, refs: list[str]) -> tuple[object, ...]:
    """Every value before and after a recalculation, bit for bit, and the report in its order."""
    before = [_key(wb.value(r)) for r in refs]
    report = wb.report
    reported = (
        list(report.unsupported.items()),
        list(report.refused.items()),
        sorted(report.uncomputed),
        sorted(report.from_kept),
        sorted(report.circular),
    )
    wb.recalculate()
    return before, reported, [_key(wb.value(r)) for r in refs], sorted(wb.report.uncomputed)


class Reads:
    """Counts ``xlio``'s ``Sheet.cells`` calls and the cells each gives, by sheet name."""

    def __init__(self, monkeypatch: pytest.MonkeyPatch) -> None:
        self.calls: Counter[str] = Counter()
        self.cells: Counter[str] = Counter()
        original = xlio.Sheet.cells

        def counted(sheet: xlio.Sheet, *args: object, **kwargs: object) -> Iterator[xlio.Cell]:
            self.calls[sheet.name] += 1
            for cell in original(sheet, *args, **kwargs):  # type: ignore[arg-type]
                self.cells[sheet.name] += 1
                yield cell

        monkeypatch.setattr(xlio.Sheet, "cells", counted)


def _edit(wb: xlrecalc.Workbook) -> xlrecalc.Workbook:
    """A change to every sheet's last row, made as soon as the workbook is first given: a sheet
    recorded again afterwards would undo it."""
    for name in wb.sheet_names:
        wb.set(f"{name}!A5", 99.0)
    return wb


def _feed(
    rng: random.Random, data: bytes, refs: list[str], mode: str, reads: Reads
) -> tuple[tuple[object, ...], str | None, set[str], list[str]]:
    """A ``Loader`` fed at random; its state, its final ``opening``, the sheets the caller fed in
    full before values were asked, and the values asked while feeding, in order (asked of the
    eager workbook too, so that the comparison does not rest on the order of reads changing
    nothing, which ``tests/test_iteration_kept.py`` checks)."""
    edited: list[xlrecalc.Workbook] = []

    def workbook() -> xlrecalc.Workbook:
        wb = loader.workbook()
        if not edited:
            edited.append(_edit(wb))
        return wb

    with xlio.open(io.BytesIO(data)) as book:
        direct = {s.name: [(c.row, c.column, c.value) for c in s.cells()] for s in book.sheets}
        reads.calls.clear()
        reads.cells.clear()
        loader = xlrecalc.Loader(book, recalc_on_load=mode)  # type: ignore[arg-type]
        early = loader.opening
        order = list(book.sheets)
        rng.shuffle(order)
        held: list[tuple[str, list[tuple[int, int, object]], Iterator[xlio.Cell]]] = []
        whole: set[str] = set()
        fed: set[str] = set()
        asked: list[str] = []
        for sheet in order:
            if rng.random() < 0.2:  # values asked between two sheets
                asked.append(rng.choice(refs))
                workbook().value(asked[-1])
            answer = loader.opening
            assert answer is None or early is None or answer == early
            early = answer if answer is not None else early
            how = rng.choice(["full", "full", "part", "held", "none"])
            if how == "none":
                continue
            cells = loader.cells(sheet)
            fed.add(sheet.name)
            if how == "full":
                got = [(c.row, c.column, c.value) for c in cells]
                assert got == direct[sheet.name]
                if not asked:
                    whole.add(sheet.name)
                continue
            first = [(c.row, c.column, c.value) for c in itertools.islice(cells, rng.randint(0, 4))]
            if how == "held":
                held.append((sheet.name, first, cells))
        finished = rng.random() < 0.5
        wb = loader.finish() if finished else workbook()
        if not edited:
            edited.append(_edit(wb))
        for name, first, cells in held:  # the rest still comes, unrecorded
            assert first + [(c.row, c.column, c.value) for c in cells] == direct[name]
        for sheet in book.sheets:  # every sheet can still be fed: passed through
            if sheet.name in fed:
                with pytest.raises(ValueError, match="already been fed"):
                    loader.cells(sheet)
            elif rng.random() < 0.5:
                got = [(c.row, c.column, c.value) for c in loader.cells(sheet)]
                assert got == direct[sheet.name]
        assert (loader.workbook() if finished else loader.finish()) is wb
        opening = loader.opening
        assert early is None or opening == early
        return state(wb, refs), opening, whole, asked


def _checked(data: bytes, mode: str) -> str:
    """``RecalcCheck``'s answer, fed every formula cell but those Excel refuses in a file."""
    check = xlrecalc.RecalcCheck(io.BytesIO(data), recalc_on_load=mode)  # type: ignore[arg-type]
    with xlio.open(io.BytesIO(data)) as book:
        for sheet in book.sheets:
            for cell in sheet.cells():
                f = cell.formula
                if f is not None and (f.text is None or refused_in_file(f.text) is None):
                    ca = f.attributes.get("ca") in ("1", "true")
                    check.add(f.text, cached=cell.value is not None, ca=ca)
    return check.finish()


def _check(seed: int, calc: str, mode: str, monkeypatch: pytest.MonkeyPatch) -> None:
    rng = random.Random(f"{seed}/{calc}/{mode}")
    data, refs = _book(rng, CALCS[calc])
    eager = xlrecalc.load(io.BytesIO(data), recalc_on_load=mode)  # type: ignore[arg-type]
    opening = _checked(data, mode)
    reads = Reads(monkeypatch)
    got, got_opening, whole, asked = _feed(rng, data, refs, mode, reads)
    _edit(eager)
    for ref in asked:  # the same cells evaluated first, in the same order
        eager.value(ref)
    assert got == state(eager, refs)
    assert got_opening == opening
    for name in whole:  # fed in full before any value was asked: read once
        assert reads.calls[name] == 1, (name, reads.calls)


@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("calc", list(CALCS))
@pytest.mark.parametrize("seed", range(12))
def test_any_feeding_order_gives_load_s_workbook(
    seed: int, calc: str, mode: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    _check(seed, calc, mode, monkeypatch)


@pytest.mark.slow
@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("calc", list(CALCS))
def test_any_feeding_order_gives_load_s_workbook_many(
    calc: str, mode: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    for seed in range(12, 200):
        _check(seed, calc, mode, monkeypatch)
        monkeypatch.undo()
