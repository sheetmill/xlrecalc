"""Files written by ``recalculate`` and the writer, opened by Excel (ARCHITECTURE.md §9.2, §9.4).

Run with ``--tier excel`` on a Mac with Excel 16 (``CAPTURING.md`` §2: workbooks live in Excel's
container, alerts are off and restored, only workbooks these tests create are opened). Excel
opens each written file and saves a copy; a file Excel would offer to repair fails to open
with error -50, which fails the test.

- Every capture suite's workbook, written by openpyxl and recalculated here: Excel recalculates
  it on open (``fullCalcOnLoad``), and every formula cell's value in its saved copy must equal
  the value written here, except the known deviations of ``tests/test_captures.py``.
- A file Excel saved, given sentinel cached values of every kind with ``full_calc_on_load``
  off: Excel trusts them, and its saved copy must hold exactly what was written, so Excel reads
  the written text (whitespace, escapes, 17-digit numbers) as intended.
- The benchmark workbooks (``bench/data/``), recalculated: every formula cell equal.
- Every corpus formula, set through ``Workbook.set_formula`` and saved: Excel opens the file
  (a refused set is halved until each refused formula stands alone).
- A file Excel saved (shared formulas, a calculation chain), given edits of every kind through
  the API: Excel opens it, and with ``fullCalcOnLoad`` recalculates it to the same values; without
  it shows the written ones.
"""

from __future__ import annotations

import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest

import xlrecalc
from tests.test_captures import (
    KNOWN_DEVIATIONS,
    relative,
    same,
    tolerance,
    within,
    within_relative,
)
from xlrecalc import CellRef, FormulaError, recalculate
from xlrecalc._formula._fileform import display_form
from xlrecalc._load import load
from xlrecalc._values import ErrorValue
from xlrecalc._workbook import _file_formula
from xlrecalc._xlsx import Book
from xlrecalc._xlsx._write import write

pytestmark = pytest.mark.excel_app

ROOT = Path(__file__).resolve().parents[1]
VOLATILE = ("NOW(", "RAND(", "RANDBETWEEN(")


@pytest.fixture(scope="module")
def container() -> Path:
    from capture.pipeline import CONTAINER

    CONTAINER.mkdir(parents=True, exist_ok=True)
    return CONTAINER


def save_copy(src: Path, copy: Path) -> None:
    """Have Excel open ``src`` (alerts off) and save it as ``copy``; fail on a repair."""
    copy.unlink(missing_ok=True)
    result = subprocess.run(
        ["osascript", str(ROOT / "capture" / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, f"Excel did not open {src.name}: {result.stderr.strip()}"


def values(path: Path, sheet: str | None = None) -> dict[tuple[int, int], object]:
    """Each formula cell's cached value, as the engine holds values."""
    out: dict[tuple[int, int], object] = {}
    with Book(path) as book:
        info = book.sheets[0] if sheet is None else book.sheet(sheet)
        for c in book.cells(info):
            if c.formula is None:
                continue
            if c.type == "e":
                out[c.row, c.column] = ErrorValue(str(c.value))
            elif c.type == "none":
                out[c.row, c.column] = None
            else:
                out[c.row, c.column] = c.value
    return out


def suite_names() -> list[str]:
    try:  # the suites build workbooks with openpyxl, a development dependency
        from capture import suites
    except ImportError:
        return []
    return [name for name in suites.SUITES if name != "iteration"]  # evidence, not matched yet


@pytest.mark.parametrize("name", suite_names())
def test_capture_suites_recalculated_here_match_excel(container, name):
    from capture import suites
    from capture.pipeline import build

    suite = suites.SUITES[name]
    cases = suite.cases()
    src = container / f"xlrecalc_live_{name}.xlsx"
    copy = container / f"xlrecalc_live_{name}_saved.xlsx"
    src.unlink(missing_ok=True)
    build(src, cases, suite.date1904, suite.fixed)
    recalculate(src)
    save_copy(src, copy)
    ours, excel = values(src), values(copy)
    differ = []
    for row, case in enumerate(cases, 1):
        if any(v in case.formula.upper() for v in VOLATILE) or (name, row) in KNOWN_DEVIATIONS:
            continue
        a, b = ours.get((row, 1)), excel.get((row, 1))
        if a is None:  # uncomputed here: the cached value is left as it was
            continue
        if (
            not same(a, b)
            and not within(a, b, tolerance(case.formula))
            and not within_relative(a, b, relative(case.formula))
        ):
            differ.append((row, case.formula, a, b))
    src.unlink()
    copy.unlink()
    assert not differ, differ[:10]


SENTINELS: list[tuple[str, float | str | bool]] = [
    ("n", 0.1 + 0.2),
    ("n", 1 / 3),
    ("n", 1e16),
    ("n", -2.5e-300),
    ("n", 123456789012345680.0),
    ("n", 42.0),
    ("str", " leading"),
    ("str", "trailing "),
    ("str", "  "),
    ("str", "line\nbreak"),
    ("str", "carriage\rreturn"),
    ("str", "tab\there"),
    ("str", "a_x0041_b"),
    ("str", "control\x01char"),
    ("str", "Résumé 日本 🙂"),
    ("str", ""),
    ("str", "&<>\"'"),
    ("b", True),
    ("b", False),
    *[("e", code) for code in ("#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NUM!", "#NULL!")],
]


def test_excel_reads_written_values_exactly(container):
    import openpyxl

    src = container / "xlrecalc_live_sentinels.xlsx"
    base = container / "xlrecalc_live_sentinels_base.xlsx"
    patched = container / "xlrecalc_live_sentinels_patched.xlsx"
    copy = container / "xlrecalc_live_sentinels_saved.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r in range(1, len(SENTINELS) + 1):
        ws.cell(r, 1, "=ROW()*1000")
    wb.save(src)
    save_copy(src, base)  # Excel's own file: its calcId, its values, its calcChain

    def lookup(row: int, col: int):
        return SENTINELS[row - 1] if col == 1 and row <= len(SENTINELS) else None

    assert write(base, patched, {"probe": lookup}, full_calc_on_load=False) == len(SENTINELS)
    save_copy(patched, copy)
    got = values(copy)
    wrong = []
    for r, (kind, value) in enumerate(SENTINELS, 1):
        expected = ErrorValue(str(value)) if kind == "e" else value
        if not same(got[r, 1], expected):
            wrong.append((r, kind, value, got[r, 1]))
    for p in (src, base, patched, copy):
        p.unlink()
    assert not wrong, wrong


@pytest.mark.parametrize("n", [1000, 10000, 100000])
def test_benchmark_workbooks(container, n):
    data = ROOT / "bench" / "data" / f"bench_{n}.xlsx"
    if not data.exists():
        pytest.skip(f"{data.name} not generated (bench/gen.py)")
    src = container / f"xlrecalc_live_bench_{n}.xlsx"
    copy = container / f"xlrecalc_live_bench_{n}_saved.xlsx"
    shutil.copyfile(data, src)
    report = recalculate(src)
    assert not report.uncomputed
    assert not report.circular
    save_copy(src, copy)
    ours, excel = values(src), values(copy)
    assert len(ours) == 3 * n + 3
    differ = [(k, ours[k], excel.get(k)) for k in ours if not same(ours[k], excel.get(k))]
    assert not differ, differ[:10]
    engine = load(src)  # the written values read back; fullCalcOnLoad makes every formula stale
    assert engine.formula_count() == 3 * n + 3
    src.unlink()
    copy.unlink()


def _printed_workbook(container: Path, formulas: list[str], tag: str) -> Path:
    """The formula_forms workbook with ``formulas`` set through the API and saved here."""
    from capture.formula_forms import COLUMN, build

    src = container / f"xlrecalc_live_printer_{tag}.xlsx"
    build(src, ["0"] * len(formulas))
    wb = xlrecalc.load(src)
    for k, formula in enumerate(formulas):
        wb.set_formula(CellRef("Sheet1", 2 * k + 1, COLUMN), formula)
    wb.save()
    return src


def _opens(container: Path, path: Path) -> bool:
    copy = container / (path.stem + "_saved.xlsx")
    copy.unlink(missing_ok=True)
    result = subprocess.run(
        ["osascript", str(ROOT / "capture" / "save_copy.applescript"), str(path), str(copy)],
        capture_output=True,
        text=True,
    )
    copy.unlink(missing_ok=True)
    if result.returncode and "-50" not in result.stderr:
        raise RuntimeError(result.stderr.strip())
    return result.returncode == 0


def test_excel_opens_every_printed_corpus_formula(container):
    from capture.formula_forms import corpus

    accepted, refused_by_api = [], []
    for written in corpus():
        typed = "=" + display_form(written)
        try:
            _file_formula(typed)
        except FormulaError:
            refused_by_api.append(typed)
            continue
        accepted.append(typed)
    refused_by_excel: list[str] = []
    pending = [accepted]
    while pending:  # halve a refused set until each refused formula stands alone
        group = pending.pop()
        path = _printed_workbook(container, group, str(len(pending)))
        if not _opens(container, path):
            if len(group) == 1:
                refused_by_excel.append(group[0])
            else:
                pending += [group[: len(group) // 2], group[len(group) // 2 :]]
        path.unlink()
    print(
        f"\nprinter: {len(accepted)} corpus formulas written, {len(refused_by_api)} refused by "
        f"set_formula as the formula bar refuses them"
    )
    assert not refused_by_excel, refused_by_excel


def test_edits_of_every_kind_open_and_match_excel(container):
    import openpyxl

    src = container / "xlrecalc_live_edits.xlsx"
    base = container / "xlrecalc_live_edits_base.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 11):
        ws.cell(r, 1, float(r))
        ws.cell(r, 2, f"=A{r}*2")  # filled down: Excel saves a shared formula
        ws.cell(r, 3, f"=B{r}+1")
    ws["E1"] = "=SUM(C1:C10)"
    wb.save(src)
    save_copy(src, base)  # Excel's own file: shared formulas, calcChain, its calcId
    with zipfile.ZipFile(base) as z:
        assert 't="shared"' in z.read("xl/worksheets/sheet1.xml").decode()
        assert "xl/calcChain.xml" in z.namelist()
    for full in (True, False):
        book = xlrecalc.load(base)
        book.set("A1", 100)  # a constant
        book.set_formula("B1", "=A1*3")  # the shared formula's master
        book.set("C5", 7)  # a formula made a constant: the chain must go
        book.set("A12", " text ")  # a new row with outer whitespace
        book.set_formula("D3", '=IFNA(C3,0)&"!"')  # a new cell in a row, a prefixed function
        book.set_formula("F20", "=SUM(A1:A12)")  # a new row at the end
        book.set("C7", None)  # a formula cleared
        out = container / f"xlrecalc_live_edits_{full}.xlsx"
        book.save(out, full_calc_on_load=full)
        copy = container / f"xlrecalc_live_edits_{full}_saved.xlsx"
        save_copy(out, copy)
        ours, excel = _all_values(out), _all_values(copy)
        differ = {k: (ours[k], excel.get(k)) for k in ours if not same(ours[k], excel.get(k))}
        assert not differ, (full, differ)
        out.unlink()
        copy.unlink()
    src.unlink()
    base.unlink()


def _all_values(path: Path) -> dict[tuple[int, int], object]:
    out: dict[tuple[int, int], object] = {}
    with Book(path) as book:
        for c in book.cells(book.sheets[0]):
            value = c.value
            if c.type == "e":
                value = ErrorValue(str(value))
            if value is not None:
                out[c.row, c.column] = value
    return out


def test_array_formulas_written_here_match_excel(container):
    """The array probe's workbook, recalculated here: Excel opens it and agrees on every cell."""
    from capture.array_probe import build

    src = container / "xlrecalc_live_arrays.xlsx"
    copy = container / "xlrecalc_live_arrays_saved.xlsx"
    build(src)
    recalculate(src)
    save_copy(src, copy)
    ours, excel = _all_values(src), _all_values(copy)
    differ = {
        k: (ours.get(k), excel.get(k))
        for k in set(ours) | set(excel)
        if not same(ours.get(k), excel.get(k))
    }
    assert not differ, differ
    src.unlink()
    copy.unlink()


def test_dynamic_arrays_set_through_the_api(container):
    import openpyxl

    src = container / "xlrecalc_live_dynamic.xlsx"
    copy = container / "xlrecalc_live_dynamic_saved.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, 6):
        ws.cell(r, 1, float(r))
    ws["F3"] = "in the way"
    wb.save(src)
    book = xlrecalc.load(src)
    book.set_formula("C1", "=A1:A5*2")  # spills
    book.set_formula("D1", "=SUM(A1:A5)")  # a plain formula
    book.set_formula("E3", "=@A1:A5")  # implicit intersection, in a dynamic array
    book.set_formula("F1", "=A1:A5")  # blocked: #SPILL!
    book.set_formula("G1", "=C1#+1")  # the spill range of C1
    book.set_formula("H1", "=SUM(C1#)")
    book.set_formula("I1", "={1,2;3,4}*A1:A2")
    book.save()
    save_copy(src, copy)
    ours, excel = _all_values(src), _all_values(copy)
    differ = {
        k: (ours.get(k), excel.get(k))
        for k in set(ours) | set(excel)
        if not same(ours.get(k), excel.get(k))
    }
    assert not differ, differ
    with zipfile.ZipFile(copy) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    for ref in ("C1:C5", "G1:G5", "I1:J2"):
        assert f'ref="{ref}"' in sheet
    src.unlink()
    copy.unlink()


def test_dynamic_references_written_here_match_excel(container):
    """The dynamic-reference probe's workbook, recalculated here: Excel agrees on every cell."""
    import re

    from capture.dynref_probe import build

    src = container / "xlrecalc_live_dynrefs.xlsx"
    copy = container / "xlrecalc_live_dynrefs_saved.xlsx"
    build(src)
    report = recalculate(src)
    save_copy(src, copy)
    ours, excel = _all_values(src), _all_values(copy)
    # CELL is volatile: Excel evaluates it again when it saves the copy, under the copy's name
    excel = {
        k: v.replace(copy.name, src.name) if isinstance(v, str) else v for k, v in excel.items()
    }
    with Book(src) as book:
        skip = {
            (c.row, c.column)
            for c in book.cells(book.sheets[0])
            if c.formula is not None
            and re.search(
                r'CELL\("(format|protect|width|prefix|color|parentheses)"', c.formula.text or ""
            )
        }
    differ = {
        k: (ours.get(k), excel.get(k))
        for k in set(ours) | set(excel)
        if k not in skip and not same(ours.get(k), excel.get(k))
    }
    assert not differ, differ
    assert len(report.circular) == 2  # INDIRECT("C1") in C1, OFFSET(C4,-1,0) in C3
    src.unlink()
    copy.unlink()
