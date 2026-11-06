"""Circular references with iteration off, as Excel leaves them (ARCHITECTURE.md §7.2).

Captures ``circular`` (alerts off) and ``circular_alerts`` (alerts on, the warning dismissed by a
person), from ``capture/circular_probe.py``: Excel calculates no cell on a cycle or depending on
one, and saves each without a cached value as the empty text. Capture ``circular_cached``, from
``capture/circular_cached_probe.py``: one with a cached value keeps it, on opening, with
``fullCalcOnLoad``, after a full recalculation and after saving.
"""

import ast
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import pytest

import xlrecalc
from tests.test_xlsx import package
from xlrecalc._engine import Engine
from xlrecalc._formula._tokens import column_number

CAPTURES = Path(__file__).with_name("captures")

# Excel finds cycles through the references a formula evaluates, and so does the engine: a
# reference in a branch IF does not take closes no cycle (=IF(FALSE,A13,5) in A13 is 5).
KNOWN_DEVIATIONS: dict[str, str] = {}


def rows(name: str) -> list[list[str]]:
    text = (CAPTURES / f"{name}.tsv").read_text(encoding="utf-8")
    return [line.split("\t") for line in text.splitlines() if not line.startswith("#")]


CIRCULAR = rows("circular")
ALERTS = rows("circular_alerts")


@pytest.fixture(scope="module")
def engine() -> Engine:
    sheet: dict[str, object] = {}
    for number, _tag, formula, constant, *_ in CIRCULAR:
        sheet[f"A{number}"] = "=" + formula
        if constant != "None":
            sheet[f"B{number}"] = ast.literal_eval(constant)
    e = Engine.from_cells({"probe": sheet})
    e.recalculate()
    return e


@pytest.mark.excel
@pytest.mark.parametrize("row", CIRCULAR, ids=lambda r: f"{r[0]}:{r[1]}")
def test_engine_matches_excel(engine, row):
    number, tag, _formula, _constant, kind, text, *_ = row
    if tag in KNOWN_DEVIATIONS:
        pytest.xfail(KNOWN_DEVIATIONS[tag])
    got = engine.get(0, int(number) - 1, 0)
    expected = ast.literal_eval(text)
    if kind == "n":
        assert got == expected
    else:
        assert kind == "str"
        assert got == expected == ""
    on_cycle = ("probe", int(number), 1) in engine.report.circular
    assert on_cycle is (kind == "str")


@pytest.mark.excel
@pytest.mark.parametrize("pair", list(zip(CIRCULAR, ALERTS, strict=True)), ids=lambda p: p[0][0])
def test_alerts_do_not_change_the_result(pair):
    quiet, alerts = pair
    assert quiet[6] == alerts[6]  # the saved cell, alerts off and on
    assert quiet[7] == f"{alerts[4]}|{alerts[5]}"  # what Excel shows


@pytest.mark.excel
@pytest.mark.parametrize("row", CIRCULAR, ids=lambda r: f"{r[0]}:{r[1]}")
def test_cached_values_of_cycles_are_shown(row):
    """A file whose calculation settings let Excel trust its cached values shows them."""
    number, tag, *_rest, shown_patched = row
    if tag in ("pair", "self", "self_range", "triple", "text", "iferror", "untaken"):
        assert shown_patched == f"{900 + int(number)}.0|{900 + int(number)}"


INDEX_RANGE = [
    line.rstrip("\n").split("\t")
    for line in (Path(__file__).with_name("captures") / "index_range.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if not line.startswith("#")
]


@pytest.mark.parametrize("row", INDEX_RANGE, ids=lambda r: r[0])
def test_a_range_over_a_formulas_own_column_is_circular_only_as_read(row):
    """Capture ``index_range``: an exact lookup stops at its match, and ``INDEX`` reads nothing."""
    text, address, _kind, value = row
    column = address.rstrip("0123456789")
    number = column_number(column)
    cells: dict[str, object] = {f"{column}{r}": float(r) for r in range(1, 6)}
    cells[address] = "=" + text
    engine = Engine.from_cells({"probe": cells})
    engine.recalculate()
    r = int(address[len(column) :])
    assert engine.get(0, r - 1, number - 1) == ast.literal_eval(value)


STATIC_CYCLES = [
    line.rstrip("\n").split("\t")
    for line in (Path(__file__).with_name("captures") / "static_cycle.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if not line.startswith("#")
]


@pytest.mark.parametrize("group", sorted({r[0] for r in STATIC_CYCLES}))
def test_only_references_evaluated_make_a_cycle(group):
    """Capture ``static_cycle``: a loop through a branch not taken is not circular."""
    rows = [r for r in STATIC_CYCLES if r[0] == group]
    cells = {r[1]: ast.literal_eval(r[2]) for r in rows}
    engine = Engine.from_cells({"probe": cells})
    engine.recalculate()
    for _, address, formula, value in rows:
        if isinstance(ast.literal_eval(formula), str):
            column = address.rstrip("0123456789")
            r = int(address[len(column) :])
            assert engine.get(0, r - 1, column_number(column) - 1) == ast.literal_eval(value), (
                address
            )


CACHED = rows("circular_cached")
HELD = [r for r in CACHED if not r[0].startswith("excel_edited")]
EDITED = [r for r in CACHED if r[0].startswith("excel_edited")]


def shown(row: list[str]) -> object:
    """What Excel showed, as the engine holds it."""
    value = row[6].split("|")[0]
    try:
        return float(value)
    except ValueError:
        return value


@pytest.mark.excel
@pytest.mark.parametrize("row", HELD, ids=lambda r: f"{r[0]}:{r[1]}:{r[2]}")
def test_excel_keeps_the_cached_values_of_a_cycle(row):
    _variant, _action, _number, tag, _formula, cached, _shown, saved = row
    if tag == "control":
        return
    if cached == "None":
        assert shown(row) == ""
        assert "<v />" in saved
    else:
        value = ast.literal_eval(cached)
        assert str(shown(row)).removesuffix(".0") == value
        assert f"<v>{value}</v>" in saved


def book(tmp_path: Path, cells: list[tuple[str, str, str | None, str | None]], calc: str) -> Path:
    """A workbook of column A's rows as the capture has them, with the constants it reads."""
    xml = []
    for number, (_tag, formula, cached, kind) in enumerate(cells, 1):
        t = f' t="{kind}"' if kind else ""
        v = "<v/>" if cached is None else f"<v>{cached}</v>"
        constant = '<c r="B1"><v>1</v></c>' if number == 1 else ""
        constant = '<c r="B6"><v>2</v></c>' if number == 6 else constant
        xml.append(
            f'<row r="{number}"><c r="A{number}"{t}><f>{escape(formula)}</f>{v}</c>{constant}</row>'
        )
    path = tmp_path / "book.xlsx"
    path.write_bytes(package("".join(xml), workbook_extra=calc))
    return path


def capture_cells(variant: str) -> list[tuple[str, str, str | None, str | None]]:
    out = []
    for r in CACHED:
        if r[0] == variant and r[1] == "none":
            cached = ast.literal_eval(r[5])
            kind = "str" if isinstance(cached, str) and not cached.isdigit() else None
            out.append((r[3], r[4], cached, kind))
    return out


@pytest.mark.parametrize("variant", ["excel_both_flag", "excel_held_flag"])
@pytest.mark.parametrize("full", [False, True])
def test_the_engine_keeps_them_as_excel_does(tmp_path, variant, full):
    cells = capture_cells(variant)
    source = book(tmp_path, cells, '<calcPr fullCalcOnLoad="1"/>')
    out = tmp_path / "out.xlsx"
    report = xlrecalc.recalculate(source, out, full=full)
    wb = xlrecalc.load(source, trust_cached=not full)
    wb.recalculate()
    with zipfile.ZipFile(out) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    excel = [r for r in CACHED if r[0] == variant and r[1] == "full"]
    assert len(excel) == len(cells)
    for row in excel:
        number = int(row[2])
        assert wb.value(f"A{number}") == shown(row), number
        found = re.search(rf'<c r="A{number}"[^>]*>(.*?)</c>', sheet)
        assert found is not None
        cached = ast.literal_eval(row[5])
        if row[3] != "control" and cached is not None:
            assert f"<v>{cached}</v>" in found[1]
    assert len(report.circular) == len(cells) - 1
    assert not report.uncomputed


def _edited_engine(tmp_path) -> xlrecalc.Workbook:
    cells = capture_cells("excel_edited")
    if not cells:
        cells = [
            ("pair", "A2+B1", "901", None),
            ("pair", "A1+1", "902", None),
            ("dependent", "A1*2", "903", None),
            ("control", "B1*2", "904", None),
        ]
    return xlrecalc.load(book(tmp_path, cells, "<calcPr/>"))


@pytest.mark.parametrize(
    "action", ["set_input", "set_formula", "break", "new_cycle", "new_dependent"]
)
def test_an_edit_does_what_it_does_in_excel(tmp_path, action):
    wb = _edited_engine(tmp_path)
    wb.recalculate(full=True)  # the engine has met the cycle, as Excel has
    if action == "set_input":
        wb.set("B1", 5.0)
    elif action == "set_formula":
        wb.set_formula("A2", "=A1+2")
    elif action == "break":
        wb.set_formula("A2", "=7")
    elif action == "new_cycle":
        wb.set_formula("A5", "=A5+1")
    else:
        wb.set_formula("A6", "=A1*3")
    wb.recalculate()
    for row in [r for r in EDITED if r[1] == action]:
        if row[3] == "control" and action != "set_input":  # the full recalculation above reached it
            continue
        expected = shown(row)
        expected = None if expected == "" else expected
        assert wb.value(f"A{row[2]}") == expected, row


def test_a_new_formula_reading_a_cycle_no_evaluation_has_reached(tmp_path):
    """QUIRKS.md: a trusting load has not met the cycle, so this reads its cached value."""
    wb = _edited_engine(tmp_path)
    wb.set_formula("A6", "=A1*3")
    assert wb.value("A6") == 2703.0
    assert not wb.report.circular


def test_many_cycles_feeding_one_chain_walk_it_once(tmp_path):
    """A hundred thousand loops all read by one long chain: not a walk of the chain per loop."""
    cells: dict[str, object] = {f"A{r}": f"=A{r}+1" for r in range(1, 50_001)}
    cells["B1"] = "=SUM(A1:A50000)"
    cells.update({f"B{r}": f"=B{r - 1}+1" for r in range(2, 50_001)})
    engine = Engine.from_cells({"probe": cells})
    engine.recalculate()
    assert len(engine.report.circular) == 100_000
