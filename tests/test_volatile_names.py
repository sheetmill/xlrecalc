"""What opening recomputes, against capture ``volatile_names`` (Excel 16.115 opening each variant).

Each row is a workbook Excel saved, its cached values then planted wrong (A2 999, A3 777, A4 555;
A1 = 2, A2 = A1*3, A3 the formula under test, A4 = A3+1). Here the same parts are written from
the capture (``<fileVersion>``, ``calcPr``, the defined names and A2:A4 as planted), and
``load``, ``recalculate`` and ``RecalcCheck`` are held to what Excel showed on opening and saved.
"""

from __future__ import annotations

import ast
import io
from pathlib import Path

import pytest
import xlio

import xlrecalc
from tests.xlsx_files import package
from xlrecalc._engine import Engine
from xlrecalc._formula import FormulaSyntaxError, TokenizeError, compile_formula, parse
from xlrecalc._functions import REGISTRY
from xlrecalc._opening import VOLATILE_FUNCTIONS, Volatility

CAPTURE = Path(__file__).with_name("captures") / "volatile_names.tsv"
REFS = ("A2", "A3", "A4")
# A cell volatile only through a defined name, without Excel's ca="1" (a file another program
# wrote): Excel does not count it volatile until it is entered again, so its save keeps A3; a save
# here computes it (QUIRKS.md).
SAVE_DEVIATIONS = {"offset_name_no_ca"}
PLANTED = {"A2": 999.0, "A3": 777.0, "A4": 555.0}


def _rows() -> list[list[str]]:
    lines = CAPTURE.read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith("#")]


def _workbook(row: list[str]) -> bytes:
    names = row[3] if row[3] != "-" else ""
    extra = row[4] + (f"<definedNames>{names}</definedNames>" if names else "")
    cells = '<row r="1"><c r="A1"><v>2</v></c></row>'
    for i, cell in enumerate(row[5].split("</c> "), 2):
        cells += f'<row r="{i}">{cell if cell.endswith("</c>") else cell + "</c>"}</row>'
    return package(
        cells, workbook_extra=extra, sheets='<sheet name="probe" sheetId="1" r:id="rId1"/>'
    )


def _numbers(cells: str) -> dict[str, float]:
    """``A1=2 A2=999 A3=6 A4=7`` as numbers."""
    return {k: float(v) for k, _, v in (item.partition("=") for item in cells.split())}


def _shown(row: list[str]) -> dict[str, float]:
    shown = ast.literal_eval("[" + row[7].replace("' '", "', '") + "]")
    return dict(zip(REFS, map(float, shown), strict=True))


def _answer(values: dict[str, float]) -> str:
    """What a check answers when a load gives these values: A2 recomputed is every formula."""
    if values["A2"] != PLANTED["A2"]:
        return "all"
    return "some" if any(values[r] != PLANTED[r] for r in REFS) else "none"


def _checked(data: bytes, **options: object) -> str:
    check = xlrecalc.RecalcCheck(io.BytesIO(data), **options)  # type: ignore[arg-type]
    with xlio.open(io.BytesIO(data)) as book:
        for cell in book.sheets[0].cells():
            f = cell.formula
            if f is not None:
                ca = f.attributes.get("ca") in ("1", "true")
                check.add(f.text, cached=cell.value is not None, ca=ca)
    return check.finish()


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_a_load_computes_what_excel_recalculates_on_opening(row: list[str]) -> None:
    wb = xlrecalc.load(io.BytesIO(_workbook(row)))
    got = {ref: wb.value(ref) for ref in REFS}
    assert got == _shown(row), row[0]


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_recalculate_writes_what_excel_saves_after_opening(tmp_path: Path, row: list[str]) -> None:
    path = tmp_path / "book.xlsx"
    path.write_bytes(_workbook(row))
    xlrecalc.recalculate(path, full_calc_on_load=False)
    wb = xlrecalc.load(path, recalc_on_load="never")
    saved = _numbers(row[8])
    same = {ref: wb.value(ref) for ref in REFS} == {ref: saved[ref] for ref in REFS}
    assert same is (row[0] not in SAVE_DEVIATIONS), row[0]


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_the_check_answers_as_the_load_computes(row: list[str]) -> None:
    data = _workbook(row)
    assert _checked(data) == _answer(_shown(row)), row[0]
    assert _checked(data, saving=True) == _answer(_numbers(row[8])), row[0]


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_never_computes_the_volatile_formulas_always_every_one(row: list[str]) -> None:
    """``"never"`` computes a volatile formula, through a name too (as at every calculation)."""
    data = _workbook(row)
    never = xlrecalc.load(io.BytesIO(data), recalc_on_load="never")
    volatile = row[0].split("_no_ca")[0] not in ("plain_name", "unused_offset_name", "literal")
    volatile = volatile and "plain" not in row[0]
    assert never.value("A3") == (6.0 if volatile else 777.0), row[0]
    assert _checked(data, recalc_on_load="never") == ("some" if volatile else "none")
    always = xlrecalc.load(io.BytesIO(data), recalc_on_load="always")
    assert always.value("A2") == 6.0
    assert _checked(data, recalc_on_load="always") == "all"


NAMES = [
    ("PLAIN", "probe!$A$1"),
    ("OFF", "OFFSET(probe!$A$1,0,0)"),
    ("IND", 'INDIRECT("probe!A1")'),
    ("CHAIN", "OFF+0"),
    ("LOOP1", "LOOP2+1"),
    ("LOOP2", "LOOP1+RAND()"),
    ("NOWF", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+NOW()*0)"),
    ("PLAINF", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+0)"),
    ("N", "OFFSET(probe!$A$1,0,0)"),
]


@pytest.mark.parametrize(
    ("text", "volatile", "own"),
    [
        ("PLAIN*3", False, False),
        ("OFF*3", True, False),
        ("IND*3", True, False),
        ("CHAIN*3", True, False),
        ("LOOP1", True, False),
        ("NOWF(A1)*3", True, False),
        ("PLAINF(A1)*3", False, False),
        ('LEN("RAND()")+A1', False, False),
        ('"OFF"&A1', False, False),
        ("OFFSET(A1,0,0)", True, True),
        ("_xlfn.RANDARRAY(2)", True, True),
        ("rand()", True, True),
        ("N(A1)", False, False),  # the function N, not the name
        ("[1]!OFF", False, False),  # another workbook's name
        ("probe!OFF+1", True, False),
        ("SUM(A1:A3)", False, False),
    ],
)
def test_the_detector(text: str, volatile: bool, own: bool) -> None:
    v = Volatility(NAMES)
    assert v(text) is volatile
    assert v.own(text) is own


def test_own_volatility_is_the_compiler_s() -> None:
    """``own`` agrees with the functions the compiler records as called, over the corpus."""
    from tests.formula_corpus import CORPUS

    engine, v = Engine(), Volatility([])
    for text in CORPUS:
        try:
            compiled = compile_formula(parse(text), 0, 0, REGISTRY, engine.ops, engine.quirks)
        except (FormulaSyntaxError, TokenizeError):  # texts the parser refuses call nothing
            continue
        assert v.own(text) is bool(compiled.calls & VOLATILE_FUNCTIONS), text


def test_result_before_finish_and_add_after() -> None:
    data = package(
        '<row r="1"><c r="A1"><f>RAND()</f><v>0.5</v></c></row>',
        workbook_extra='<fileVersion appName="xl"/><calcPr calcId="191029"/>',
    )
    check = xlrecalc.RecalcCheck(io.BytesIO(data))
    assert check.result is None
    check.add("A1+1", cached=False)
    assert check.result is None  # "some" so far: a later formula could still make it "all"
    check.add("RAND()")  # a volatile function, not marked
    assert check.result == "all"
    assert check.finish() == "all"
    with pytest.raises(ValueError, match="after finish"):
        check.add("A1")
    old = package("", workbook_extra='<calcPr calcId="152511"/>')
    assert xlrecalc.RecalcCheck(io.BytesIO(old)).result == "all"  # decided by the file alone
    with pytest.raises(ValueError, match="recalc_on_load"):
        xlrecalc.RecalcCheck(io.BytesIO(old), recalc_on_load="sometimes")  # type: ignore[arg-type]


def test_manual_calculation_load_keeps_and_recalculate_computes(tmp_path: Path) -> None:
    """Manual calculation: ``load`` keeps every cached value, as Excel shows them on opening;
    ``recalculate``, which saves, writes what Excel's save recalculates, the cells marked
    ``ca="1"`` (capture ``volatile_names``, ``manual_rand_in_cell``)."""
    row = next(r for r in _rows() if r[0] == "manual_rand_in_cell")
    assert 'calcMode="manual"' in row[4]
    assert 'ca="1"' in row[5]
    data = _workbook(row)
    loaded = xlrecalc.load(io.BytesIO(data))
    assert {ref: loaded.value(ref) for ref in REFS} == PLANTED
    path = tmp_path / "manual.xlsx"
    path.write_bytes(data)
    xlrecalc.recalculate(path, full_calc_on_load=False)
    with xlio.open(path) as book:  # the cached values as written, not computed again
        written = {f"A{c.row}": c.value for c in book.sheets[0].cells() if c.formula is not None}
    assert written == {"A2": 999.0, "A3": 6.0, "A4": 7.0}
    assert _numbers(row[8]) == {"A1": 2.0, "A2": 999.0, "A3": 6.0, "A4": 7.0}  # Excel's saved copy
    assert _checked(data) == "none"
    assert _checked(data, saving=True) == "some"
