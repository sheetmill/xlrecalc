"""SUBTOTAL over hidden rows, against Excel (capture ``hidden_rows``)."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

import xlrecalc

ROWS = [
    line.split("\t")
    for line in (Path(__file__).with_name("captures") / "hidden_rows.tsv")
    .read_text(encoding="utf-8")
    .splitlines()
    if not line.startswith("#")
]


@pytest.fixture(scope="module")
def workbook(tmp_path_factory: pytest.TempPathFactory) -> xlrecalc.Workbook:
    pytest.importorskip("openpyxl")
    from capture.hidden_probe import build

    path = tmp_path_factory.mktemp("hidden") / "hidden.xlsx"
    build(path)
    return xlrecalc.load(path)


@pytest.mark.excel
@pytest.mark.parametrize("row", ROWS, ids=lambda r: f"{r[0]}:{ast.literal_eval(r[1])}")
def test_matches_excel(workbook: xlrecalc.Workbook, row: list[str]) -> None:
    number, _formula, _kind, text = row
    assert workbook.value(f"R!A{number}") == ast.literal_eval(text)


def test_filter_mode_read_where_rows_are_hidden(workbook: xlrecalc.Workbook) -> None:
    """``filterMode`` is read on a sheet with hidden rows, the only case where it matters."""
    engine = workbook._engine
    modes = {s.name: s.filter_mode for s in engine.sheets}
    assert modes["Filtered"]
    assert modes["Both"]
    assert not modes["Hand"]
    assert all(engine.sheets[i].hidden_rows for i in range(3))


def test_an_empty_row_is_not_hidden() -> None:
    """A ``<row>`` with no cells and no formats, which xlio reports, hides nothing."""
    import io

    from tests.xlsx_files import package
    from xlrecalc._load import load

    rows = (
        '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'
        '<row r="3"/><row r="4"></row><row r="5" hidden="1"><c r="A5"><v>4</v></c></row>'
        '<row r="6" hidden="0"/><row r="7"><c r="A7"><f>SUBTOTAL(109,A1:A6)</f></c></row>'
    )
    engine = load(io.BytesIO(package(rows)))
    assert engine.sheets[0].hidden_rows == frozenset({4})
    engine.recalculate(full=True)
    assert engine.get(0, 6, 0) == 3.0  # row 5 left out, the empty rows count nothing
