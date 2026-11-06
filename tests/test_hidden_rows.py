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
