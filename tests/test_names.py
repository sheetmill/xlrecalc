"""Defined names as references, against Excel (capture ``names``, ``capture/name_probe.py``).

A name for one cell is a reference, as a name for a range is, so ``OFFSET``, ``ROW`` and the range
operator take it. A name for anything else is a value: ``ROW(Number)`` is ``#VALUE!``.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from xlrecalc._engine import Engine
from xlrecalc._values import ErrorValue

LINES = (
    (Path(__file__).with_name("captures") / "names.tsv").read_text(encoding="utf-8").splitlines()
)
NAMES: dict[str, str] = ast.literal_eval(
    next(line for line in LINES if line.startswith("# names: ")).removeprefix("# names: ")
)
ROWS = [line.split("\t") for line in LINES if not line.startswith("#")]


@pytest.fixture(scope="module")
def engine() -> Engine:
    sheet: dict[str, object] = {}
    for r in range(1, 6):
        sheet[f"Z{r}"] = float(r)
        sheet[f"AA{r}"] = float(10 * r)
    for number, formula, *_ in ROWS:
        where = f"A{number}" if number.isdigit() else number
        sheet[where] = "=" + ast.literal_eval(formula)
    other = {f"A{r}": float(100 * r) for r in range(1, 4)}
    for c, value in zip("CDE", (11.0, 12.0, 13.0), strict=True):
        other[f"{c}5"] = value
        other[f"{c}7"] = value * 2
        other[f"{c}8"] = value * 3
    for r, value in enumerate((21.0, 22.0, 23.0), 1):
        other[f"G{r}"] = value
        sheet[f"AC{r}"] = value * 2
    e = Engine.from_cells({"S": sheet, "Other": other})
    for name, text in NAMES.items():
        e.define_name(name, text)
    e.recalculate()
    return e


@pytest.mark.excel
@pytest.mark.parametrize("row", ROWS, ids=lambda r: f"{r[0]}:{ast.literal_eval(r[1])}")
def test_matches_excel(engine: Engine, row: list[str]) -> None:
    number, _formula, kind, text = row
    value = ast.literal_eval(text)
    if number.isdigit():
        got = engine.get(0, int(number) - 1, 0)
    else:
        got = engine.get(0, int(number[1:]) - 1, ord(number[0]) - 65)
    if kind == "e":
        assert got == ErrorValue(value)
    else:
        assert got == value
        assert type(got) is type(value)
