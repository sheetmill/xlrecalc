"""What Excel's repair does to formulas it refuses in a file (ARCHITECTURE.md §9.5)."""

import ast
from pathlib import Path

import pytest

CAPTURE = Path(__file__).with_name("captures") / "repair.tsv"
ROWS = [
    line.split("\t")
    for line in CAPTURE.read_text(encoding="utf-8").splitlines()
    if not line.startswith("#")
]


@pytest.mark.excel
@pytest.mark.parametrize("row", ROWS, ids=lambda r: f"{r[0]}:{ast.literal_eval(r[2])!r}")
def test_repair(row):
    number, tag, _written, formula_after, value_after, _text, saved = row
    cached = 100 + int(number)
    formula_after, value_after, saved = map(ast.literal_eval, (formula_after, value_after, saved))
    if tag == "refused":
        # The formula is dropped; its cached value stays as a constant.
        assert formula_after == str(cached)
        assert float(value_after) == cached
        assert "<f>" not in saved
    else:
        # A control formula survives and is recalculated.
        assert formula_after.startswith("=")
        assert "<f>" in saved
        assert value_after != f"{cached}.0"


UNCACHED = [
    line.split("\t")
    for line in (CAPTURE.with_name("repair_uncached.tsv")).read_text(encoding="utf-8").splitlines()
    if not line.startswith("#")
]


@pytest.mark.excel
@pytest.mark.parametrize("row", UNCACHED, ids=lambda r: f"{r[0]}:{ast.literal_eval(r[2])!r}")
def test_repair_without_a_cached_value(row):
    """A refused formula without a cached value becomes a blank cell, as the reader makes it."""
    _number, tag, _written, formula_after, _value, _text, saved, saved_c = row
    formula_after, saved, saved_c = map(ast.literal_eval, (formula_after, saved, saved_c))
    if tag == "refused":
        assert formula_after == ""
        assert saved == ""  # no <c> at all
        assert saved_c.endswith("<v>1</v></c>")  # ISBLANK is TRUE
    else:
        assert formula_after.startswith("=")
        assert saved_c.endswith("<v>0</v></c>")
