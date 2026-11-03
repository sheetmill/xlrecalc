"""Circular references with iteration off, as Excel leaves them (ARCHITECTURE.md §7.2).

Captures ``circular`` (alerts off) and ``circular_alerts`` (alerts on, the warning dismissed by a
person), from ``capture/circular_probe.py``: Excel calculates no cell on a cycle or depending on
one, and saves each as the empty text.
"""

import ast
from pathlib import Path

import pytest

from xlrecalc._engine import Engine

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
