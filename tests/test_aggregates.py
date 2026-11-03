"""SUM over one reference (a single loop) against the general path, value for value."""

import random

import pytest

from xlrecalc._coerce import numbers
from xlrecalc._engine import Engine
from xlrecalc._functions._math import _total
from xlrecalc._quirks import Quirks
from xlrecalc._values import ErrorValue


def _value(rng: random.Random) -> object:
    kind = rng.random()
    if kind < 0.55:
        magnitude = 10 ** rng.randint(-20, 20)
        return rng.choice([1, -1]) * rng.random() * magnitude
    if kind < 0.65:
        return rng.choice([0.1, 0.2, 0.3, -0.3, 1e16, -1e16, 1.0, -1.0])  # cancellation
    if kind < 0.75:
        return rng.choice(["text", "1", True, False])
    if kind < 0.78:
        return ErrorValue.NA
    return None  # blank


@pytest.mark.parametrize("snapping", [True, False])
@pytest.mark.parametrize("seed", range(30))
def test_sum_of_a_reference_matches_the_general_path(seed, snapping):
    rng = random.Random(seed)
    quirks = Quirks(cancellation_snapping=snapping)
    cells = {f"A{r}": _value(rng) for r in range(1, 41)}
    cells["B1"] = "=SUM(A1:A40)"
    engine = Engine.from_cells({"S": cells}, quirks=quirks)
    engine.recalculate()
    ref = engine.box(None, 0, 0, 39, 0)
    expected = _total(numbers((ref,), engine), engine)
    got = engine.get(0, 0, 1)
    assert type(got) is type(expected)
    assert got == expected or (got != got and expected != expected)
