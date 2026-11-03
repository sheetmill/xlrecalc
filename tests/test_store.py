"""The store's template ids against a plain dict, through dense, growing and sparse columns."""

import random

import pytest

from xlrecalc._store import COLS, MAX_ROW, TemplateIds


@pytest.mark.parametrize("seed", range(20))
def test_template_ids_match_a_dict(seed):
    rng = random.Random(seed)
    ids, reference = TemplateIds(), {}
    for _ in range(3000):
        col = rng.randrange(4)
        kind = rng.random()
        if kind < 0.6:  # mostly in order down a column, as a file is read
            row = max((k // COLS for k in reference if k % COLS == col), default=-1) + 1
        elif kind < 0.8:
            row = rng.randrange(200)  # inside or above the span: fill and prepend
        else:
            row = rng.randrange(MAX_ROW + 1)  # far away: the column becomes sparse
        key = row * COLS + col
        if rng.random() < 0.15:
            assert ids.remove(row, col) == (reference.pop(key, None) is not None)
        else:
            tid = rng.randrange(50)
            assert ids.put(row, col, tid) == (key not in reference)
            reference[key] = tid
        assert len(ids) == len(reference)
    assert sorted(ids) == sorted(reference)
    for key, tid in reference.items():
        assert ids[key] == tid
        assert key in ids
    for _ in range(500):
        key = rng.randrange(MAX_ROW + 1) * COLS + rng.randrange(5)
        assert ids.get(key) == reference.get(key)
        assert (key in ids) == (key in reference)
    with pytest.raises(KeyError):
        ids[MAX_ROW * COLS + 9]


def test_a_column_spanning_the_sheet_becomes_sparse():
    ids = TemplateIds()
    ids.put(0, 2, 1)
    ids.put(MAX_ROW, 2, 3)
    assert ids.at(0, 2) == 1
    assert ids.at(MAX_ROW, 2) == 3
    assert 2 in ids._sparse
    assert 2 not in ids._dense
