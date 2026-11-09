"""Counts QUIRKS.md states, derived from the captures and the deviation lists and checked.

Six groups: the sizes of suites given as ``suite`` (N rows), the CUMIPMT and CUMPRINC rows with
the group's total (the README's largest group, ``tests/test_readme.py``) and the TBILLEQ row,
VDB's listed share of its sweeps, the MULTINOMIAL rows, the iterative rows, and the
empty-argument rows. Other figures there (character pairs, units in the last place, dates) are
not checked here.
"""

from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from tests import test_captures as tc
from tests.captures import load

ROOT = Path(__file__).parents[1]
WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6}
# The VDB sweeps of fractional periods are the rows of capture finance from this case on.
_VDB_SWEEP_FIRST = 932


def _quirks() -> str:
    return " ".join((ROOT / "QUIRKS.md").read_text(encoding="utf-8").split())


def _n(text: str) -> int:
    return WORDS.get(text.lower()) or int(text.replace(",", ""))


def _listed(reason: str, suite: str | None = None) -> int:
    return sum(1 for (s, _), r in tc.KNOWN_DEVIATIONS.items() if r == reason and suite in (None, s))


def _one(pattern: str) -> re.Match[str]:
    found = re.search(pattern, _quirks())
    assert found, f"QUIRKS.md: no match for {pattern!r}"
    return found


def test_suite_sizes():
    stated = re.findall(r"`(\w+)` \(([\d,]+) rows[):]", _quirks())
    assert len(stated) >= 5
    for suite, n in stated:
        assert len(load(suite).rows) == _n(n), suite


def test_cumulative_noise_and_tbilleq():
    found = _one(
        r"([\d,]+) rows of `CUMIPMT` and `CUMPRINC` are further.*?with one more in `empty_args`, "
        r"([\d,]+) rows in all.*?(\w+) row of `TBILLEQ`"
    )
    assert _n(found[1]) == _listed(tc._CUMULATIVE_NOISE, "finance_family")
    assert _n(found[2]) == _listed(tc._CUMULATIVE_NOISE)
    assert _n(found[2]) - _n(found[1]) == _listed(tc._CUMULATIVE_NOISE, "empty_args") == 1
    assert _n(found[3]) == _listed(tc._TBILLEQ_NOISE)
    largest, size = Counter(tc.KNOWN_DEVIATIONS.values()).most_common(1)[0]
    assert largest == tc._CUMULATIVE_NOISE  # the README's largest group, by the same count
    assert size == _n(found[2])


def test_vdb_listed_share():
    found = _one(r"in its last bits on ([\d,]+) of the ([\d,]+) rows of `finance`")
    sweep = [r for r in load("finance").rows if r.case >= _VDB_SWEEP_FIRST]
    assert all(r.tag == "vdb" for r in sweep)
    assert (_n(found[1]), _n(found[2])) == (_listed(tc._VDB_FRACTIONAL), len(sweep))


def test_multinomial_rows():
    found = _one(
        r"\(279\.99999999999966 for 280\); ([\d,]+) rows of `math_integers` listed.*?"
        r"(\w+) of `statistics_family` and (\w+) of `empty_args`, ([\d,]+) in all"
    )
    by_suite = [_listed(tc._MULTINOMIAL, s) for s in ("math_integers", "statistics_family")]
    by_suite.append(_listed(tc._MULTINOMIAL, "empty_args"))
    assert [_n(found[k]) for k in (1, 2, 3)] == by_suite
    assert _n(found[4]) == _listed(tc._MULTINOMIAL) == sum(by_suite)


def test_iterative_rows():
    found = _one(r"Of the ([\d,]+) rows of `iterative`, ([\d,]+) match.*?and ([\d,]+) are listed")
    rows = load("iterative").rows
    listed = sum(1 for r in rows if ("iterative", r.case) in tc.KNOWN_DEVIATIONS)
    assert (_n(found[1]), _n(found[2]), _n(found[3])) == (len(rows), len(rows) - listed, listed)


def test_empty_args_rows():
    found = _one(
        r"`VDB` with an empty `factor` is `#DIV/0!` in Excel and 0 here \((\w+) rows of "
        r"`empty_args`\)\. (\w+) others there carry the known noise"
    )
    noise = (tc._MULTINOMIAL, tc._CUMULATIVE_NOISE, tc._XIRR)
    assert _n(found[1]) == _listed(tc._EMPTY_ODD, "empty_args")
    assert _n(found[2]) == sum(_listed(r, "empty_args") for r in noise)
