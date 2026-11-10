"""Counts QUIRKS.md states, derived from the captures and the deviation lists and checked.

Six groups: the sizes of suites given as ``suite`` (N rows), the CUMIPMT and CUMPRINC rows with
the group's total (the README's largest group, ``tests/test_readme.py``) and the TBILLEQ row,
VDB's listed share of its sweeps, the MULTINOMIAL rows, the iterative rows, and the
empty-argument rows. Other figures there (character pairs, units in the last place, dates) are
not checked here.
"""

from __future__ import annotations

import math
import re
from collections import Counter
from pathlib import Path

from tests import test_captures as tc
from tests.captures import load

ROOT = Path(__file__).parents[1]
WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7}
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


def test_distribution_rows():
    found = _one(
        r"([\d,]+) rows of the discrete ones are further.*?\(the largest group of known "
        r"deviations\), and ([\d,]+) of the others"
    )
    assert _n(found[1]) == _listed(tc._DISCRETE_TERMS)
    assert _n(found[2]) == _listed(tc._DISTRIBUTION_LIBRARY)
    largest, size = Counter(tc.KNOWN_DEVIATIONS.values()).most_common(1)[0]
    assert largest == tc._DISCRETE_TERMS  # the README's largest group, by the same count
    assert size == _n(found[1])


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


def _continuous_spread() -> dict[str, float]:
    """Per function, the largest relative difference from Excel on its rows not listed."""
    outer = re.compile(r"^(?:_xlfn\.)?([A-Z][A-Z0-9.]*)\(")
    suite = load("continuous_family")
    engine, sheet = tc.run(suite)
    spread: dict[str, float] = {}
    for k, row in enumerate(suite.rows):
        found, got, want = outer.match(row.formula), engine.get(sheet, k, 0), row.expected
        if ("continuous_family", row.case) in tc.KNOWN_DEVIATIONS or not found:
            continue
        if isinstance(got, float) and isinstance(want, float) and want and not tc.same(got, want):
            spread[found[1]] = max(spread.get(found[1], 0.0), abs(got - want) / abs(want))
    return spread


def test_continuous_tolerances():
    from xlrecalc._functions import SPECS, _continuous

    family = (_continuous._REL_TAILS, _continuous._REL_SEARCH, _continuous._REL_TEST)
    declared = {k.removeprefix("_XLFN."): s.rel for k, s in SPECS.items() if s.rel_reason in family}
    entry = _one(r"Declared tolerances, `continuous_family`.*?rows are listed")[0]
    stated: dict[str, tuple[float, str]] = {}
    for tol, names in re.findall(r"(1E-\d+): ((?:`[A-Z.0-9]+` \([\d.E-]+\)(?:, )?)+)", entry):
        for name, largest in re.findall(r"`([A-Z.0-9]+)` \(([\d.E-]+)\)", names):
            stated[name] = (float(tol), largest)
    assert {n: tol for n, (tol, _) in stated.items()} == declared
    if tc._reference_platform():
        spread = _continuous_spread()
        assert {n: largest for n, (_, largest) in stated.items()} == {
            n: f"{spread[n]:.1E}" for n in stated
        }


def test_continuous_listed_rows():
    groups = [
        (tc._SEARCH_TINY_P, r"probability of 1E-300, where Excel's search"),
        (tc._SEARCH_HUGE_DF, r"`CONFIDENCE.T` at 1E10 degrees of freedom"),
        (tc._SEARCH_1E6, r"`CONFIDENCE.T` at 1E6 degrees of freedom"),
        (tc._T_CENTER, r"`T.INV.2T` near 1 at one degree of freedom"),
        (tc._T_HALF, r"the t distribution at 1E-10, exactly 0.5"),
        (tc._EXTREME_NUM, r"`#NUM!` where the result exists"),
        (tc._EXTREME_SEARCH, r"or a search that stops short"),
        (tc._CHISQ_ONE_DF, r"`CHIINV` at one degree of freedom invert"),
        (tc._BOTH_HUGE, r"`BETA.DIST` with both shapes 1E300"),
    ]
    found = _one(r"([\d,]+) rows are listed\. In ([\d,]+) of them Excel's result")
    listed = [_listed(reason, "continuous_family") for reason, _ in groups]
    assert _n(found[1]) == sum(listed) == _listed_suite("continuous_family")
    ours = _listed(tc._CHISQ_ONE_DF) + _listed(tc._BOTH_HUGE)
    assert _n(found[2]) == sum(listed) - ours
    for (_, phrase), n in zip(groups, listed, strict=True):
        assert _n(_one(re.escape(phrase).replace(r"\ ", " ") + r".*? \(([\d,]+) rows?\)")[1]) == n


def test_continuous_cases_excel_does_not_finish():
    """The formulas past the time limit and the others timed, from tests/captures/slow_cases.tsv."""
    lines = (
        (ROOT / "tests" / "captures" / "slow_cases.tsv").read_text(encoding="utf-8").splitlines()
    )
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    past = [r for r in rows if r[4] == "past"]
    timed = [float(r[6]) for r in rows if r[4] != "past"]
    found = _one(
        r"(\w+) formulas at a shape of 1\.7E308 .*? past 90 seconds, alone in a workbook, "
        r"where the other ([\d,]+) timed.*? take at most ([\d.]+) seconds"
    )
    assert _n(found[1].lower()) == len(past)
    assert all("1.7e+308" in r[3] for r in past)
    assert _n(found[2]) == len(timed)
    assert math.ceil(max(timed)) == float(found[3])


def _listed_suite(suite: str) -> int:
    return sum(1 for s, _ in tc.KNOWN_DEVIATIONS if s == suite)
