"""The incremental refusal probe (``capture/refusals.py``) against a stand-in for Excel."""

from __future__ import annotations

from typing import NamedTuple

from capture.refusals import BATCH, Probe


class Case(NamedTuple):  # as capture.pipeline's, which needs openpyxl
    tag: str
    formula: str


CASES = [Case("t", f"F({i})") for i in range(150)]
REFUSES = {"F(3)", "F(70)", "F(71)", "F(149)"}


def CASES_ORDER(formula: str) -> int:  # noqa: N802
    return int(formula[2:-1])


def _probe() -> Probe:
    return Probe(lambda cases: not any(c.formula in REFUSES for c in cases))


def test_a_first_run_finds_every_refusal_by_bisection():
    probe = _probe()
    assert probe.run(CASES, set(), same_build=True, full=False) == sorted(REFUSES, key=CASES_ORDER)
    assert probe.opens > 3


def test_a_rerun_with_nothing_new_opens_one_workbook_a_batch():
    probe = _probe()
    found = probe.run(CASES, set(REFUSES), same_build=True, full=False)
    assert found == sorted(REFUSES, key=CASES_ORDER)
    assert probe.opens == -(-(len(CASES) - len(REFUSES)) // BATCH)


def test_known_refusals_are_probed_again_under_another_build_or_full():
    for same_build, full in ((False, False), (True, True)):
        probe = _probe()
        assert probe.run(CASES, set(REFUSES), same_build, full) == sorted(REFUSES, key=CASES_ORDER)
        assert probe.opens > -(-len(CASES) // BATCH)


def test_a_new_refusal_among_known_accepted_forms_is_news():
    probe = Probe(lambda cases: not any(c.formula in REFUSES | {"F(10)"} for c in cases))
    found = probe.run(CASES, set(REFUSES), same_build=True, full=False)
    assert set(found) == REFUSES | {"F(10)"}


ACCEPTED = {c.formula for c in CASES} - REFUSES


def test_known_accepted_forms_are_confirmed_in_one_open():
    probe = _probe()
    found = probe.run(CASES, set(REFUSES), True, False, ACCEPTED)
    assert found == sorted(REFUSES, key=CASES_ORDER)
    assert probe.opens == 1


def test_after_a_build_change_refusals_are_checked_by_entry_first():
    """Those still refused set in a cell are kept; only a changed verdict is opened in a file."""
    probe = Probe(_probe().open_cases, lambda formulas: set(formulas) - {"F(70)"})
    found = probe.run(CASES, set(REFUSES), False, False, ACCEPTED)
    assert found == sorted(REFUSES, key=CASES_ORDER)  # the file still refuses F(70)
    assert probe.opens == 2  # the accepted forms together, then F(70) alone
    now = Probe(lambda cases: not any(c.formula in REFUSES - {"F(70)"} for c in cases))
    now.entry = lambda formulas: set(formulas) - {"F(70)"}
    assert "F(70)" not in now.run(CASES, set(REFUSES), False, False, ACCEPTED)
