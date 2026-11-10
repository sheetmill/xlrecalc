"""The incremental refusal probe (``capture/refusals.py``) against a stand-in for Excel."""

from __future__ import annotations

from capture.pipeline import Case
from capture.refusals import BATCH, Probe

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
