"""The sizes ``CAPTURING.md`` §1 states for each capture, derived from the capture files.

A row of the table names one suite or several (``57 each``, or a total over the ``dyn_*``
suites) and a size: rows of a TSV, ``cells`` (the ``cell`` rows of a capture that records its
inputs beside them, otherwise its rows), or ``cases`` of the dynamic-array captures (one formula
at column B of its block). ``generated`` is not a size.
"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).parents[1]
CAPTURES = ROOT / "tests" / "captures"


def _table() -> list[tuple[list[str], str]]:
    text = (ROOT / "CAPTURING.md").read_text(encoding="utf-8")
    section = text.split("## 1. What exists", 1)[1].split("\n## ", 1)[0]
    out = []
    for line in section.splitlines():
        if line.startswith("| `"):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            out.append((re.findall(r"`(\w+)`", cells[0]), cells[1]))
    return out


def _lines(suite: str) -> list[list[str]]:
    text = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8")
    rows = [line.split("\t") for line in text.splitlines() if line and not line.startswith("#")]
    return rows[1:] if rows and rows[0][0] == "case" else rows  # a probe's column header


def _size(suite: str, unit: str) -> int:
    rows = _lines(suite)
    if unit == "cells" and any(r[0] == "input" for r in rows):  # inputs and cells interleaved
        return sum(1 for r in rows if r[0] == "cell")
    if unit == "cases":  # an anchor in column B: B1, B13, B25... (a legacy array: B109:B115)
        return sum(
            1
            for r in rows
            if r[0] == "input"
            and r[1] in ("dynamic", "legacy", "cse")
            and re.fullmatch(r"B\d+(?::[A-Z]+\d+)?", r[2])
        )
    return len(rows)


def test_every_suite_in_the_table_has_a_capture():
    for suites, _ in _table():
        for suite in suites:
            assert (CAPTURES / f"{suite}.tsv").exists(), suite


def test_stated_sizes_are_the_captures():
    table = _table()
    assert len(table) > 50
    for suites, stated in table:
        if stated == "generated":
            continue
        found = re.fullmatch(r"([\d,]+)(?: (cells|cases|each))?", stated)
        assert found, f"CAPTURING.md: {stated!r} for {suites}"
        n, unit = int(found[1].replace(",", "")), found[2] or "rows"
        if unit == "each":
            for suite in suites:
                assert _size(suite, "rows") == n, suite
        else:
            assert sum(_size(s, unit) for s in suites) == n, (suites, stated)


def test_a_suite_no_test_compares_says_so():
    """A suite that neither capture test runs and no other test names must say in its row that it
    is not compared, as evidence kept for a finding or for a later comparison.

    A test file that names a suite only in a docstring or comment counts as naming it: the check
    is by name, not by what the test does with the capture.
    """
    from tests.test_captures import SUITES as COMPARED
    from tests.test_dynamic_arrays import SUITES as DYNAMIC

    compared = set(COMPARED) | {f"dyn_{suite}" for suite in DYNAMIC}
    tests = [
        p.read_text(encoding="utf-8")
        for p in (ROOT / "tests").glob("*.py")
        if p.name != Path(__file__).name
    ]
    text = (ROOT / "CAPTURING.md").read_text(encoding="utf-8")
    section = text.split("## 1. What exists", 1)[1].split("\n## ", 1)[0]
    missing = []
    for line in section.splitlines():
        if not line.startswith("| `"):
            continue
        for suite in re.findall(r"`(\w+)`", line.split("|")[1]):
            named = any(re.search(rf"(?<!\w){suite}(?!\w)", t) for t in tests)
            if suite not in compared and not named and "Not compared" not in line:
                missing.append(suite)
    assert missing == []
