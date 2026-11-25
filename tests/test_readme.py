"""Every figure README.md states is derived here and checked, so a stale figure fails.

The evidence sentence (formulas, matches, known deviations) is checked against the capture run
in ``tests/test_captures.py``; the rest is here: the function counts (``docs/coverage.md``), the
largest group of known deviations and its size, the iterative functions' share of listed rows
(the deviation list and capture ``iterative``), and the relative tolerances (the registry).
"""

from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from tests.captures import load
from tests.test_captures import KNOWN_DEVIATIONS

ROOT = Path(__file__).parents[1]


def _readme() -> str:
    return " ".join((ROOT / "README.md").read_text(encoding="utf-8").split())


def _figure(text: str) -> int:
    return int(text.replace(",", ""))


def test_function_counts_are_the_coverage_pages():
    coverage = (ROOT / "docs" / "coverage.md").read_text(encoding="utf-8")
    counts = re.search(r"(\d+) implemented, (\d+) partial, (\d+) absent", coverage)
    assert counts, "docs/coverage.md: the totals line is missing"
    implemented, partial = int(counts[1]), int(counts[2])
    stated = re.search(r"([\d,]+) of Excel's functions \(([\d,]+) of them partly\)", _readme())
    assert stated, "README.md: the function count is missing"
    assert (_figure(stated[1]), _figure(stated[2])) == (implemented + partial, partial)


def test_largest_deviation_group_is_the_deviation_lists():
    groups = Counter(KNOWN_DEVIATIONS.values())
    (reason, size), (_, runner_up) = groups.most_common(2)
    assert size > runner_up, "two groups share the largest size: name both or neither"
    stated = re.search(r"The largest group of deviations is (.+?) \(([\d,]+) rows\)", _readme())
    assert stated, "README.md: the largest group of deviations is missing"
    assert _figure(stated[2]) == size
    names = set(re.findall(r"\b[A-Z][A-Z0-9.]{2,}\b", reason.split(":")[0]))
    assert names, reason
    assert all(f"`{name}`" in stated[1] for name in names), (stated[1], names)


def test_iterative_share_is_the_iterative_capture():
    rows = [r for r in load("iterative").rows if r.tag in ("irr", "xirr", "rate")]
    listed = sum(1 for r in rows if ("iterative", r.case) in KNOWN_DEVIATIONS)
    stated = re.search(
        r"([\d,]+) of the ([\d,]+) captured rows of `IRR`, `XIRR` and `RATE`", _readme()
    )
    assert stated, "README.md: the iterative functions' share is missing"
    assert (_figure(stated[1]), _figure(stated[2])) == (listed, len(rows))


def test_relative_tolerances_are_the_registry():
    """The relative tolerances stated: the iterative functions', the regression's, the continuous
    family's."""
    from xlrecalc._functions import SPECS, _continuous

    family = (_continuous._REL_TAILS, _continuous._REL_SEARCH, _continuous._REL_TEST)
    rels = [s.rel for s in SPECS.values() if s.rel_reason in family]
    others = {k: s.rel for k, s in SPECS.items() if s.rel and s.rel_reason not in family}
    text = " ".join(_readme().split())
    stated = re.search(
        r"a relative one: (.*?), ([\dE.-]+) to ([\dE.-]+) "
        r"for the beta, gamma, chi-squared, F and t distributions",
        text,
    )
    assert stated, "README.md: the relative tolerances are missing"
    groups = re.findall(r"([\dE.-]+) for ((?:`[A-Z]+`(?:, | and )?)+)", stated[1])
    named = {name: float(rel) for rel, names in groups for name in re.findall(r"`([A-Z]+)`", names)}
    assert named == others
    assert (float(stated[2]), float(stated[3])) == (min(rels), max(rels))
