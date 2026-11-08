"""The coverage matrix: every Excel worksheet function, its status here, and its evidence.

Joins the function registry, Excel's function list (``capture/functions_list.py``) and the
capture suites (``tests/captures``): for each function its category, status (implemented,
partial, absent), the note on what a partial one lacks, how many captured rows call it, how
many of those match Excel exactly, are within a declared tolerance or are listed deviations,
and its tolerance. A row counts for every function its formula calls.

    uv run python tools/coverage.py            # print the matrix
    uv run python tools/coverage.py --write    # write docs/coverage.md

``tests/test_coverage.py`` fails when ``docs/coverage.md`` is out of date. Run from the
repository root: the matrix needs the captures and the capture tests, which are not in the
package.
"""

from __future__ import annotations

import collections
import platform
import re
import sys
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from capture.functions_list import FUNCTIONS  # noqa: E402
from tests.captures import load  # noqa: E402
from tests.test_captures import (  # noqa: E402
    _CALL,
    KNOWN_DEVIATIONS,
    SUITES,
    relative,
    run,
    same,
    tolerance,
    within,
    within_relative,
)

from xlrecalc._functions import SPECS  # noqa: E402

OUTPUT = ROOT / "docs" / "coverage.md"
_PREFIXES = ("_XLFN._XLWS.", "_XLFN.")


def display(name: str) -> str:
    """A registry name as the formula bar shows it: no ``_xlfn.`` prefix."""
    upper = name.upper()
    for prefix in _PREFIXES:
        if upper.startswith(prefix):
            return upper[len(prefix) :]
    return upper


@dataclass
class Evidence:
    rows: int = 0
    exact: int = 0
    tolerated: int = 0
    relative: int = 0
    listed: int = 0


def evidence() -> dict[str, Evidence]:
    """Per function (formula-bar name), the captured rows that call it and how they fare."""
    found: dict[str, Evidence] = collections.defaultdict(Evidence)
    for name in SUITES:
        suite = load(name)
        engine, sheet = run(suite)
        for k, row in enumerate(suite.rows, 1):
            called = {display(n) for n in _CALL.findall(row.formula)}
            called &= {display(n) for n in SPECS}
            if not called:
                continue
            got = engine.get(sheet, k - 1, 0)
            if (name, row.case) in KNOWN_DEVIATIONS:  # first: a random row may match by chance
                outcome = "listed"
            elif same(got, row.expected):
                outcome = "exact"
            elif within(got, row.expected, tolerance(row.formula)):
                outcome = "tolerated"
            elif within_relative(got, row.expected, relative(row.formula)):
                outcome = "relative"
            else:
                outcome = ""
            for fn in called:
                e = found[fn]
                e.rows += 1
                if outcome:
                    setattr(e, outcome, getattr(e, outcome) + 1)
    _dynamic_evidence(found)
    return found


def _dynamic_evidence(found: dict[str, Evidence]) -> None:
    """The functions that make arrays (captures ``dyn_*``): a case is a row, exact when every
    cell of its result and its readers matches Excel."""
    from tests import test_dynamic_arrays as dyn

    for suite in dyn.SUITES:
        e = dyn.engine_for(suite)
        rich = dyn._rich_errors(suite)
        rows = dyn._rows(suite)
        cells = {r[1]: r for r in rows if r[0] == "cell"}
        for _, kind, where, text in (r for r in rows if r[0] == "input"):
            if kind not in ("dynamic", "legacy") or dyn._cell(where)[1] != 2:
                continue
            case = dyn._case(where)
            names = {n for n in _CALL.findall(text) if n.upper() in SPECS}  # not a bare new name
            called = {display(n) for n in names}
            if not called:
                continue
            mine = [a for a in cells if dyn._cell(a)[1] < 27 and dyn._case(a) == case]
            if any((suite, a) in dyn.KNOWN for a in mine):
                outcome = "listed"
            else:
                ok = True
                for a in mine:
                    r, c = dyn._cell(a)
                    want = dyn.expected(cells[a], rich)
                    got = e.get(0, r - 1, c - 1)
                    ok = ok and (got is dyn.BLANK if want is None else same(got, want))
                outcome = "exact" if ok else ""
            for fn in called:
                ev = found[fn]
                ev.rows += 1
                if outcome:
                    setattr(ev, outcome, getattr(ev, outcome) + 1)


def _declared(spec: object) -> str:
    ulps, rel = getattr(spec, "ulps", 0), getattr(spec, "rel", 0.0)
    return f"{ulps} ulps" if ulps else (f"{rel:g} relative" if rel else "")


def platform_label() -> str:
    """The platform the split between exact and ulps rows is measured on, as the page records it.

    The math library's last bits come with the operating system: on macOS, its major version
    and the machine's architecture (``macOS 26, arm64``).
    """
    machine = platform.machine()
    if sys.platform == "darwin":
        return f"macOS {platform.mac_ver()[0].split('.')[0]}, {machine}"
    return f"{platform.system()}, {machine}"


def recorded_platform(page: str) -> str | None:
    """The platform a generated page records, or ``None``."""
    found = re.search(r"Generated by `tools/coverage\.py` on (.+?) \(", " ".join(page.split()))
    return found[1] if found else None


def matrix() -> str:
    """The matrix as a Markdown page."""
    specs = {display(n): s for n, s in SPECS.items()}
    names = sorted(set(FUNCTIONS) | set(specs))
    found = evidence()
    counts = collections.Counter(specs[n].status if n in specs else "absent" for n in names)
    lines = [
        "# Function coverage",
        "",
        "Every worksheet function Excel has, whether this package computes it, and the",
        "evidence: captured rows that call the function (a row counts for every function in",
        "its formula), how many match Excel exactly, are within the function's declared",
        "tolerance in units in the last place (ulps) or its relative tolerance (rel, for the",
        "iterative functions), or are listed deviations with their reasons",
        "(`QUIRKS.md`, `tests/test_captures.py`).",
        "",
        f"Generated by `tools/coverage.py` on {platform_label()} (the reference platform is",
        "macOS): the split between exact and ulps can differ on another platform, whose math",
        "library differs in the last bits; the rows that match either way do not.",
        "",
        f"{counts['implemented']} implemented, {counts['partial']} partial, "
        f"{counts['absent']} absent.",
        "",
        "| function | category | status | rows | exact | ulps | rel | listed | tolerance | note |",
        "|---|---|---|---:|---:|---:|---:|---:|---|---|",
    ]
    for n in names:
        spec = specs.get(n)
        e = found.get(n, Evidence())
        if spec is None:
            lines.append(f"| `{n}` | | absent | | | | | | | |")
            continue
        lines.append(
            f"| `{n}` | {spec.category} | {spec.status} | {e.rows} | {e.exact} | "
            f"{e.tolerated} | {e.relative} | {e.listed} | {_declared(spec)} | {spec.note} |"
        )
    return "\n".join(lines) + "\n"


def main() -> None:
    text = matrix()
    if "--write" in sys.argv[1:]:
        OUTPUT.write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()
