"""Which formulas Excel refuses in a file, probed incrementally (``CAPTURING.md`` §2).

Excel refuses a whole file for one formula it does not take (-50), so a probe opens formulas in
batches and halves a refused batch until each refusal is opened alone. Most answers are settled
already, by the build that gave them: a refused list (``# excel:`` in its header) and the suite
captured without its refusals. So a run keeps each known refusal of the same build without an
open, and opens everything else in full batches: the forms known to be accepted, which confirm
in one open a batch when nothing changed (a refusal there is news, and is bisected), the forms
not seen before, and the refusals of another build. A rerun with nothing new costs one open per
batch.
"""

from __future__ import annotations

import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path

from capture.pipeline import Case

BATCH = 60
MAX_OPENS = 2500


def _build(lines: Iterable[str]) -> str | None:
    for line in lines:
        if line.startswith("# excel: "):
            return line.removeprefix("# excel: ").split(" for Mac", 1)[0]
    return None


def known_refused(path: Path) -> tuple[set[str], str | None]:
    """The refused formulas a list holds, and the build that refused them (``None``: unstated)."""
    if not path.exists():
        return set(), None
    lines = path.read_text(encoding="utf-8").split("\n")
    return {line for line in lines if line and not line.startswith("# ")}, _build(lines)


def known_accepted(path: Path) -> set[str]:
    """The formulas a captured suite holds: each opened in a file Excel took."""
    if not path.exists():
        return set()
    out = set()
    for line in path.read_text(encoding="utf-8").split("\n"):
        if line and not line.startswith("#"):
            out.add(line.split("\t")[3])
    return out


@dataclass
class Probe:
    """One run: what it opened, how long the opens took, and the refusals it holds."""

    open_cases: Callable[[list[Case]], bool]
    opens: int = 0
    seconds: float = 0.0
    refused: list[str] = field(default_factory=list)

    def opened(self, cases: list[Case]) -> bool:
        """Whether Excel opened a workbook of these cases; one open, timed."""
        self.opens += 1
        if self.opens > MAX_OPENS:
            raise SystemExit(f"more than {MAX_OPENS} opens; so far {len(self.refused)} refused")
        start = time.monotonic()
        try:
            return self.open_cases(cases)
        finally:
            self.seconds += time.monotonic() - start

    def walk(self, cases: list[Case]) -> None:
        """Open the cases together, and halve a refused batch until each refusal is alone."""
        if not cases or self.opened(cases):
            return
        if len(cases) == 1:
            self.refused.append(cases[0].formula)
            return
        middle = len(cases) // 2
        self.walk(cases[:middle])
        self.walk(cases[middle:])

    def run(self, cases: list[Case], refused: set[str], same_build: bool, full: bool) -> list[str]:
        """The refused formulas among ``cases``, opening only what is not settled."""
        settled = set() if full or not same_build else refused
        kept = [c.formula for c in cases if c.formula in settled]
        pending = [c for c in cases if c.formula not in settled]
        for start in range(0, len(pending), BATCH):
            self.walk(pending[start : start + BATCH])
        found = set(self.refused)
        self.refused = [c.formula for c in cases if c.formula in found or c.formula in kept]
        return self.refused

    def summary(self, total: int) -> str:
        per = self.seconds / self.opens if self.opens else 0.0
        return (
            f"{total} formulas, {self.opens} opens in {self.seconds:.0f} s ({per:.2f} s an open), "
            f"{len(self.refused)} refused"
        )
