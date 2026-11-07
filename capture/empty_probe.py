"""Which formulas with an empty argument Excel refuses in a file (``CAPTURING.md`` §1).

    uv run python -m capture.empty_probe

Excel refuses a whole file for one such formula (-50),
so the suite ``empty_args`` is opened in small
batches, each halved until the refused formulas are named. They go to
``tests/captures/empty_args_refused.tsv``; the suite leaves them out, so it can be captured at once.
"""

from __future__ import annotations

import sys
from pathlib import Path

from capture.pipeline import CAPTURES, Case, run_excel
from capture.suites import _EMPTY_TABLE, empty_args_all

BATCH = 60
MAX_OPENS = 2500


def main() -> None:
    opens = 0
    refused: list[str] = []

    def walk(cases: list[Case]) -> None:
        nonlocal opens
        opens += 1
        if opens > MAX_OPENS:
            sys.exit(f"more than {MAX_OPENS} opens; so far {len(refused)} refused")
        try:
            run_excel(cases, "empty_probe", False, _EMPTY_TABLE)
        except RuntimeError:
            if len(cases) == 1:
                refused.append(cases[0].formula)
                return
            middle = len(cases) // 2
            walk(cases[:middle])
            walk(cases[middle:])

    cases = empty_args_all()
    for start in range(0, len(cases), BATCH):
        walk(cases[start : start + BATCH])
    path: Path = CAPTURES / "empty_args_refused.tsv"
    lines = ["# formulas Excel refuses in a file (capture/empty_probe.py)", *refused]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{len(cases)} formulas, {opens} opens, {len(refused)} refused -> {path.name}")


if __name__ == "__main__":
    main()
