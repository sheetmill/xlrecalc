"""Which formulas with an empty argument Excel refuses in a file (``CAPTURING.md`` §1).

    uv run python -m capture.empty_probe [--full]
    uv run python -m capture.empty_probe --check FORMULA [FORMULA ...]

Excel refuses a whole file for one such formula (-50), so the suite ``empty_args`` is opened in
batches, each refused batch halved until the refused formulas are named (``capture/refusals.py``):
a known refusal of the same build is kept without an open, and the rest confirmed in full
batches, unless ``--full``. They go to ``tests/captures/empty_args_refused.tsv``; the suite leaves
them out, so it can be captured at once.
"""

from __future__ import annotations

import sys

from capture.edition import edition_line
from capture.pipeline import CAPTURES, Case, excel_version, run_excel
from capture.refusals import Probe, known_accepted, known_refused
from capture.suites import _EMPTY_TABLE, empty_args_all

REFUSED = CAPTURES / "empty_args_refused.tsv"


def _opens(cases: list[Case]) -> bool:
    try:
        run_excel(cases, "empty_probe", False, _EMPTY_TABLE)
    except RuntimeError as exc:
        if "-50" not in str(exc):  # not a refusal: Excel failed otherwise
            sys.exit(f"Excel failed otherwise: {exc}")
        return False
    return True


def main() -> None:
    if sys.argv[1:2] == ["--check"]:  # given formulas, each alone, with Excel's message
        for formula in sys.argv[2:]:
            try:
                run_excel([Case("check", formula)], "empty_probe", False, _EMPTY_TABLE)
                print(f"opens\t{formula}", flush=True)
            except RuntimeError as exc:
                print(f"fails\t{formula}\t{exc}", flush=True)
        return
    cases = empty_args_all()
    build = excel_version()
    refused, refused_build = known_refused(REFUSED)
    accepted = known_accepted(CAPTURES / "empty_args.tsv")
    new = sum(1 for c in cases if c.formula not in refused and c.formula not in accepted)
    probe = Probe(_opens)
    probe.run(cases, refused, refused_build == build, "--full" in sys.argv)
    lines = [
        "# formulas Excel refuses in a file (capture/empty_probe.py)",
        f"# excel: {build} for Mac, locale en_US",
        edition_line(),
        *probe.refused,
    ]
    REFUSED.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{probe.summary(len(cases))} ({new} not seen before) -> {REFUSED.name}")


if __name__ == "__main__":
    main()
