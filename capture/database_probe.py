"""Which error and constant arguments of the database functions Excel refuses in a file.

    uv run python -m capture.database_probe [--full]

Excel refuses a whole file for one such formula (-50, ``CAPTURING.md`` §2), so the forms are
opened in batches, each refused batch halved until every refusal is opened alone
(``capture/refusals.py``): a known refusal of the same build is kept without an open, and the
rest confirmed in full batches, unless ``--full``. The refused formulas go to
``tests/captures/database_refused.tsv``, which the suite ``database_family`` leaves out, so it can
be captured at once.
"""

from __future__ import annotations

import sys

from capture.edition import edition_line
from capture.pipeline import CAPTURES, Case, excel_version, run_excel
from capture.refusals import Probe, known_accepted, known_refused
from capture.suites import _DB_EXTRA, MISC_TABLE, database_errors_all

FIXED = {**MISC_TABLE, **_DB_EXTRA}
REFUSED = CAPTURES / "database_refused.tsv"


def _opens(cases: list[Case]) -> bool:
    try:
        run_excel(cases, "database_probe", False, FIXED)
    except RuntimeError as exc:
        if "-50" not in str(exc):  # not a refusal: Excel failed otherwise
            sys.exit(f"Excel failed otherwise: {exc}")
        return False
    return True


def main() -> None:
    cases = database_errors_all()
    build = excel_version()
    refused, refused_build = known_refused(REFUSED)
    accepted = known_accepted(CAPTURES / "database_family.tsv")
    new = sum(1 for c in cases if c.formula not in refused and c.formula not in accepted)
    probe = Probe(_opens)
    probe.run(cases, refused, refused_build == build, "--full" in sys.argv)
    lines = [
        "# formulas Excel refuses in a file (capture/database_probe.py)",
        f"# excel: {build} for Mac, locale en_US",
        edition_line(),
        *probe.refused,
    ]
    REFUSED.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{probe.summary(len(cases))} ({new} not seen before) -> {REFUSED.name}")


if __name__ == "__main__":
    main()
