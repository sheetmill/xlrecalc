"""Which error and constant arguments of the database functions Excel refuses in a file.

    uv run python -m capture.database_probe

Excel refuses a whole file for one such formula (-50, ``CAPTURING.md`` §2). Each form is opened
alone for ``DSUM``; for every other database function, the forms ``DSUM`` accepted are opened
together, and each form ``DSUM`` refused is opened alone, so every refusal listed was seen by
itself. The refused formulas go to ``tests/captures/database_refused.tsv``, which the suite
``database_family`` leaves out, so it can be captured at once.
"""

from __future__ import annotations

from pathlib import Path

from capture.edition import edition_line
from capture.pipeline import CAPTURES, Case, excel_version, run_excel
from capture.suites import _DB_EXTRA, MISC_TABLE, database_errors_all

FIXED = {**MISC_TABLE, **_DB_EXTRA}


def opens(cases: list[Case]) -> bool:
    try:
        run_excel(cases, "database_probe", False, FIXED)
    except RuntimeError:
        return False
    return True


def main() -> None:
    cases = database_errors_all()
    by_fn: dict[str, list[Case]] = {}
    for case in cases:
        by_fn.setdefault(case.formula.split("(", 1)[0], []).append(case)
    count = 0
    refused_forms: set[str] = set()
    refused: list[str] = []
    for case in by_fn.pop("DSUM"):
        count += 1
        if not opens([case]):
            refused_forms.add(case.formula.removeprefix("DSUM"))
            refused.append(case.formula)
    for fn, group in by_fn.items():
        accepted = [c for c in group if c.formula.removeprefix(fn) not in refused_forms]
        count += 1
        if not opens(accepted):
            raise SystemExit(f"{fn}: a form DSUM accepts was refused; bisect before going on")
        for case in group:
            if case.formula.removeprefix(fn) in refused_forms:
                count += 1
                if opens([case]):
                    raise SystemExit(f"{case.formula}: accepted where DSUM's form was refused")
                refused.append(case.formula)
    path: Path = CAPTURES / "database_refused.tsv"
    lines = [
        "# formulas Excel refuses in a file (capture/database_probe.py)",
        f"# excel: {excel_version()} for Mac, locale en_US",
        edition_line(),
        *refused,
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{len(cases)} formulas, {count} opens, {len(refused)} refused -> {path.name}")


if __name__ == "__main__":
    main()
