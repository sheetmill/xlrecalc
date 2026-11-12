"""Whether setting a formula in an open workbook gives a file's verdict on it (CAPTURING.md §2).

    uv run python -m capture.entry_probe [--limit N]
    uv run python -m capture.entry_probe --batches N [N ...]

Excel refuses a whole file for one formula it does not take (-50), so the probes that name
refused formulas (``empty_probe``, ``database_probe``) open one workbook for each known refusal
after a build changes. A formula Excel refuses, set through AppleScript's ``formula`` of a cell,
leaves the cell unchanged without an error. This sets every formula those probes settled on the
current build, accepted and refused, in its formula-bar form (``display_form``), into one open
workbook (``set_formulas.applescript``), reads each back, and compares: accepted if the cell holds
a formula afterwards. A formula longer than 255 characters cannot be set this way and is counted
apart. Prints the agreements, every disagreement, and how long the sets took.

``--batches`` opens workbooks of the first N formulas ``empty_args`` holds that Excel takes in a
file (none refused), for each N given, and prints whether each opened and how long it took: how
large a batch can confirm a known-accepted set in one open.
"""

from __future__ import annotations

import re
import sys
import tempfile
import time
from pathlib import Path

import openpyxl

from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER
from capture.refusals import known_accepted, known_refused
from xlrecalc._formula._fileform import display_form

HERE = Path(__file__).resolve().parent
US = "\x1f"
SOURCES = (
    ("empty", CAPTURES / "empty_args_refused.tsv", CAPTURES / "empty_args.tsv"),
    ("database", CAPTURES / "database_refused.tsv", CAPTURES / "database_family.tsv"),
)


def _filled(formula: str) -> str:
    """A suite's formula with each input placeholder ``{k}`` as a cell of row 1 (syntax only)."""
    return re.sub(r"\{(\d+)\}", lambda m: f"{chr(65 + int(m[1]))}1", formula)


def forms() -> list[tuple[str, str, bool]]:
    """(source, file form, refused in a file) for every form the probes settled."""
    out = []
    for source, refused_path, accepted_path in SOURCES:
        refused, _ = known_refused(refused_path)
        out += [(source, f, True) for f in sorted(refused)]
        out += [(source, _filled(f), False) for f in sorted(known_accepted(accepted_path))]
    return out


def batches(sizes: list[int]) -> None:
    from capture.pipeline import run_excel
    from capture.suites import _EMPTY_TABLE, empty_args_all

    refused, _ = known_refused(CAPTURES / "empty_args_refused.tsv")
    cases = [c for c in empty_args_all() if c.formula not in refused]
    for n in sizes:
        start = time.monotonic()
        try:
            run_excel(cases[:n], "entry_batch", False, _EMPTY_TABLE)
            verdict = "opened"
        except RuntimeError as exc:
            verdict = f"failed: {str(exc)[-80:]}"
        print(
            f"{min(n, len(cases))} formulas: {verdict} in {time.monotonic() - start:.1f} s",
            flush=True,
        )


def entered(file_forms: list[str]) -> list[bool | None]:
    """For each formula (file form), whether Excel takes it set in an open workbook.

    ``None`` for one longer than AppleScript's 255 characters, which cannot be set this way.
    """
    typed = ["=" + display_form(f) for f in file_forms]
    long = {k for k, t in enumerate(typed) if len(t) > 256}
    path = CONTAINER / "xlrecalc_entry.xlsx"
    if path.exists():
        sys.exit(f"{path.name} exists: remove it or close it in Excel first")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    openpyxl.Workbook().save(path)
    with tempfile.NamedTemporaryFile("w", suffix=".txt", encoding="utf-8", delete=False) as f:
        f.write("\n".join("" if k in long else t for k, t in enumerate(typed)))
        listing = f.name
    try:
        out = osascript_or_exit(
            str(HERE / "set_formulas.applescript"), str(path), listing, seconds=3600
        ).stdout
    finally:
        path.unlink(missing_ok=True)
        Path(listing).unlink(missing_ok=True)
    back = {}
    for line in out.split("\n"):
        if US in line:
            k, text = line.split(US, 1)
            back[int(k) - 1] = text
    return [
        None if k in long else bool(back.get(k)) and not back[k].startswith("error ")
        for k in range(len(typed))
    ]


def still_refused(file_forms: list[str]) -> set[str]:
    """The formulas Excel refuses set in an open workbook (``Probe.entry``)."""
    return {f for f, ok in zip(file_forms, entered(file_forms), strict=True) if ok is False}


def main() -> None:
    if "--batches" in sys.argv:
        batches([int(a) for a in sys.argv[sys.argv.index("--batches") + 1 :]])
        return
    limit = int(sys.argv[sys.argv.index("--limit") + 1]) if "--limit" in sys.argv else None
    cases = forms()[:limit]
    typed = ["=" + display_form(f) for _, f, _ in cases]
    start = time.monotonic()
    verdicts = entered([f for _, f, _ in cases])
    seconds = time.monotonic() - start
    long = {k for k, v in enumerate(verdicts) if v is None}
    agree, differ = 0, []
    for k, (source, form, refused) in enumerate(cases):
        if k in long:
            continue
        if verdicts[k] != (not refused):
            outcome = "refused in a file" if refused else "opened in a file"
            differ.append((source, form, typed[k], outcome, "taken" if verdicts[k] else "refused"))
        else:
            agree += 1
    print(
        f"{len(cases)} forms, {len(long)} longer than 255, {agree} agree, {len(differ)} differ; "
        f"{seconds:.0f} s to set them"
    )
    for row in differ:
        print("\t".join(row))


if __name__ == "__main__":
    main()
