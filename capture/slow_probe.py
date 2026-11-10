"""How long Excel takes over one formula alone in a workbook, or that it passes a time limit.

    uv run python -m capture.slow_probe [--limit SECONDS] {shapes,third}... | all

Each case is built into a workbook of its own (inputs in cells, as the pipeline writes them),
opened, calculated and saved by ``save_copy.applescript`` under a time limit. A case Excel
finishes gives its value and time; one past the limit is recorded as such, and Excel is ended when
``capture.excel_watch.may_end`` allows it (the run began with no workbook open, and only the
probe's windows are on screen); otherwise the run stops, as any request past its limit does.
Results go to ``tests/captures/slow_cases.tsv``, replacing the rows of each set run and keeping
the others. ``shapes`` times functions of the beta and gamma families as their shape grows;
``third`` is the third set of ``continuous_family``, which Excel must finish before the pipeline
captures it.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

from capture.edition import edition_line
from capture.excel_watch import ExcelTimeoutError, osascript
from capture.pipeline import CONTAINER, HERE, build, excel_version, read_results, result_repr
from capture.suites import Case, continuous_third_set

OUT = Path(__file__).resolve().parents[1] / "tests" / "captures" / "slow_cases.tsv"


def shapes() -> list[Case]:
    """A beta, a gamma and an inverse at shapes from 1E5 to 1E100."""
    steps = (1e5, 1e10, 1e20, 1e50, 1e100)
    return [
        *(Case("beta.dist", "_xlfn.BETA.DIST(0.5,{1},0.5,TRUE)", (s,)) for s in steps),
        *(Case("gamma.dist", "_xlfn.GAMMA.DIST({1},{1},1,TRUE)", (s,)) for s in steps),
        *(Case("beta.inv", "_xlfn.BETA.INV(0.5,{1},0.5)", (s,)) for s in steps),
    ]


def third() -> list[Case]:
    """``continuous_family``'s third set with the seven formulas the suite leaves out."""
    return continuous_third_set(slow=True)


def run(cases: list[Case], limit: float) -> list[tuple[Case, str, str, float | None]]:
    """Each case alone: its kind, value and seconds, or ``past`` and None."""
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_slow_probe.xlsx"
    copy = CONTAINER / "xlrecalc_slow_probe_saved.xlsx"
    out: list[tuple[Case, str, str, float | None]] = []
    for case in cases:
        for p in (src, copy):
            p.unlink(missing_ok=True)
        build(src, [case])
        print(f"{case.formula} {case.inputs} ...", end=" ", flush=True)
        start = time.monotonic()
        try:
            result = osascript(
                str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=limit,
                end_allowed=True,
            )  # fmt: skip
        except ExcelTimeoutError as timeout:
            if not timeout.ended:
                sys.exit(str(timeout))
            print(f"past {limit:.0f} s, Excel ended", flush=True)
            out.append((case, "past", f"{limit:.0f} s", None))
            continue
        seconds = time.monotonic() - start
        if result.returncode:
            sys.exit(f"{case.formula}: {result.stderr.strip()}")
        kind, value = result_repr(read_results(copy, 1)[0])
        print(f"{value} in {seconds:.1f} s", flush=True)
        out.append((case, kind, value, seconds))
    for p in (src, copy):
        p.unlink(missing_ok=True)
    return out


SETS = {"shapes": shapes, "third": third}


def main(argv: list[str]) -> None:
    limit = 90.0
    if "--limit" in argv:
        at = argv.index("--limit")
        limit = float(argv[at + 1])
        del argv[at : at + 2]
    chosen = list(SETS) if argv == ["all"] else argv
    if not chosen or any(which not in SETS for which in chosen):
        sys.exit(
            f"usage: python -m capture.slow_probe [--limit SECONDS] {{{','.join(SETS)}}}... | all"
        )
    kept: dict[str, list[str]] = {which: [] for which in SETS}
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line and not line.startswith("#"):
                kept[line.split("\t", 1)[0]].append(line)
    for which in chosen:
        kept[which] = []
        for case, kind, value, seconds in run(SETS[which](), limit):
            shown = "" if seconds is None else f"{seconds:.1f}"
            kept[which].append(
                f"{which}\t{case.tag}\t{case.formula}\t{case.inputs}\t{kind}\t{value}\t{shown}"
            )
    lines = [
        "# suite: slow_cases (how long Excel takes over one formula alone in a workbook)",
        f"# excel: {excel_version()} for Mac, locale en_US",
        f"# captured: {time.strftime('%Y-%m-%d')}",
        edition_line(),
        f"# limit: {limit:.0f} s a case; past it Excel was ended (capture.excel_watch.may_end)",
        "# columns: set, tag, formula, inputs, kind, value, seconds",
    ]
    lines += [line for which in SETS for line in kept[which]]
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1:])
