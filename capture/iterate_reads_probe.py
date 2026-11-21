"""What Excel's opening run gives loops apart or reading each other (ARCHITECTURE.md §7.2).

    uv run python -m capture.iterate_reads_probe

The machinery of ``capture/iterate_uncached_probe.py``, one file per variant, every formula
cached 0, ``fullCalcOnLoad`` set and 100 iterations: what Excel shows after opening, after one
``calculate``, and in a copy saved after opening again. The variants test one reading of capture
``iterate_calc``: that opening makes one run over every loop together, in Excel's order, with one
stopping test for all of them. Two counters that never settle (each run adds the count); a loop of
one cell that settles in about 11 iterations beside a loop of three that needs about 66 (with one
stopping test the first runs as long as the second); a loop read by another. Writes
``tests/captures/iterate_reads.tsv``.

Needs Excel with no workbook open, and refuses to run otherwise.
"""

from __future__ import annotations

from capture.iterate_uncached_probe import main
from capture.iterate_uncached_probe import one as formulas

COUNT = 'iterateCount="100"'
VARIANTS = [
    ("counters", COUNT, formulas(("B1", "B1+1"), ("C1", "C1+1"), ("D1", "B1+C1"))),
    (
        "lengths",
        COUNT,
        formulas(("B1", "B1/2+1"), ("C1", "E1*0.9+1"), ("D1", "C1"), ("E1", "D1")),
    ),
    ("feeds", COUNT, formulas(("B1", "B1/2+1"), ("C1", "C1/2+B1"))),
]
# every formula cached 0 (main's "cached_0"), with fullCalcOnLoad: Excel's opening run
FORMS = {"cached_0": '<calcPr calcId="124519" iterate="1" {} fullCalcOnLoad="1"/>'}

METHOD = [
    "# method: each variant as one file, every formula cached 0, with fullCalcOnLoad",
    "#   (cached_0); shown after opening, calculated after one calculate, saved the value",
    "#   in a copy saved after opening again",
]

if __name__ == "__main__":
    main(VARIANTS, FORMS, "iterate_reads", "capture/iterate_reads_probe.py", METHOD)
