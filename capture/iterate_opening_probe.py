"""What Excel's opening does to a loop with no cached values (ARCHITECTURE.md §7.2).

    uv run python -m capture.iterate_opening_probe

The machinery of ``capture/iterate_uncached_probe.py``, one file per variant as openpyxl writes
it (no cached values, ``fullCalcOnLoad``), read after opening, after one ``calculate`` and saved.
The variants test one reading of capture ``iterate_uncached``: that opening leaves a cell
uncalculated when its formula first reads a cell not yet calculated. The same loops in other row
and column orders, entered from a constant outside or not, with a formula outside it, with
IFERROR in each cell in turn, and with counts 1 to 3 to count opening's iterations. Writes
``tests/captures/iterate_opening.tsv``.
"""

from __future__ import annotations

from capture.iterate_probe import Sheets
from capture.iterate_uncached_probe import main
from capture.iterate_uncached_probe import one as formulas


def one(*cells: tuple[str, str]) -> Sheets:
    """Formulas without cached values; ``("C1", "")`` is the constant 5."""
    sheets = formulas(*cells)
    return [(s, [(r, f, "5" if not f else None) for r, f, _ in cs]) for s, cs in sheets]


F = 'iterateCount="5"'
VARIANTS = [
    # Two cells, IFERROR in the first or the second, in rows, far rows, columns, reversed.
    ("pf_rows", F, one(("A1", "IFERROR(A2,0)+1"), ("A2", "A1+1"))),
    ("ps_rows", F, one(("A1", "A2+1"), ("A2", "IFERROR(A1,0)+1"))),
    ("pf_far", F, one(("A1", "IFERROR(A9,0)+1"), ("A9", "A1+1"))),
    ("ps_far", F, one(("A1", "A9+1"), ("A9", "IFERROR(A1,0)+1"))),
    ("pf_cols", F, one(("A1", "IFERROR(B1,0)+1"), ("B1", "A1+1"))),
    ("ps_cols", F, one(("A1", "B1+1"), ("B1", "IFERROR(A1,0)+1"))),
    ("pf_written_reversed", F, one(("A2", "A1+1"), ("A1", "IFERROR(A2,0)+1"))),
    # Entered from outside: a constant, and a formula, feeding the loop.
    ("pf_constant_in", F, one(("C1", ""), ("A1", "IFERROR(A2,0)+C1"), ("A2", "A1+1"))),
    ("ps_constant_in", F, one(("C1", ""), ("A1", "A2+C1"), ("A2", "IFERROR(A1,0)+1"))),
    ("pf_formula_in", F, one(("C1", "2+3"), ("A1", "IFERROR(A2,0)+C1"), ("A2", "A1+1"))),
    ("ps_formula_in", F, one(("C1", "2+3"), ("A1", "A2+C1"), ("A2", "IFERROR(A1,0)+1"))),
    ("pf_formula_below", F, one(("A1", "IFERROR(A2,0)+A9"), ("A2", "A1+1"), ("A9", "2+3"))),
    # Three cells, IFERROR in each in turn.
    ("t1", F, one(("A1", "IFERROR(A3,0)+1"), ("A2", "A1+1"), ("A3", "A2+1"))),
    ("t2", F, one(("A1", "A3+1"), ("A2", "IFERROR(A1,0)+1"), ("A3", "A2+1"))),
    ("t3", F, one(("A1", "A3+1"), ("A2", "A1+1"), ("A3", "IFERROR(A2,0)+1"))),
    # A self-reference read by other tests of error, and by none.
    ("self_iserror", F, one(("A1", "IF(ISERROR(A1),0,A1)+1"))),
    ("self_isnumber_late", F, one(("A1", "IF(ISNUMBER(A1+0),A1+1,0)"))),
    ("self_if_first", F, one(("A1", "IF(TRUE,IFERROR(A1,0)+1)"))),
    # Opening's own count.
    *[
        (f"ps_count_{n}", f'iterateCount="{n}"', one(("A1", "A2+1"), ("A2", "IFERROR(A1,0)+1")))
        for n in (1, 2, 3)
    ],
    *[
        (f"isnum_count_{n}", f'iterateCount="{n}"', one(("A1", "IF(ISNUMBER(A1),A1+1,0)")))
        for n in (1, 2, 3)
    ],
]
# The tech lead's eight, from the public path's first comparison.
B = 'iterateCount="100"'
VARIANTS += [
    ("tl_kind2", B, one(("A1", "IFERROR(A2,0)+1"), ("A2", "A1+1"))),
    ("tl_conv", B, one(("B1", ""), ("A1", "0.5*IFERROR(A2,0)+10"), ("A2", "A1/2+B1"))),
    (
        "tl_interest",
        B,
        one(("B1", ""), ("B2", "IFERROR(B4,0)*0.05"), ("B3", "B1+B2"), ("B4", "B3*0.5")),
    ),
    ("tl_count1", 'iterateCount="1"', one(("A1", "IFERROR(A2,0)+1"), ("A2", "A1*2"))),
    (
        "tl_sumself",
        'iterateCount="50"',
        one(("A1", ""), ("A2", ""), ("A3", ""), ("A4", ""), ("A5", "SUM(A1:A4)+IFERROR(A5,0)*0.1")),
    ),
    ("tl_selfinc", B, one(("A1", "IF(ISNUMBER(A1),A1+1,0)"))),
    (
        "tl_bigdelta",
        'iterateCount="100" iterateDelta="10"',
        one(("A1", "IFERROR(A1,0)+3"), ("B1", "A1*2")),
    ),
    ("tl_below", B, one(("A1", "A3*2"), ("A2", "IFERROR(A3,0)+1"), ("A3", "A2+1"))),
]
CONSTANTS = {
    "tl_conv": {"B1": "3"},
    "tl_interest": {"B1": "1000"},
    "tl_sumself": {"A1": "1", "A2": "2", "A3": "3", "A4": "4"},
}
for i, (name, attrs, sheets) in enumerate(VARIANTS):
    values = CONSTANTS.get(name)
    if values:
        cells = [(r, f, values.get(r, c)) for r, f, c in sheets[0][1]]
        VARIANTS[i] = (name, attrs, [(sheets[0][0], cells)])
FORMS = {"flag": '<calcPr calcId="124519" iterate="1" {} fullCalcOnLoad="1"/>'}

if __name__ == "__main__":
    main(VARIANTS, FORMS, "iterate_opening", "capture/iterate_opening_probe.py")
