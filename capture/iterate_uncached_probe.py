"""How Excel iterates a workbook whose formulas have no cached values (ARCHITECTURE.md §7.2).

    uv run python -m capture.iterate_uncached_probe

Non-interactive. Each variant is one small workbook with iterative calculation on and no cached
value in any formula, as openpyxl writes it, with ``fullCalcOnLoad`` set (as openpyxl sets it) and,
as a second file, without it; a third file gives every formula a cached 0. Each is opened alone:
what Excel shows after opening and after one ``calculate`` (``iterate_read.applescript``), then,
opened again, a saved copy's exact values (``save_copy.applescript``). The counts are small, so
each candidate rule (an initial pass outside the count, how a cell not yet calculated reads to
itself and to others, which changes the stopping test sees) predicts a different number. Writes
``tests/captures/iterate_uncached.tsv``.

Needs Excel with no workbook open, and refuses to run otherwise.
"""

from __future__ import annotations

import datetime as dt
import sys

from capture.iterate_probe import HERE, US, Sheets, build, osascript, saved_values
from capture.pipeline import CAPTURES, CONTAINER, excel_version


def one(*cells: tuple[str, str]) -> Sheets:
    return [("s", [(ref, formula, None) for ref, formula in cells])]


# (name, calcPr attributes after iterate="1", cells)
VARIANTS: list[tuple[str, str, Sheets]] = [
    ("self_plain", 'iterateCount="5"', one(("A1", "A1+1"))),
    ("self_iferror", 'iterateCount="5"', one(("A1", "IFERROR(A1,0)+1"))),
    ("self_isnumber", 'iterateCount="5"', one(("A1", "IF(ISNUMBER(A1),A1+1,0)"))),
    ("self_iferror_d10", 'iterateCount="5" iterateDelta="10"', one(("A1", "IFERROR(A1,0)+3"))),
    ("self_iferror_d2", 'iterateCount="5" iterateDelta="2"', one(("A1", "IFERROR(A1,0)+3"))),
    ("pair_plain", 'iterateCount="5"', one(("A1", "A2+1"), ("A2", "A1+1"))),
    ("pair_first", 'iterateCount="5"', one(("A1", "IFERROR(A2,0)+1"), ("A2", "A1+1"))),
    ("pair_second", 'iterateCount="5"', one(("A1", "A2+1"), ("A2", "IFERROR(A1,0)+1"))),
    ("pair_both", 'iterateCount="5"', one(("A1", "IFERROR(A2,0)+1"), ("A2", "IFERROR(A1,0)+1"))),
    (
        "pair_first_d10",
        'iterateCount="5" iterateDelta="10"',
        one(("A1", "IFERROR(A2,0)+1"), ("A2", "A1+1")),
    ),
    ("pair_first_count1", 'iterateCount="1"', one(("A1", "IFERROR(A2,0)+1"), ("A2", "A1+1"))),
    (
        "dep_self",
        'iterateCount="5" iterateDelta="10"',
        one(("A1", "IFERROR(A1,0)+3"), ("B1", "A1*2")),
    ),
    (
        "three_below",
        'iterateCount="5"',
        one(("A1", "A3*2"), ("A2", "IFERROR(A3,0)+1"), ("A3", "A2+1")),
    ),
    (
        "isblank_self",
        'iterateCount="5"',
        one(("A1", 'IF(ISBLANK(A1),"blank",IF(ISERROR(A1),"error",A1+1))')),
    ),
    ("type_self", 'iterateCount="3"', one(("A1", "TYPE(A1)*10+IFERROR(A1,0)"))),
    ("type_other", 'iterateCount="3"', one(("A1", "TYPE(A2)"), ("A2", "A1+0"))),
    ("half_iferror", 'iterateCount="100"', one(("A1", "IFERROR(A2,0)/2+1"), ("A2", "A1"))),
]
CALC = {
    "flag": '<calcPr calcId="124519" iterate="1" {} fullCalcOnLoad="1"/>',
    "no_flag": '<calcPr calcId="191029" iterate="1" {}/>',
    "cached_0": '<calcPr calcId="191029" iterate="1" {}/>',
}


def main() -> None:
    count = osascript("-e", 'tell application "Microsoft Excel" to count of workbooks')
    if count != "0":
        sys.exit(f"Excel has {count} workbooks open: iteration is taken from the first one opened")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    body: list[str] = []
    made = []
    try:
        for name, attrs, sheets in VARIANTS:
            for form, calc in CALC.items():
                cells = sheets
                if form == "cached_0":
                    cells = [(s, [(r, f, "0") for r, f, _ in cs]) for s, cs in sheets]
                path = CONTAINER / f"xlrecalc_iter_unc_{name}_{form}.xlsx"
                copy = path.with_name(path.stem + "_saved.xlsx")
                made += [path, copy]
                copy.unlink(missing_ok=True)
                build(path, calc.format(attrs), cells)
                refs = [f"{s}!{r}" for s, cs in cells for r, _, _ in cs]
                fields = osascript(str(HERE / "iterate_read.applescript"), str(path), *refs)
                fields_list = fields.split(US)
                if fields_list[0] != "open":
                    sys.exit(f"{path.name} did not open: {fields_list}")
                k = len(refs)
                first, second = fields_list[1 : 1 + 2 * k : 2], fields_list[2 + 2 * k :: 2]
                osascript(str(HERE / "save_copy.applescript"), str(path), str(copy))
                saved = saved_values(copy)
                for i, (ref, (_, formula, _)) in enumerate(
                    zip(refs, [c for _, cs in cells for c in cs], strict=True)
                ):
                    kind, value = saved.get(ref, ("none", ""))
                    row = [name, form, attrs, ref, formula, first[i], second[i], kind, value]
                    body.append("\t".join(row))
                    print("  ".join(row), flush=True)
                path.unlink(missing_ok=True)
                copy.unlink(missing_ok=True)
    finally:
        for path in made:
            path.unlink(missing_ok=True)
    header = [
        "# suite: iterate_uncached",
        f"# excel: {excel_version()} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/iterate_uncached_probe.py",
        "# method: each variant as three files: no cached values with fullCalcOnLoad and",
        "#   openpyxl's calcId (flag), no cached values without it (no_flag), every formula",
        "#   cached 0 (cached_0); shown after opening, calculated after one calculate, saved",
        "#   the value in a copy saved after opening again",
        "# columns: variant form calcPr cell formula shown calculated saved_type saved",
    ]
    out = CAPTURES / "iterate_uncached.tsv"
    out.write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    print(f"-> {out.name}, {len(body)} rows")


if __name__ == "__main__":
    main()
