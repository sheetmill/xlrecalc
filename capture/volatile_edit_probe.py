"""Which formulas an edit makes Excel recompute after it kept their values on opening.

    uv run python -m capture.volatile_edit_probe

One workbook a variant (``capture/volatile_edit_cases.py``): openpyxl writes it (``probe!A1`` =
0.1, ``B1`` the formula under test, ``C1`` = 1, ``D1 = C1*10``, ``E1 = A1*3``, and the variant's
defined names and cells) and Excel saves a copy, so the file has Excel's own ``calcId``,
``<fileVersion>`` and cell forms. In that copy B1's ``ca="1"`` mark is removed or written, as the
variant asks, the marks of the cells in ``UNMARKED`` are removed, and the cached values of B1,
D1, E1, F1, G1 and H1 are planted (7 for B1, 999 for the others: only the ``<v>`` text). Excel
opens it with alerts off and shows the cells, then takes the variant's steps (``STEPS``: edits,
``calculate``, ``calculate full``), showing the cells after each (``edit_steps.applescript``);
the workbook is closed without saving. The manual variants come last, and
``calc_automatic.applescript`` sets Excel's mode back after each.

B1 reads C1 only through ``INDIRECT``'s text in ``name_no_ca``, so it changes after the edit
only if Excel computes it as volatile; ``name_dependent_no_ca`` names C1
(``OFFSET(probe!$C$1,0,0)``), so there B1 is a dependent of the edit too. D1 is a plain
dependent of an edit of C1; E1 of an edit of A1. ``PREDICTED`` holds what was expected before
the run, per cell: after opening, after a first full calculation where the steps start with one,
then after each further step.

The workbooks live in Excel's container and are removed after; the capture is written to
``tests/captures/volatile_edit.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys

from capture.edition import edition_line
from capture.excel_watch import preflight
from capture.open_calc_probe import calc_attributes, edit_calc
from capture.pipeline import CAPTURES
from capture.volatile_edit_cases import (
    FULL_ON_LOAD,
    MANUAL,
    NAMES,
    UNMARKED,
    VARIANTS,
    columns,
    steps,
)
from capture.volatile_name_probe import cell_xml
from capture.writeback_probe import CONTAINER, osascript

US = "\x1f"
PLANTED = {"B1": "7", "D1": "999", "E1": "999", "F1": "999", "G1": "999", "H1": "999"}
NEW = "new"  # a value in [2, 3) other than the one shown at the step before

# D1 and E1 after an edit of C1 to 2, of A1 to 2, and of C1 to 2 then A1 to 3.
_C1 = {"D1": (999, 20, 20), "E1": (999, 999, 0.3)}
_A1 = {"D1": (999, 999, 10), "E1": (999, 6, 6)}
_C1_A1 = {"D1": (999, 20, 20, 20), "E1": (999, 999, 9, 9)}

# Expected before the run: (after opening, after the edit, after the full calculation).
PREDICTED = {
    "name_no_ca": {"B1": (7, 7, 2), "D1": (999, 20, 20), "E1": (999, 999, 0.3)},
    "name_ca": {"B1": (1, 2, 2), "D1": (999, 20, 20), "E1": (999, 999, 0.3)},
    "own_no_ca": {"B1": (1, 2, 2), "D1": (10, 20, 20), "E1": (0.3, 0.3, 0.3)},
    "now_no_ca": {
        "B1": (0.1, 0.1, 0.1),
        "D1": (10, 20, 20),
        "E1": (0.3, 0.3, 0.3),
        "F1": (3, 3, 3),
    },
    "name_dependent_no_ca": {"B1": (7, 2, 2), "D1": (999, 20, 20), "E1": (999, 999, 0.3)},
    # plain_name, name_offset_plus, name_static_dependent, name_no_ca_full_first and
    # name_no_ca_full_on_load, expected before they ran (after opening, [after a first full
    # calculation,] after the edit, after the full calculation): a name's references followed as
    # any others but OFFSET's reference argument, and a cell Excel computed tracked as volatile
    # from then on.
    "plain_name": {"B1": (7, 2, 2)},
    "name_offset_plus": {"B1": (7, 2, 2)},
    "name_static_dependent": {"B1": (7, 3, 3), "D1": (999, 999, 10), "E1": (999, 6, 6)},
    "name_no_ca_full_first": {"B1": (7, 1, 2, 2), "D1": (999, 10, 20, 20)},
    "name_no_ca_full_on_load": {"B1": (1, 2, 2), "D1": (10, 20, 20), "E1": (0.3, 0.3, 0.3)},
    # The variants from own_plus_offset on, expected before they ran: a cell volatile only
    # through a defined name, not computed since opening (dormant), has edges from its own text
    # only, whatever the volatile function in the name; an edit of a cell its text names computes
    # it, the name evaluated as usual, and from then on it is volatile; nothing reaches its
    # dependents through it, and a dependent computed for another reason reads its cached value;
    # a plain name beside the volatile one counts as the cell's own text; in manual calculation an
    # edit computes nothing, and calculate computes the edit's dependents but not the dormant cell.
    # dep_marked: Excel writes ca="1" on G1 when it saves (seen in G1's planted XML).
    "own_plus_offset": {"B1": (7, 4, 4), **_C1},
    "own_plus_indirect": {"B1": (7, 2.1, 2.1), **_C1},
    "own_name_ref_edit": {"B1": (7, 7, 3), **_A1},
    "dep_unmarked": {"B1": (7, 7, 2), "G1": (999, 999, 2), "H1": (999, 999, 2), **_C1},
    "dep_unmarked_plus": {"B1": (7, 7, 2), "G1": (999, 9, 4), **_C1},
    "dep_of_own_unmarked": {"B1": (7, 4, 4), "G1": (999, 4, 4), **_C1},
    "dep_marked": {"B1": (7, 7, 2), "G1": (7, 7, 2), **_C1},
    "second_indirect": {"B1": (7, 2.1, 5, 5), **_C1_A1},
    "second_offset": {"B1": (7, 2.1, 5, 5), **_C1_A1},
    "second_dormant": {"B1": (7, 7, 7, 3), **_C1_A1},
    "now_name": {"B1": (7, 7, 2), **_C1},
    "today_name": {"B1": (7, 7, 2), **_C1},
    "rand_name": {"B1": (7, 7, 2), **_C1},
    "randbetween_name": {"B1": (7, 7, 2), **_C1},
    "randarray_name": {"B1": (7, 7, 2), **_C1},
    "cell_name": {"B1": (7, 7, 2), **_C1},
    "awake_unrelated": {
        "B1": (7, NEW, NEW, NEW),
        "D1": (999, 20, 20, 20),
        "E1": (999, 999, 999, 0.3),
    },
    "plain_and_volatile": {"B1": (7, 3, 3), **_A1},
    "nested_name": {"B1": (7, 7, 2), **_C1},
    "manual_name": {"B1": (7, 7, 7, 2), "D1": (999, 999, 20, 20), "E1": (999, 999, 999, 0.3)},
}


def build(path: object, variant: str) -> None:
    import openpyxl  # imported here: the tests read VARIANTS without it
    from openpyxl.workbook.defined_name import DefinedName

    formula, name, _, extra = VARIANTS[variant]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    ws["A1"] = 0.1
    ws["B1"] = formula
    ws["C1"] = 1
    ws["D1"] = "=C1*10"
    ws["E1"] = "=A1*3"
    for ref, text in extra.items():
        ws[ref] = text
    names = {} if name is None else {"vn": name}
    for key, text in {**names, **NAMES.get(variant, {})}.items():
        wb.defined_names[key] = DefinedName(key, attr_text=text)
    wb.save(path)


def plant(xml: str, marked: bool, unmarked: tuple[str, ...] = ()) -> str:
    """B1 marked or not, the ``unmarked`` cells' marks removed; the planted cells' ``<v>`` text
    replaced, the rest as Excel wrote it. F1 keeps the mark Excel wrote."""
    for ref in ("B1", *unmarked):
        xml = re.sub(rf'(<c r="{ref}"[^>]*><f)\s+ca="1"', r"\1", xml, count=1)
    if marked:
        xml, n = re.subn(r'(<c r="B1"[^>]*><f)>', r'\1 ca="1">', xml, count=1)
        if n != 1:
            raise RuntimeError("B1's formula has attributes: cannot mark it")
    for ref, value in PLANTED.items():
        cell = re.compile(rf'(<c r="{ref}"[^>]*>.*?<v>)(.*?)(</v>)', re.S)
        xml = cell.sub(lambda m, v=value: m.group(1) + v + m.group(3), xml, count=1)
    return xml


def _same(text: str) -> str:
    return text


def _flag(text: str) -> str:
    """``fullCalcOnLoad="1"`` in workbook.xml's ``<calcPr>``."""
    return re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', text, count=1)


def _manual(text: str) -> str:
    """``calcMode="manual"`` in workbook.xml's ``<calcPr>``."""
    return re.sub(r"<calcPr\b", '<calcPr calcMode="manual"', text, count=1)


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    rows: list[str] = []
    calc = "-"
    for variant, (_, _, marked, extra) in VARIANTS.items():
        src = CONTAINER / f"xlrecalc_ve_{variant}_src.xlsx"
        base = CONTAINER / f"xlrecalc_ve_{variant}_base.xlsx"
        path = CONTAINER / f"xlrecalc_ve_{variant}.xlsx"
        cells = ["A1", "B1", "C1", "D1", "E1", *extra]
        after = steps(variant)
        try:
            build(src, variant)
            osascript("save_copy.applescript", str(src), str(base))
            workbook = _flag if variant in FULL_ON_LOAD else _manual if variant in MANUAL else _same
            unmarked = UNMARKED.get(variant, ())
            edit_calc(base, path, lambda x, m=marked, u=unmarked: plant(x, m, u), workbook)
            if workbook is _same:
                calc = calc_attributes(path)
            planted = {ref: cell_xml(path, ref) for ref in cells}
            refs = [f"probe!{ref}" for ref in cells]
            fields = osascript("edit_steps.applescript", str(path), ";".join(after), *refs)
        finally:
            if variant in MANUAL and base.exists():  # the application's mode back to automatic
                print("  restore: " + osascript("calc_automatic.applescript", str(base)))
            for p in (src, base, path):
                p.unlink(missing_ok=True)
        fields_list = fields.split(US)
        if fields_list[0] != "open":
            sys.exit(f"{variant} did not open: {fields_list}")
        k = len(cells)
        shown = [fields_list[1 : 1 + k]]
        for at in range(1 + k, len(fields_list), 1 + k):  # each step: its label, then the cells
            shown.append(fields_list[at + 1 : at + 1 + k])
        if after[0] != "full":
            shown.insert(1, ["-"] * k)
        if len(shown) != columns(variant):
            sys.exit(f"{variant}: {len(shown)} readings for {columns(variant)} columns")
        for i, ref in enumerate(cells):
            row = [variant, f"probe!{ref}", planted[ref], *(step[i] for step in shown)]
            rows.append("\t".join(row))
            print("  ".join(row), flush=True)
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
        capture_output=True,
        text=True,
    ).stdout.strip()
    header = [
        "# suite: volatile_edit",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/volatile_edit_probe.py",
        f"# workbook: {calc}",
        "# method: Excel's saved copy, B1 marked or not, the marks of listed dependents",
        "#   removed, B1 D1 E1 F1 G1 H1 planted, fullCalcOnLoad or calcMode=manual written where",
        "#   the variant asks; shown after opening, then after each of the variant's steps",
        "#   (STEPS in capture/volatile_edit_cases.py: an edit, calculate, calculate full); the",
        "#   workbook line is that of the files without the flag or the mode",
        "# columns: variant cell planted_xml shown_after_opening shown_after_full_first ('-'",
        "#   where the steps do not start with calculate full) then one column per further step;",
        "#   the last is always after calculate full",
    ]
    out = CAPTURES / "volatile_edit.tsv"
    out.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")
    print(f"-> {out.name}, {len(rows)} rows")


if __name__ == "__main__":
    main()
