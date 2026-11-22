"""Capture ``volatile_edit``'s variants, read by its probe and by ``tests/test_volatile.py``.

Kept apart from the probe so that reading them imports neither openpyxl nor Excel's tooling.
"""

from __future__ import annotations

READ = 'INDIRECT("probe!C1")'
OFFSET_C1 = "OFFSET(probe!$C$1,0,0)"
OFFSET_A1 = "OFFSET(probe!$A$1,0,0)"
READ_A1 = 'INDIRECT("probe!A1")'

# name: (B1's formula, defined name vn's text or None, B1 marked ca="1", extra cells {ref: formula})
VARIANTS: dict[str, tuple[str, str | None, bool, dict[str, str]]] = {
    "name_no_ca": ("=vn", READ, False, {}),
    "name_ca": ("=vn", READ, True, {}),
    "own_no_ca": ("=" + READ, None, False, {}),
    "now_no_ca": ("=NOW()*0+A1", None, False, {"F1": "=NOW()*0+3"}),
    "name_dependent_no_ca": ("=vn", OFFSET_C1, False, {}),
    "plain_name": ("=vn+0", "probe!$C$1", False, {}),
    "name_offset_plus": ("=vn", OFFSET_C1 + "+probe!$C$1*0", False, {}),
    "name_static_dependent": ("=vn", READ + "+probe!$A$1", False, {}),
    "name_no_ca_full_first": ("=vn", READ, False, {}),
    "name_no_ca_full_on_load": ("=vn", READ, False, {}),
    # Edges from the cell's own text only: an edit of a cell it names computes it, the name
    # evaluated as usual; the name's reference alone does not.
    "own_plus_offset": ("=vn+C1", OFFSET_C1, False, {}),
    "own_plus_indirect": ("=vn+C1", READ_A1, False, {}),
    "own_name_ref_edit": ("=vn+C1", OFFSET_A1, False, {}),
    # Dependents of the cell: unmarked (the probe removes Excel's mark), or as Excel marks them.
    "dep_unmarked": ("=vn", OFFSET_C1, False, {"G1": "=B1*1", "H1": "=G1*1"}),
    "dep_unmarked_plus": ("=vn", OFFSET_C1, False, {"G1": "=B1+C1"}),
    "dep_of_own_unmarked": ("=vn+C1", OFFSET_C1, False, {"G1": "=B1*1"}),
    "dep_marked": ("=vn", OFFSET_C1, False, {"G1": "=B1*1"}),
    # Two edits: the first computes the cell through its own text (or does not), the second
    # edits only the cell the name reads.
    "second_indirect": ("=vn+C1", READ_A1, False, {}),
    "second_offset": ("=vn+C1", OFFSET_A1, False, {}),
    "second_dormant": ("=vn", READ_A1, False, {}),
    # Each volatile function in a name, beside a plain reference to the edited cell.
    "now_name": ("=vn", "NOW()*0+probe!$C$1", False, {}),
    "today_name": ("=vn", "TODAY()*0+probe!$C$1", False, {}),
    "rand_name": ("=vn", "RAND()*0+probe!$C$1", False, {}),
    "randbetween_name": ("=vn", "RANDBETWEEN(1,9)*0+probe!$C$1", False, {}),
    "randarray_name": ("=vn", "SUM(_xlfn.RANDARRAY(1))*0+probe!$C$1", False, {}),
    "cell_name": ("=vn", 'CELL("row")*0+probe!$C$1', False, {}),
    # Computed through its own text, then an edit of a cell nothing reads (J1): recomputed
    # only if the cell is now volatile, which a random name shows.
    "awake_unrelated": ("=vn+C1", "RAND()", False, {}),
    # A plain name beside the volatile one, its cell edited.
    "plain_and_volatile": ("=vn+pn", READ, False, {}),
    # A name volatile through another name.
    "nested_name": ("=vn2", OFFSET_C1, False, {}),
    # Manual calculation (calcMode="manual"): an edit, then calculate, then calculate full.
    "manual_name": ("=vn", READ, False, {}),
}
# More defined names, beside vn.
NAMES: dict[str, dict[str, str]] = {
    "plain_and_volatile": {"pn": "probe!$A$1"},
    "nested_name": {"vn2": "vn*1"},
}
# Formula cells whose ca="1" mark, written by Excel on saving, the probe removes.
UNMARKED: dict[str, tuple[str, ...]] = {
    "dep_unmarked": ("G1", "H1"),
    "dep_unmarked_plus": ("G1",),
    "dep_of_own_unmarked": ("G1",),
}
FULL_ON_LOAD = {"name_no_ca_full_on_load"}  # <calcPr fullCalcOnLoad="1"> written in the file
MANUAL = {"manual_name"}  # <calcPr calcMode="manual"> written in the file; the probe's last
RANDOM = {("awake_unrelated", "probe!B1")}  # in [2, 3) after an edit; compared as changed or not

# The steps after opening, each followed by a reading of the cells: "full" (calculate full),
# "calc" (calculate) or "edit:<cell>=<number>". The last is always "full". The capture's columns
# after the planted XML: after opening, after a first "full" ("-" when the steps start
# otherwise), then after each further step.
STEPS: dict[str, tuple[str, ...]] = {
    "name_static_dependent": ("edit:probe!A1=2", "full"),
    "name_no_ca_full_first": ("full", "edit:probe!C1=2", "full"),
    "own_name_ref_edit": ("edit:probe!A1=2", "full"),
    "second_indirect": ("edit:probe!C1=2", "edit:probe!A1=3", "full"),
    "second_offset": ("edit:probe!C1=2", "edit:probe!A1=3", "full"),
    "second_dormant": ("edit:probe!C1=2", "edit:probe!A1=3", "full"),
    "awake_unrelated": ("edit:probe!C1=2", "edit:probe!J1=5", "full"),
    "plain_and_volatile": ("edit:probe!A1=2", "full"),
    "manual_name": ("edit:probe!C1=2", "calc", "full"),
}
DEFAULT_STEPS = ("edit:probe!C1=2", "full")


def steps(variant: str) -> tuple[str, ...]:
    """The variant's steps after opening."""
    return STEPS.get(variant, DEFAULT_STEPS)


def columns(variant: str) -> int:
    """The shown values a capture row of the variant holds (after the planted XML)."""
    after = steps(variant)
    return 2 + len(after) - (after[0] == "full")


def edit(step: str) -> tuple[str, float]:
    """An ``edit:`` step's cell and number."""
    ref, number = step.removeprefix("edit:").split("=")
    return ref, float(number)
