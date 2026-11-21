"""Capture ``volatile_edit``'s variants, read by its probe and by ``tests/test_volatile.py``.

Kept apart from the probe so that reading them imports neither openpyxl nor Excel's tooling.
"""

from __future__ import annotations

READ = 'INDIRECT("probe!C1")'

# name: (B1's formula, defined name text or None, B1 marked ca="1", extra cells {ref: formula})
VARIANTS: dict[str, tuple[str, str | None, bool, dict[str, str]]] = {
    "name_no_ca": ("=vn", READ, False, {}),
    "name_ca": ("=vn", READ, True, {}),
    "own_no_ca": ("=" + READ, None, False, {}),
    "now_no_ca": ("=NOW()*0+A1", None, False, {"F1": "=NOW()*0+3"}),
    "name_dependent_no_ca": ("=vn", "OFFSET(probe!$C$1,0,0)", False, {}),
    "plain_name": ("=vn+0", "probe!$C$1", False, {}),
    "name_offset_plus": ("=vn", "OFFSET(probe!$C$1,0,0)+probe!$C$1*0", False, {}),
    "name_static_dependent": ("=vn", READ + "+probe!$A$1", False, {}),
    "name_no_ca_full_first": ("=vn", READ, False, {}),
    "name_no_ca_full_on_load": ("=vn", READ, False, {}),
}
EDITED = {"name_static_dependent": "probe!A1"}  # the cell the edit sets to 2; else probe!C1
FULL_FIRST = {"name_no_ca_full_first"}  # a full calculation, and a reading, before the edit
FULL_ON_LOAD = {"name_no_ca_full_on_load"}  # <calcPr fullCalcOnLoad="1"> written in the file
