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
}
