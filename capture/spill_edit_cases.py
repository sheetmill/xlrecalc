"""The cases of capture ``spill_edit``: an edit that makes a dynamic array spill further or less.

Each case is one workbook: its cells (a number, a plain formula, or a dynamic-array formula, in
the file's form), the steps taken after opening (``edit:sheet!A1=<number>``, as
``edit_steps.applescript`` takes them) and the cells read after opening and after each step.
``PREDICTED`` holds what was expected before the capture ran: each read cell's value after
opening and after each step (``None`` for a blank cell).

Kept apart from the probe so that the tests read them without openpyxl or Excel.
"""

from __future__ import annotations

SEQ = "_xlfn.SEQUENCE"
ANC = "_xlfn.ANCHORARRAY"

Cell = float | tuple[str, str]  # a number; ("=", text) a formula; ("#", text) a dynamic array

CASES: dict[str, tuple[dict[str, Cell], tuple[str, ...], tuple[str, ...]]] = {
    # D1 spills SEQUENCE(E1); F1 reads a cell the spill reaches only once it grows; G1 and H1
    # read a range over it.
    "grow": (
        {
            "probe!E1": 2.0,
            "probe!D1": ("#", f"{SEQ}(E1)"),
            "probe!F1": ("=", "D3+1"),
            "probe!G1": ("=", "SUM(D1:D5)"),
            "probe!H1": ("=", "COUNT(D1:D5)"),
        },
        ("edit:probe!E1=3",),
        ("probe!D1", "probe!D2", "probe!D3", "probe!D4", "probe!F1", "probe!G1", "probe!H1"),
    ),
    "shrink": (
        {
            "probe!E1": 3.0,
            "probe!D1": ("#", f"{SEQ}(E1)"),
            "probe!F1": ("=", "D3+1"),
            "probe!G1": ("=", "SUM(D1:D5)"),
            "probe!H1": ("=", "COUNT(D1:D5)"),
        },
        ("edit:probe!E1=2",),
        ("probe!D1", "probe!D2", "probe!D3", "probe!D4", "probe!F1", "probe!G1", "probe!H1"),
    ),
    "grow_then_shrink": (
        {
            "probe!E1": 2.0,
            "probe!D1": ("#", f"{SEQ}(E1)"),
            "probe!F1": ("=", "D4+1"),
            "probe!G1": ("=", "SUM(D1:D5)"),
        },
        ("edit:probe!E1=4", "edit:probe!E1=1"),
        ("probe!D1", "probe!D2", "probe!D3", "probe!D4", "probe!D5", "probe!F1", "probe!G1"),
    ),
    # I1 spills D1's spill times 10 (D1#*10): it grows with D1; J1 and K1 read its new cells.
    "spill_ref": (
        {
            "probe!E1": 2.0,
            "probe!D1": ("#", f"{SEQ}(E1)"),
            "probe!I1": ("#", f"{ANC}(D1)*10"),
            "probe!J1": ("=", "I3+1"),
            "probe!K1": ("=", "SUM(I1:I5)"),
        },
        ("edit:probe!E1=3",),
        ("probe!I1", "probe!I2", "probe!I3", "probe!I4", "probe!J1", "probe!K1"),
    ),
    # D3 holds 7: growing to three cells blocks the spill (#SPILL!, ERROR.TYPE 9), and D2 is
    # left blank; shrinking back spills again.
    "blocked": (
        {
            "probe!E1": 2.0,
            "probe!D3": 7.0,
            "probe!D1": ("#", f"{SEQ}(E1)"),
            "probe!F1": ("=", "IFERROR(ERROR.TYPE(D1),0)"),
            "probe!G1": ("=", "D2+D3"),
        },
        ("edit:probe!E1=3", "edit:probe!E1=2"),
        ("probe!D2", "probe!D3", "probe!F1", "probe!G1"),
    ),
    # A formula on another sheet reads a cell the spill reaches only once it grows.
    "other_sheet": (
        {
            "probe!E1": 2.0,
            "probe!D1": ("#", f"{SEQ}(E1)"),
            "other!A1": ("=", "probe!D3*2"),
        },
        ("edit:probe!E1=3",),
        ("probe!D3", "other!A1"),
    ),
}

# Expected before the capture ran: after opening, then after each step.
PREDICTED: dict[str, dict[str, tuple[float | None, ...]]] = {
    "grow": {
        "probe!D1": (1, 1),
        "probe!D2": (2, 2),
        "probe!D3": (None, 3),
        "probe!D4": (None, None),
        "probe!F1": (1, 4),
        "probe!G1": (3, 6),
        "probe!H1": (2, 3),
    },
    "shrink": {
        "probe!D1": (1, 1),
        "probe!D2": (2, 2),
        "probe!D3": (3, None),
        "probe!D4": (None, None),
        "probe!F1": (4, 1),
        "probe!G1": (6, 3),
        "probe!H1": (3, 2),
    },
    "grow_then_shrink": {
        "probe!D1": (1, 1, 1),
        "probe!D2": (2, 2, None),
        "probe!D3": (None, 3, None),
        "probe!D4": (None, 4, None),
        "probe!D5": (None, None, None),
        "probe!F1": (1, 5, 1),
        "probe!G1": (3, 10, 1),
    },
    "spill_ref": {
        "probe!I1": (10, 10),
        "probe!I2": (20, 20),
        "probe!I3": (None, 30),
        "probe!I4": (None, None),
        "probe!J1": (1, 31),
        "probe!K1": (30, 60),
    },
    "blocked": {
        "probe!D2": (2, None, 2),
        "probe!D3": (7, 7, 7),
        "probe!F1": (0, 9, 0),
        "probe!G1": (9, 7, 9),
    },
    "other_sheet": {
        "probe!D3": (None, 3),
        "other!A1": (0, 6),
    },
}
