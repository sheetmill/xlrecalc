"""Every sum of doubles in the package is the same on every Python version.

Python's ``sum`` compensates its additions from 3.12 (Neumaier) and does not on 3.11, so a sum of
doubles through it differs in the last bits between versions. The package adds doubles with
``_arith.plain_sum`` (left to right, as Excel adds) and keeps ``sum`` for integers and logicals,
which no version changes. This lists every ``sum`` call left in ``src/``: a new one must be a
count, added here, or a ``plain_sum``.
"""

from __future__ import annotations

import ast
import math
from pathlib import Path

from xlrecalc._arith import plain_sum

SRC = Path(__file__).parents[1] / "src" / "xlrecalc"

# (module, the call's text): each sums integers or logicals.
COUNTS = {
    ("_engine.py", "sum(len(s.templates) for s in self.sheets)"),
    ("_compare.py", 'sum(1 for ch in text if ch > "\\uffff")'),
    ("_functions/_statistical.py", 'sum(1 for v in numbers(args, ctx, direct_text="skip") if not isinstance(v, ErrorValue))'),  # noqa: E501
    ("_functions/_statistical.py", "sum(1 for _ in arg.cells())"),
    ("_functions/_statistical.py", "sum(len(row) for row in arg)"),
    ("_functions/_statistical.py", "sum(len(row) for row in arg.head)"),
    ("_functions/_statistical.py", 'sum(1 for _, _, v in ref.cells() if v is not BLANK and v != "")'),  # noqa: E501
    ("_functions/_statistical.py", "sum(1 for v in values if (v < x if order else v > x))"),
    ("_functions/_statistical.py", "sum(1 for v in values if v == x)"),
    ("_functions/_statistical.py", "sum(v < x for v in values)"),
    ("_functions/_reshape.py", "sum(len(p) for p in parts)"),
    ("_functions/_reshape.py", "sum(len(p[0]) for p in parts)"),
    ("_functions/_logical.py", "sum(values)"),  # XOR: logicals
    ("_functions/_math.py", "sum(1 for v in values if v is not BLANK)"),
    ("_functions/_math.py", "sum(values)"),  # MULTINOMIAL: integer parts
    ("_functions/_regression.py", "sum(null)"),  # logicals
    ("_functions/_database.py", "sum(type(v) is float for v in values)"),
    ("_functions/_text.py", 'sum("\\udc00" <= u <= "\\udfff" for u in units[:i])'),
    ("_functions/_datetime.py", "sum(1 for i in range(rest) if (first + i) % 7 not in off)"),
    ("_functions/_datetime.py", "sum(1 for h in hol if lo <= h <= hi and _day_of_week(h, ctx) not in off)"),  # noqa: E501
    ("_functions/_bonds.py", "sum(366 if _isleap(y) else 365 for y in range(y1, y2 + 1))"),
}  # fmt: skip


def _sum_calls() -> set[tuple[str, str]]:
    found = set()
    for path in SRC.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        for node in ast.walk(ast.parse(text)):
            if (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Name)
                and node.func.id == "sum"
            ):
                found.add((path.relative_to(SRC).as_posix(), ast.get_source_segment(text, node)))
    return found


def test_every_sum_left_counts_integers() -> None:
    assert _sum_calls() == COUNTS


def test_plain_sum_adds_left_to_right() -> None:
    # 1E16 + 1 + 1 is 1E16 added in order (each 1 is lost), 1E16 + 2 correctly rounded.
    values = [1e16, 1.0, 1.0]
    assert plain_sum(values) == 1e16
    assert math.fsum(values) == 1e16 + 2
    assert plain_sum([]) == 0.0
