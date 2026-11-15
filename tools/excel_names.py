"""Excel's worksheet function names, for telling a name Excel knows from one it does not.

    uv run python tools/excel_names.py        # rewrites src/xlrecalc/_functions/_excel_names.py

Source: ``capture/functions_list.py``, the names the ``formula_bar`` capture typed. ``evaluate``
raises for a function Excel has and the package does not compute, and gives ``#NAME?`` for any
other name, as Excel does. ``tests/test_evaluate.py`` checks the committed module against it.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from capture.functions_list import FUNCTIONS  # noqa: E402

MODULE = ROOT / "src" / "xlrecalc" / "_functions" / "_excel_names.py"


def module_text() -> str:
    lines = [
        '"""Excel\'s worksheet function names (from tools/excel_names.py)."""',
        "",
        "EXCEL_NAMES: frozenset[str] = frozenset(",
        "    {",
        *(f'        "{name}",' for name in sorted(FUNCTIONS)),
        "    }",
        ")",
        "",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    MODULE.write_text(module_text(), encoding="utf-8")
    print(f"{len(FUNCTIONS)} names -> {MODULE.relative_to(ROOT)}")
