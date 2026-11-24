"""Excel's worksheet function names, for telling a name Excel knows from one it does not.

    uv run python tools/excel_names.py        # rewrites src/xlrecalc/_functions/_excel_names.py

Source: ``capture/functions_list.py``, the names the ``formula_bar`` capture typed, except those in
``NOT_EXCELS``. ``evaluate`` raises for a function Excel has and the package does not compute, and
gives ``#NAME?`` for any other name, as Excel does. ``tests/test_evaluate.py`` checks the committed
module against it.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from capture.functions_list import FUNCTIONS  # noqa: E402

MODULE = ROOT / "src" / "xlrecalc" / "_functions" / "_excel_names.py"

# Names of the list that no en-US Excel knows: JIS gives #NAME? with and without _xlfn., on the Mac
# and on Windows (capture function_names, cases 265 and 266, and windows/function_names).
NOT_EXCELS = frozenset({"JIS"})


def module_text() -> str:
    lines = [
        '"""Excel\'s worksheet function names (from tools/excel_names.py)."""',
        "",
        "EXCEL_NAMES: frozenset[str] = frozenset(",
        "    {",
        *(f'        "{name}",' for name in sorted(set(FUNCTIONS) - NOT_EXCELS)),
        "    }",
        ")",
        "",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    MODULE.write_text(module_text(), encoding="utf-8")
    print(f"{len(set(FUNCTIONS) - NOT_EXCELS)} names -> {MODULE.relative_to(ROOT)}")
