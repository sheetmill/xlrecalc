"""The file-form prefix of each function, derived from captures (ARCHITECTURE.md §5.2).

    uv run python tools/prefix_table.py        # rewrites src/xlrecalc/_formula/_prefixes.py

Sources, all under ``tests/captures/``: ``formula_bar`` and ``formula_bar_ranges`` (the prefix
Excel's file form gives a call typed in the formula bar) and ``prefixes`` (the form Excel
computes when the formula bar could not settle it: with ``_xlfn.`` a value, without it
``#NAME?``). A function that needs none is left out. ``tests/test_fileform.py`` checks the
committed module against this derivation.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CAPTURES = ROOT / "tests" / "captures"
MODULE = ROOT / "src" / "xlrecalc" / "_formula" / "_prefixes.py"
_SAVED = re.compile(r"(_xlfn\._xlws\.|_xlfn\.)?([A-Z][A-Z0-9.]*)\(")


def _rows(name: str) -> list[list[str]]:
    text = (CAPTURES / f"{name}.tsv").read_text(encoding="utf-8")
    return [line.split("\t") for line in text.splitlines() if not line.startswith("#")]


def prefix_table() -> dict[str, str]:
    table: dict[str, str] = {}
    for suite in ("formula_bar", "formula_bar_ranges"):
        for tag, typed, saved in _rows(suite):
            if tag != "function":
                continue
            text = ast.literal_eval(saved)
            m = _SAVED.match(text)
            if text == "REJECTED" or m is None:
                continue
            name = ast.literal_eval(typed)[1:].split("(")[0]
            if m[2] == name and m[1]:
                table[name] = m[1]
    computed: dict[str, dict[str, str]] = {}
    for name, prefix, _formula, kind, value in _rows("prefixes"):
        computed.setdefault(name, {})[ast.literal_eval(prefix)] = "" if kind == "e" else value
    for name, forms in computed.items():
        if forms.get("_xlfn.") and not forms.get(""):
            table[name] = "_xlfn."
    return dict(sorted(table.items()))


def module_text() -> str:
    lines = [
        '"""The prefix each function\'s name takes in a file (from tools/prefix_table.py)."""',
        "",
        "PREFIXES: dict[str, str] = {",
        *(f'    "{name}": "{prefix}",' for name, prefix in prefix_table().items()),
        "}",
        "",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    MODULE.write_text(module_text(), encoding="utf-8")
    print(f"{len(prefix_table())} functions -> {MODULE.relative_to(ROOT)}")
