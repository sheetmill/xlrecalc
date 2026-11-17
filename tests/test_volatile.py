"""The volatile functions the documents list are the ones the registry marks volatile."""

from __future__ import annotations

import re
from pathlib import Path

from xlrecalc._functions import SPECS

ROOT = Path(__file__).resolve().parents[1]


def _registry() -> set[str]:
    """The registry's volatile functions as the documents name them (no file prefix)."""
    return {name.rsplit(".", 1)[-1] if name.startswith("_XLFN.") else name
            for name, spec in SPECS.items() if spec.volatile}  # fmt: skip


def _names(text: str) -> set[str]:
    return set(re.findall(r"`([A-Z][A-Z.]*)`", text))


def test_architecture_lists_the_registry_volatile_functions() -> None:
    text = (ROOT / "ARCHITECTURE.md").read_text(encoding="utf-8")
    paragraph = re.search(r"\*\*Volatility\.\*\*(.*?)\n\n", text, re.S)
    assert paragraph, "ARCHITECTURE.md: the Volatility paragraph is missing"
    listed = re.search(r"marks volatile, as in Excel, are (.*?)\(", paragraph[1], re.S)
    assert listed
    assert _names(listed[1]) == _registry()


def test_the_glossary_lists_the_registry_volatile_functions() -> None:
    text = (ROOT / "docs" / "glossary.md").read_text(encoding="utf-8")
    entry = re.search(r"\nVolatile function\n: (.*?)\n\n", text, re.S)
    assert entry, "docs/glossary.md: the Volatile function entry is missing"
    listed = entry[1].split(":", 1)[1].split(". ", 1)[0]
    assert _names(listed) == _registry()


def test_no_other_list_of_volatile_functions() -> None:
    """Elsewhere, a mention points to the list: no sentence on volatility names three or more."""
    volatile = _registry()
    for path in [ROOT / "ARCHITECTURE.md", ROOT / "README.md", ROOT / "QUIRKS.md"]:
        text = path.read_text(encoding="utf-8")
        text = re.sub(r"\*\*Volatility\.\*\*.*?\n\n", "", text, flags=re.S)
        for sentence in re.split(r"(?<=[.;])\s", text):
            if "volatile" in sentence.lower():
                named = _names(sentence) & volatile
                assert len(named) < 3, f"{path.name}: {sorted(named)} beside 'volatile'"
