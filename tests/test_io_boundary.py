"""xlsx reading and writing go through ``xlio``'s public names only (ARCHITECTURE.md §9.1).

``xlio`` keeps its modules private (``xlio._...``); anything else would tie the package, its
tooling or its tests to internals that may change in any release of ``xlio``.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
import xlio

ROOT = Path(__file__).resolve().parents[1]
FILES = sorted(
    p
    for folder in ("src", "capture", "tools", "tests", "bench")
    for p in (ROOT / folder).rglob("*.py")
    if ".venv" not in p.parts and "build" not in p.parts
)
PUBLIC = set(xlio.__all__)


def uses(tree: ast.AST) -> list[str]:
    """Every ``xlio`` module imported and every name taken from it, as written."""
    aliases = {"xlio"}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                if a.name == "xlio" or a.name.startswith("xlio."):
                    aliases.add(a.asname or "xlio")
    out = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            out += [a.name for a in node.names if a.name.startswith("xlio.")]
        elif (
            isinstance(node, ast.ImportFrom) and node.module and node.module.split(".")[0] == "xlio"
        ):
            if node.module != "xlio":
                out.append(node.module)
            else:
                out += [f"xlio.{a.name}" for a in node.names if a.name not in PUBLIC]
        elif (
            isinstance(node, ast.Attribute)
            and isinstance(node.value, ast.Name)
            and node.value.id in aliases
            and node.attr not in PUBLIC
            and not node.attr.startswith("__")
        ):
            out.append(f"xlio.{node.attr}")
    return out


def test_the_rule_finds_a_private_name():
    tree = ast.parse("import xlio\nfrom xlio._write import write\nfrom xlio import _x\nxlio._book")
    assert uses(tree) == ["xlio._write", "xlio._x", "xlio._book"]


@pytest.mark.parametrize("path", FILES, ids=lambda p: str(p.relative_to(ROOT)))
def test_only_public_names(path):
    assert uses(ast.parse(path.read_text(encoding="utf-8"))) == []


def test_no_reader_of_its_own():
    assert not (ROOT / "src" / "xlrecalc" / "_xlsx").exists()
