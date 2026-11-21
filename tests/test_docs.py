"""The design document's references resolve: section numbers and names in the package."""

import ast
import re
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "src" / "xlrecalc"
ARCHITECTURE = (ROOT / "ARCHITECTURE.md").read_text(encoding="utf-8")
# Two levels at most in ARCHITECTURE.md; a reference with three or more parts, such as
# ECMA-376's §18.8.19, cites a standard.
SECTION = re.compile(r"§(\d+(?:\.\d+)?[a-z]?)(?![\d.]*\d)")


def tracked() -> list[Path]:
    """The repository's text files: the tracked ones, or under an sdist the files it holds."""
    try:
        listed = subprocess.run(
            ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True
        ).stdout.split()
        files = [ROOT / name for name in listed]
    except (OSError, subprocess.CalledProcessError):
        files = [p for p in ROOT.rglob("*") if ".venv" not in p.parts and "_build" not in p.parts]
    suffixes = {".md", ".py", ".applescript", ".yml", ".yaml", ".toml", ".html"}
    return sorted(p for p in files if p.suffix in suffixes and p.is_file())


def headings() -> set[str]:
    found = set()
    for line in ARCHITECTURE.splitlines():
        if match := re.match(r"#{2,3} (\d+(?:\.\d+)?[a-z]?)\.? ", line):
            found.add(match[1])
    return found


def test_section_references_resolve():
    known = headings()
    assert {"4.4", "5.4", "9.4", "12"} <= known
    dangling = [
        f"{path.relative_to(ROOT)}:{number}: §{ref}"
        for path in tracked()
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1)
        for ref in SECTION.findall(line)
        if ref not in known
    ]
    assert dangling == []


ROADMAP_ITEM = re.compile(r"(?:§13|[Rr]oadmap),?\s+items?\s+\d")


def test_roadmap_items_are_cited_by_subject():
    """The Roadmap's items are renumbered as they ship, so a reference names the item's subject."""
    cited = [
        f"{path.relative_to(ROOT)}: {match[0]!r}"
        for path in tracked()
        for match in ROADMAP_ITEM.finditer(path.read_text(encoding="utf-8"))
    ]
    assert cited == []


def names(module: Path) -> set[str]:
    """Top-level names of a module, and ``Class.member`` for each class member."""
    tree = ast.parse(module.read_text(encoding="utf-8"))
    found = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
            found.add(node.name)
        if isinstance(node, ast.ClassDef):
            found |= {
                f"{node.name}.{item.name}"
                for item in node.body
                if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef))
            }
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            found |= {t.id for t in targets if isinstance(t, ast.Name)}
    return found


# Excel's own function-name prefixes look like module references; they are not.
EXCEL_PREFIXES = {"_xlfn", "_xlws", "_xludf", "_xlpm"}


def _docstrings(path: Path) -> str:
    """Every docstring of a Python file, the module's included."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    nodes = [tree, *ast.walk(tree)]
    kinds = (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)
    return "\n".join(ast.get_docstring(n) or "" for n in nodes if isinstance(n, kinds))


def _reference_texts() -> dict[str, str]:
    """The texts whose module references are checked: the design documents and docstrings."""
    texts = {"ARCHITECTURE.md": ARCHITECTURE}
    for doc in ("CAPTURING.md", "QUIRKS.md"):
        if (ROOT / doc).is_file():
            texts[doc] = (ROOT / doc).read_text(encoding="utf-8")
    for path in [*sorted((ROOT / "capture").glob("*.py")), *sorted(PACKAGE.rglob("*.py"))]:
        texts[str(path.relative_to(ROOT))] = _docstrings(path)
    return texts


_TICK = r"``?"
# `_module.name` (or `_module.py`), `NAME` in `_module.py`, and ``NAME`` (``_module.py``)
_DOTTED = re.compile(rf"{_TICK}(_[a-z_]+)\.(\w+(?:\.\w+)?){_TICK}")
_IN_MODULE = re.compile(
    rf"{_TICK}([A-Za-z_]\w*(?:\.\w+)?){_TICK}(?: in | \(){_TICK}(_[a-z_]+)\.py{_TICK}"
)


def _references() -> list[tuple[str, str, str]]:
    found = set()
    for source, text in _reference_texts().items():
        found |= {(source, module, name) for module, name in _DOTTED.findall(text)}
        found |= {(source, module, name) for name, module in _IN_MODULE.findall(text)}
    return sorted(ref for ref in found if ref[1] not in EXCEL_PREFIXES)


def _module_files(module: str) -> list[Path]:
    return sorted(PACKAGE.rglob(f"{module}.py"))


@pytest.mark.parametrize("reference", _references(), ids=lambda r: f"{r[0]}:{r[1]}.{r[2]}")
def test_module_references_resolve(reference):
    """A module a document or docstring cites exists and holds the name it is cited for."""
    source, module, name = reference
    files = _module_files(module)
    assert files, f"{source}: no module {module}.py"
    if name != "py":
        assert any(name in names(f) for f in files), f"{source}: {name} is not in {module}.py"


def test_cited_public_names_exist():
    """Outside the Roadmap, ARCHITECTURE.md presents only API that exists.

    Private modules (``xlrecalc._engine``) are named for the design, not as API.
    """
    import xlrecalc

    current = ARCHITECTURE.split("\n## 13. Roadmap", 1)[0]
    cited = set(re.findall(r"\bxlrecalc\.([A-Za-z]\w*)", current))
    assert cited
    assert cited <= set(xlrecalc.__all__), sorted(cited - set(xlrecalc.__all__))


def test_every_benchmark_result_is_described():
    """Every file under bench/results/ is named in its SUMMARY.md."""
    results = ROOT / "bench" / "results"
    if not results.is_dir():
        pytest.skip("bench/ is not in the sdist")
    summary = (results / "SUMMARY.md").read_text(encoding="utf-8")
    missing = [p.name for p in results.iterdir() if p.is_file() and p.name != "SUMMARY.md"]
    assert [name for name in missing if f"`{name}`" not in summary] == []


def test_no_registered_function_is_called_not_implemented():
    """A sentence of ARCHITECTURE.md or QUIRKS.md saying "not implemented" names no function the
    registry has: such a sentence goes stale when the function ships."""
    from xlrecalc._functions import SPECS

    registered = {k.upper().replace("_XLFN.", "").replace("_XLWS.", "") for k in SPECS}
    stale = []
    for name in ("ARCHITECTURE.md", "QUIRKS.md"):
        text = (ROOT / name).read_text(encoding="utf-8")
        for sentence in re.split(r"(?<=[.;:])\s", text):
            if "not implemented" in sentence:  # the names said to be so come before it
                before = sentence[: sentence.index("not implemented")]
                cited = re.findall(r"`(?:_xlfn\.)?([A-Z][A-Z0-9.]*[A-Z0-9])`", before)
                stale += [(name, c) for c in cited if c in registered]
    assert not stale, stale
