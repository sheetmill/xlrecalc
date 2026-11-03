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


@pytest.mark.parametrize(
    "reference",
    sorted(
        ref
        for ref in set(re.findall(r"`(_[a-z_]+)\.(\w+(?:\.\w+)?)`", ARCHITECTURE))
        if ref[0] not in EXCEL_PREFIXES
    ),
)
def test_module_references_resolve(reference):
    module, name = reference
    if name == "py":
        assert (PACKAGE / f"{module}.py").is_file(), f"{module}.py"
    else:
        assert name in names(PACKAGE / f"{module}.py"), f"{module}.{name}"


def test_cited_public_names_exist():
    """Outside the Roadmap, ARCHITECTURE.md presents only API that exists.

    Private modules (``xlrecalc._xlsx``) are named for the design, not as API.
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
