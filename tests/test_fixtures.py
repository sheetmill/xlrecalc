"""Committed documents carry no personal data that Office saves into a file."""

import re
import subprocess
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SUFFIXES = (".xlsx", ".docx", ".pptx")
# Excel saves the folder a workbook was saved in (x15ac:absPath in xl/workbook.xml) and the user
# who saved it (cp:lastModifiedBy in docProps/core.xml).
MARKERS = (b"absPath", b"lastModifiedBy")
# Home-directory paths, in any part: a cached CELL("filename") value, an external link, a
# saved path. Excel writes / and \ as they are, and XML may escape neither.
HOME = re.compile(rb"/Users/|/home/|[A-Za-z]:\\Users\\", re.I)
# The author Excel records (dc:creator in docProps/core.xml) is the user's name; a fixture may
# name only the tool that wrote it.
CREATOR = re.compile(rb"<dc:creator>([^<]*)</dc:creator>")
TOOLS = frozenset({b"", b"openpyxl", b"xlrecalc"})
# Elements of docProps/app.xml that name an organization or a person when filled in.
FILLED = re.compile(rb"<(Company|Manager)>[^<]+</")
# Parts holding comments name their authors: Excel's comments and threaded comments and their
# persons, Word's and PowerPoint's comments and comment authors.
COMMENTS = re.compile(r"(^|/)(comments?\d*|threadedComment\w*|persons?|commentAuthors)\.xml$", re.I)


def committed(suffixes: tuple[str, ...]) -> list[Path]:
    """Committed files with these suffixes: tracked ones, or under an sdist the files it holds."""
    try:
        listed = subprocess.run(
            ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True
        ).stdout.split()
        files = [ROOT / name for name in listed]
    except (OSError, subprocess.CalledProcessError):
        files = [p for p in ROOT.rglob("*") if ".venv" not in p.parts]
    return sorted(p for p in files if p.suffix.lower() in suffixes and p.is_file())


DOCUMENTS = committed(SUFFIXES)
# Parts taken out of a workbook Excel saved and committed alone, such as the rich-value parts of
# the dyn_* captures.
PARTS = [p for p in committed((".xml",)) if "captures" in p.parts]


def _check_part(name: str, data: bytes) -> None:
    for marker in MARKERS:
        assert marker not in data, f"{name} contains {marker.decode()}"
    assert not FILLED.search(data), f"{name} names a company or manager"
    assert not HOME.search(data), f"{name} holds a home-directory path"
    for creator in CREATOR.findall(data):
        assert creator.strip() in TOOLS, f"{name}: dc:creator {creator.decode()!r}"


@pytest.mark.parametrize("path", DOCUMENTS, ids=lambda p: p.relative_to(ROOT).as_posix())
def test_no_personal_data(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert "docProps/custom.xml" not in names, "custom document properties"
        for name in names:
            assert not COMMENTS.search(name), f"{name}: comments name their authors"
            _check_part(name, archive.read(name))


@pytest.mark.parametrize("path", PARTS, ids=lambda p: p.relative_to(ROOT).as_posix())
def test_no_personal_data_in_a_loose_part(path):
    assert not COMMENTS.search(path.name), f"{path.name}: comments name their authors"
    _check_part(path.name, path.read_bytes())


def test_loose_parts_are_found():
    assert any(p.name.startswith("dyn_") and p.name.endswith("_rdrichvalue.xml") for p in PARTS)


def test_a_person_as_creator_is_found():
    found = CREATOR.findall(b"<dc:creator>Jane Doe</dc:creator>")
    assert found == [b"Jane Doe"]
    assert found[0] not in TOOLS


def test_home_paths_are_found():
    assert HOME.search(b"<v>/Users/someone/Documents/[a.xlsx]Sheet1</v>")
    assert HOME.search(rb"<v>C:\Users\someone\a.xlsx</v>")
    assert not HOME.search(b"<v>Users</v>")
