"""Text captures hold no control character but tab and newline, C1 included (gzipped: unpacked).

git and GitHub take a file holding one for binary; a generator writes a value it cannot name
in readable form (``capture/pipeline.py`` writes inputs and results as Python ``repr``). The
captures are the text files the generators write (``.tsv``, ``.json``, ``.txt``); the ``.xml``
and ``.xlsx`` files beside them are parts and files Excel saved, kept byte for byte (an XML
declaration Excel ends with a carriage return included).
"""

from __future__ import annotations

import gzip
import re
from pathlib import Path

import pytest

ROOT = Path(__file__).parent / "captures"
TEXT = (".tsv", ".json", ".txt")
CAPTURES = sorted(
    p
    for p in ROOT.rglob("*")
    if p.suffix in TEXT or p.name.endswith(tuple(t + ".gz" for t in TEXT))
)
CONTROL = re.compile(r"[\x00-\x08\x0b-\x1f\x7f-\x9f]")
"""Control characters other than tab and newline (carriage return and the C1 controls included)."""


def _found(data: bytes) -> list[tuple[int, str]]:
    return [
        (number, line[:60])
        for number, line in enumerate(data.decode("utf-8").split("\n"), 1)  # not splitlines
        if CONTROL.search(line)
    ]


def test_the_guard_finds_each_control_character() -> None:
    """Each control character but tab and newline, U+0085 and others splitlines splits at too."""
    for code in [*range(0x00, 0x09), *range(0x0B, 0x20), *range(0x7F, 0xA0)]:
        line = f"a{chr(code)}b"
        assert _found(f"a\tb\n{line}\n".encode()) == [(2, line)]
    assert _found("a\tb\nc\u00a0d\u2028e\u2029f\n".encode()) == []


def test_every_text_capture_is_checked() -> None:
    assert len(CAPTURES) > 100
    assert any(p.parent.name == "windows" for p in CAPTURES)


@pytest.mark.parametrize("path", CAPTURES, ids=lambda p: str(p.relative_to(ROOT)))
def test_a_text_capture_holds_no_control_character(path: Path) -> None:
    data = path.read_bytes()
    if path.suffix == ".gz":
        data = gzip.decompress(data)
    found = _found(data)
    assert not found, found
