"""Every capture tool records the Excel edition beside the build (``CAPTURING.md`` §2)."""

from __future__ import annotations

import re
from pathlib import Path

CAPTURE = Path(__file__).parents[1] / "capture"


def test_every_header_with_the_build_has_the_edition():
    missing = []
    for path in sorted(CAPTURE.glob("*.py")):
        text = path.read_text(encoding="utf-8")
        builds = len(re.findall(r"# excel: |Excel \{version\} for Mac", text))
        if builds and len(re.findall(r"edition_line\(\)", text)) < builds:
            missing.append(path.name)
    assert missing == []


def test_the_edition_line_names_the_edition_and_channel(monkeypatch):
    from capture import edition

    monkeypatch.setenv("XLRECALC_EXCEL_EDITION", "Microsoft 365")
    monkeypatch.setattr(edition, "update_channel", lambda: "Current")
    assert edition.edition_line() == "# edition: Microsoft 365, update channel Current"
