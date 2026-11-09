"""A copy of an Excel-saved workbook without the folder it was saved in or the user who saved it.

Excel writes both into a file it saves (``x15ac:absPath`` in ``xl/workbook.xml``,
``cp:lastModifiedBy`` in ``docProps/core.xml``); a workbook committed as evidence carries neither
(``tests/test_fixtures.py``). Standard library only.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path


def strip(source: Path, target: Path) -> None:
    """``source`` without the folder it was saved in or the user who saved it."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = re.sub(
                    rb"<mc:AlternateContent[^>]*>.*?absPath.*?</mc:AlternateContent>",
                    b"",
                    data,
                    flags=re.S,
                )
                data = re.sub(rb"<x15ac:absPath[^>]*/>", b"", data)
            if item.filename == "docProps/core.xml":
                data = re.sub(
                    rb"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>|<cp:lastModifiedBy/>", b"", data
                )
            out.writestr(item, data)
