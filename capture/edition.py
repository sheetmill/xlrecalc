"""The Excel edition and update channel a capture ran on, for its header (CAPTURING.md §2).

The build number alone does not say which functions Excel has: Office Home & Business 2021 for
Mac at 16.114.4 lacks ``LAMBDA`` and its helpers, which Microsoft 365 and Excel 2024 at the same
build have. No file on the Mac states the license plainly, so the edition is given by the person
capturing, in the environment variable ``XLRECALC_EXCEL_EDITION`` or in the local file
``capture/build/excel_edition.txt`` (not committed), and a capture refuses to run without it.
The update channel is Microsoft AutoUpdate's setting. Standard library only, so that any probe
can import it.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

EDITION_FILE = Path(__file__).resolve().parent / "build" / "excel_edition.txt"


def excel_edition() -> str:
    """The edition, as the person capturing states it; exits when it is not stated."""
    edition = os.environ.get("XLRECALC_EXCEL_EDITION", "").strip()
    if not edition and EDITION_FILE.exists():
        edition = EDITION_FILE.read_text(encoding="utf-8").strip()
    if not edition:
        sys.exit(
            "state the Excel edition first (Excel > About Excel), for example\n"
            f"  echo 'Office Home & Business 2021' > {EDITION_FILE}\n"
            "or set XLRECALC_EXCEL_EDITION"
        )
    return edition


def update_channel() -> str:
    """Microsoft AutoUpdate's channel (``Current``, ``Preview``, ``Beta``), or ``unknown``."""
    try:
        result = subprocess.run(
            ["defaults", "read", "com.microsoft.autoupdate2"],
            capture_output=True,
            text=True,
            timeout=10,
        )
    except (OSError, subprocess.TimeoutExpired):
        return "unknown"
    for line in result.stdout.splitlines():
        key, _, value = line.strip().partition("=")
        if key.strip() == "ChannelName":
            return value.strip().strip(";").strip().strip('"') or "unknown"
    return "unknown"


def edition_line() -> str:
    """The header line: ``# edition: <edition>, update channel <channel>``."""
    return f"# edition: {excel_edition()}, update channel {update_channel()}"
