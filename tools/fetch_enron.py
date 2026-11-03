"""Fetch the Enron spreadsheet corpus for local analysis, pinned to its URL and checksums.

    uv run python tools/fetch_enron.py [DESTINATION]      # default: ../corpora/enron

The corpus (F. Hermans and E. Murphy-Hill, MSR 2015; figshare,
doi:10.6084/m9.figshare.1221767, CC BY 4.0) is about 1 GB compressed and holds real business
workbooks with personal data, so it is never committed: it lives outside the repository, is
read only with this project's own reader, and is never opened in Excel (``CAPTURING.md`` §4).
The script downloads ``spreadsheets.7z``, checks its size, MD5 (as figshare publishes it) and
SHA-256, refuses an archive that names absolute paths or ``..``, and extracts only the
``.xlsx`` files into ``DESTINATION/spreadsheets``. It needs a 7-Zip command (``7zz`` or ``7z``,
for example ``brew install sevenzip``). ``tools/function_frequency.py`` and
``tools/formula_roundtrip.py`` then read the extracted files.
"""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import sys
import urllib.request
from pathlib import Path

URL = "https://ndownloader.figshare.com/files/3242531"
SIZE = 992_635_888
MD5 = "188ec98d04c83a4725d9205cc75ce5a2"
SHA256 = "d73bd1ea5be929236541331a58ef8e255857fb8f7b2397ad629a30425bd9e149"
DEFAULT = Path(__file__).resolve().parents[2] / "corpora" / "enron"


def verified(path: Path) -> bool:
    if not path.exists() or path.stat().st_size != SIZE:
        return False
    md5, sha = hashlib.md5(usedforsecurity=False), hashlib.sha256()
    with path.open("rb") as f:
        while chunk := f.read(1 << 20):
            md5.update(chunk)
            sha.update(chunk)
    return md5.hexdigest() == MD5 and sha.hexdigest() == SHA256


def download(path: Path) -> None:
    partial = path.with_suffix(".part")
    print(f"downloading {URL} ({SIZE / 1e9:.2f} GB) to {path}")
    with urllib.request.urlopen(URL) as response, partial.open("wb") as out:
        shutil.copyfileobj(response, out, 1 << 20)
    partial.replace(path)


def main() -> int:
    destination = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    destination.mkdir(parents=True, exist_ok=True)
    archive = destination / "spreadsheets.7z"
    if not verified(archive):
        download(archive)
        if not verified(archive):
            print("the download does not match the pinned size and checksums", file=sys.stderr)
            return 1
    seven = shutil.which("7zz") or shutil.which("7z")
    if seven is None:
        print("no 7-Zip command found (7zz or 7z)", file=sys.stderr)
        return 1
    listing = subprocess.run(
        [seven, "l", "-slt", str(archive)], capture_output=True, text=True, check=True
    ).stdout
    paths = [line[7:] for line in listing.splitlines() if line.startswith("Path = ")][1:]
    if any(p.startswith(("/", "\\")) or ".." in p.replace("\\", "/").split("/") for p in paths):
        print("the archive names a path outside its directory; not extracting", file=sys.stderr)
        return 1
    target = destination / "spreadsheets"
    subprocess.run(
        [seven, "x", "-y", "-bso0", "-bsp0", f"-o{target}", str(archive), "*.xlsx"], check=True
    )
    count = sum(1 for _ in target.rglob("*.xlsx"))
    print(f"{count} .xlsx files in {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
