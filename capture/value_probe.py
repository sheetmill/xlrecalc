"""What Excel does with number forms it never writes, and with rows and columns off its grid.

    uv run python -m capture.value_probe

Each case is a workbook openpyxl writes (A1 = 111, B1 = ``=A1*2``) with A1 rewritten in the
sheet's XML: its ``<v>`` replaced by a form such as ``nan``, ``INF``, ``1e400`` or ``0x10``, or
the cell (with its row) moved to a row or column past Excel's grid. Excel opens each with alerts
off and saves a copy (``save_copy.applescript``); a file it would offer to repair fails to open
with error -50. From the copy, the script reads A1's and B1's cached values and types. Output:
``tests/captures/value_forms.tsv``.

Run it only when Excel answers and no other capture or crafted-input run is going on
(CAPTURING.md §2). Workbooks live in Excel's container and are deleted afterwards.
"""

from __future__ import annotations

import datetime
import io
import re
import subprocess
import zipfile
from pathlib import Path

import openpyxl

from capture.edition import edition_line

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc"
OUTPUT = ROOT / "tests" / "captures" / "value_forms.tsv"
SHEET = "xl/worksheets/sheet1.xml"
TIME_LIMIT = "60"

VALUES = [
    "nan", "NaN", "inf", "INF", "-INF", "Infinity", "1e400", "-1e400", "1e308", "4.9e-324",
    "1_000", "0x10", " 5", "5 ", "+5", "5.", ".5", "1E+2", "1e", "",
]  # fmt: skip
PLACES = {
    "row 0": ("0", "A0"),
    "row 1048577": ("1048577", "A1048577"),
    "row 99999999999": ("99999999999", "A99999999999"),
    "column XFE": ("1", "XFE1"),
    "row 0 without a cell reference": ("0", None),
}


def base() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 111
    ws["B1"] = "=A1*2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def rewrite(data: bytes, edit: object) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            body = src.read(info)
            if info.filename == SHEET:
                body = edit(body.decode()).encode()  # type: ignore[operator]
            dst.writestr(info, body)
    return out.getvalue()


def with_value(value: str):
    def edit(xml: str) -> str:
        new, n = re.subn(r'(<c r="A1"[^>]*>)<v>111</v>', rf"\g<1><v>{value}</v>", xml)
        assert n == 1, xml
        return new

    return edit


def moved(row: str, ref: str | None):
    def edit(xml: str) -> str:
        cell = re.search(r'<c r="A1"[^>]*>.*?</c>', xml)
        assert cell is not None
        a1 = cell.group(0) if ref is None else cell.group(0).replace('r="A1"', f'r="{ref}"')
        if ref is None:
            a1 = a1.replace(' r="A1"', "")
        xml = xml.replace(cell.group(0), "")
        return xml.replace("</sheetData>", f'<row r="{row}">{a1}</row></sheetData>')

    return edit


def cases() -> dict[str, bytes]:
    data = base()
    out = {"control": data}
    for value in VALUES:
        out[f"v {value!r}"] = rewrite(data, with_value(value))
    for name, (row, ref) in PLACES.items():
        out[name] = rewrite(data, moved(row, ref))
    return out


def save_copy(src: Path, copy: Path) -> str:
    """Excel opens ``src`` and saves ``copy``; '' or the error number Excel gave."""
    result = subprocess.run(
        ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(HERE / "save_copy.applescript"),
         str(src), str(copy)],
        capture_output=True, text=True, check=False,
    )  # fmt: skip
    if result.returncode == 0:
        return ""
    found = re.search(r"\((-?\d+)\)", result.stderr)
    return found.group(1) if found else result.stderr.strip()[:80]


def cell(xml: str, ref: str) -> str:
    found = re.search(rf'<c r="{ref}"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
    if found is None:
        return "absent"
    t = re.search(r'\bt="([^"]*)"', found.group(1))
    v = re.search(r"<v>(.*?)</v>", found.group(2) or "", re.S)
    return f"{t.group(1) if t else 'n'}:{v.group(1) if v else ''}"


def main() -> None:
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True, text=True, check=True, timeout=30,
    ).stdout.strip()  # fmt: skip
    CONTAINER.mkdir(parents=True, exist_ok=True)
    rows = []
    for name, data in cases().items():
        src, copy = CONTAINER / "xlrecalc_value_src.xlsx", CONTAINER / "xlrecalc_value_copy.xlsx"
        src.write_bytes(data)
        copy.unlink(missing_ok=True)
        error = save_copy(src, copy)
        if error:
            rows.append(f"{name}\trefused\t{error}\t\t")
        else:
            with zipfile.ZipFile(copy) as z:
                xml = z.read(SHEET).decode()
            rows.append(f"{name}\topened\t\t{cell(xml, 'A1')}\t{cell(xml, 'B1')}")
        src.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
        print(rows[-1], flush=True)
    header = [
        f"# suite: value_forms (capture.value_probe); Excel {version} for Mac, en_US",
        edition_line(),
        f"# captured: {datetime.date.today()}; method: openpyxl's workbook with A1 rewritten in "
        "the sheet XML, opened with alerts off and saved as a copy",
        "# A1 = 111 rewritten as the case says, B1 = A1*2; values as type:value from Excel's copy; "
        "refused = error -50",
        "case\toutcome\terror\tA1\tB1",
    ]
    OUTPUT.write_text("\n".join(header + rows) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
