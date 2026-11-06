"""What Excel does with arrays over whole columns and the whole sheet, and what it costs it.

    uv run python -m capture.array_size_probe [--functions [--resume]]

Each case is a workbook made here, with cached values: sheet D holds two cells (A1 = 1, B2 = 2),
and sheet S holds in A1 either a legacy array formula declared over a range (``D!A1:A1048576*1``
over ``A1:A1048576``, then ten columns, then the whole sheet; its cached value in A1 only) or
``SUMPRODUCT`` over a reference to D's whole columns (``A:A``, ``A:J``, ``A:XFD``). Smallest
first, one at a time: Excel opens the workbook (alerts off), A1 is read, Excel recalculates
everything, A1 is read again, and the workbook is closed without saving. Each step has a time
limit; Excel's memory (its physical footprint, from ``ps``) is sampled while it runs. A step
past the time limit, or Excel past 8 GB, ends Excel and the run: what was measured so far is
written to ``tests/captures/array_size.tsv``.

With ``--functions`` the cases are functions over D's columns instead (``FUNCTION_CASES``), each
over one column, then ten, then all of them, cached value 7, written to
``tests/captures/array_functions.tsv``. A function that uses more than 0.5 GB of Excel's memory
over ten columns does not go on to all of them. ``--resume`` keeps the rows already in the file
and runs the others, without that test for the ten-column rows it keeps.

Excel may use a great deal of memory here: run it only with no workbook open in Excel (the probe
refuses otherwise), when no other capture or crafted-input run is going on (CAPTURING.md §2).
"""

from __future__ import annotations

import datetime
import subprocess
import sys
import threading
import time
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc"
OUTPUT = ROOT / "tests" / "captures" / "array_size.tsv"
FUNCTIONS_OUTPUT = ROOT / "tests" / "captures" / "array_functions.tsv"
GB = 10**9
MEMORY_LIMIT = 8 * GB
TEN_COLUMN_LIMIT = GB // 2
STEP_SECONDS = {"open": 180, "read": 60, "calc": 300, "close": 60}

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

CASES = [  # (name, formula in S!A1, declared array range or None, cached value), smallest first
    ("sumproduct A:A", "SUMPRODUCT(D!A:A*1)", None, "1"),
    ("array A1:A1048576", "D!A1:A1048576*1", "A1:A1048576", "1"),
    ("sumproduct A:J", "SUMPRODUCT(D!A:J*1)", None, "3"),
    ("array A1:J1048576", "D!A1:J1048576*1", "A1:J1048576", "1"),
    ("sumproduct A:XFD", "SUMPRODUCT(D!A:XFD*1)", None, "3"),
    ("array A1:XFD1048576", "D!A1:XFD1048576*1", "A1:XFD1048576", "1"),
]

# Each template over (a range, the range moved a column right): one column, ten, then all (A:XFD,
# or A:XFC where a range moved right must fit).
_SPANS = [("A:A", "B:B"), ("A:J", "B:K"), ("A:XFC", "B:XFD")]
_TEMPLATES = [
    'COUNTIF(D!{a},"")',
    'COUNTIFS(D!{a},"")',
    'COUNTIFS(D!{a},">0")',
    'COUNTIFS(D!{a},"<>1")',
    'COUNTIFS(D!{a},"",D!{a},"<>2")',
    'SUMIF(D!{a},"<>5")',
    'SUMIF(D!{a},"",D!{b})',
    'SUMIF(D!{a},"",D!B:B)',
    'SUMIFS(D!{a},D!{a},"<>5")',
    'SUMIFS(D!{a},D!{a},">0")',
    'SUMIFS(D!{b},D!{a},"")',
    "DSUM(D!{a},1,D!A1:A2)",
    "CORREL(D!{a},D!{a})",
    "SUMPRODUCT(ROW(1:1048576)*COLUMN({a}))",
]
FUNCTION_CASES = [  # (template, formula), smallest first
    (template, template.format(a="A:XFD" if a == "A:XFC" and "{b}" not in template else a, b=b))
    for a, b in _SPANS
    for template in _TEMPLATES
    if not (a == "A:XFC" and "ROW(" in template)  # never every pair of the whole sheet
]
# Not run: Excel took over two minutes on the same formula through SUMIF, and the run ended there.
NOT_RUN = {'SUMIFS(D!A:XFD,D!A:XFD,"<>5")'}


def workbook(path: Path, formula: str, ref: str | None, cached: str) -> None:
    """A minimal workbook with cached values, which Excel opens without recalculating."""
    text = escape(formula)
    f = f'<f t="array" ref="{ref}">{text}</f>' if ref else f"<f>{text}</f>"
    sheet_s = (
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1">{f}<v>{cached}</v></c>'
        "</row></sheetData></worksheet>"
    )
    sheet_d = (
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
        '<row r="2"><c r="B2"><v>2</v></c></row></sheetData></worksheet>'
    )
    office = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    rels_type = "application/vnd.openxmlformats-package.relationships+xml"
    sheets = "".join(
        f'<Override PartName="/xl/worksheets/sheet{k}.xml" ContentType="{office}.worksheet+xml"/>'
        for k in (1, 2)
    )
    parts = {
        "[Content_Types].xml": f'<Types xmlns="{CT}">'
        f'<Default Extension="rels" ContentType="{rels_type}"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{office}.sheet.main+xml"/>'
        f"{sheets}</Types>",
        "_rels/.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
        f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>'
        '<sheet name="S" sheetId="1" r:id="rId1"/><sheet name="D" sheetId="2" r:id="rId2"/>'
        '</sheets><calcPr calcId="191029"/></workbook>',
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/>'
        "</Relationships>",
        "xl/worksheets/sheet1.xml": sheet_s,
        "xl/worksheets/sheet2.xml": sheet_d,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, body in parts.items():
            z.writestr(name, body)


def excel_pid() -> int | None:
    out = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True, text=True)
    return int(out.stdout.split()[0]) if out.stdout.split() else None


def footprint(pid: int) -> int:
    """Excel's physical footprint in bytes, from footprint(1); its resident size if that fails."""
    out = subprocess.run(["footprint", "-p", str(pid)], capture_output=True, text=True)
    for line in out.stdout.splitlines():
        if "phys_footprint:" in line:
            number, unit = line.split("phys_footprint:")[1].split()[:2]
            return int(float(number) * {"B": 1, "KB": 1e3, "MB": 1e6, "GB": 1e9}[unit])
    rss = subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True)
    return int(rss.stdout.strip() or 0) * 1024


class Watch:
    """Samples Excel's memory while a step runs; ends Excel past the limit."""

    def __init__(self, pid: int) -> None:
        self.pid, self.peak, self.ended = pid, 0, False
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self) -> None:
        while not self._stop.is_set():
            size = footprint(self.pid)
            self.peak = max(self.peak, size)
            if size > MEMORY_LIMIT:
                subprocess.run(["kill", "-9", str(self.pid)], check=False)
                self.ended = True
                return
            time.sleep(0.2)

    def stop(self) -> None:
        self._stop.set()
        self._thread.join()


def step(name: str, *args: str) -> tuple[str, float, int, str]:
    """Run one step; (result, seconds, peak bytes, why the run must end or '')."""
    pid = excel_pid()
    assert pid is not None, "Excel is not running"
    watch = Watch(pid)
    start = time.monotonic()
    seconds = str(STEP_SECONDS[name])
    out = subprocess.run(
        ["gtimeout", "-k", "5", seconds, "osascript", str(HERE / "array_size.applescript"),
         name, *args],
        capture_output=True, text=True, check=False,
    )  # fmt: skip
    elapsed = time.monotonic() - start
    watch.stop()
    if watch.ended:
        return "", elapsed, watch.peak, "Excel passed 8 GB and was ended"
    if out.returncode == 124:
        subprocess.run(["kill", "-9", str(pid)], check=False)
        return "", elapsed, watch.peak, f"no answer in {seconds} s: Excel was ended"
    result = out.stdout.strip() or out.stderr.strip()[-120:]
    return result, elapsed, watch.peak, ""


def main() -> None:
    count = subprocess.run(
        ["osascript", str(HERE / "array_size.applescript"), "count"],
        capture_output=True, text=True, check=True, timeout=30,
    ).stdout.strip()  # fmt: skip
    if count != "0":
        raise SystemExit(f"{count} workbooks are open in Excel: close them first")
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True, text=True, check=True, timeout=30,
    ).stdout.strip()  # fmt: skip
    CONTAINER.mkdir(parents=True, exist_ok=True)
    functions = "--functions" in sys.argv[1:]
    cases = [(formula, formula, None, "7") for _, formula in FUNCTION_CASES] if functions else CASES
    templates = {formula: template for template, formula in FUNCTION_CASES}
    costly: set[str] = set()
    rows = []
    if "--resume" in sys.argv[1:]:
        rows = [
            line
            for line in FUNCTIONS_OUTPUT.read_text(encoding="utf-8").splitlines()
            if not line.startswith(("#", "case\t"))
        ]
    done = {row.split("\t")[0] for row in rows}
    for name, formula, ref, cached in cases:
        template = templates.get(formula, "")
        if name in done:
            continue
        if name in NOT_RUN:
            rows.append(f"{name}\tskipped: as slow as SUMIF over the whole sheet")
            continue
        if template in costly and "A:XF" in formula:
            rows.append(f"{name}\tskipped: over ten columns it used more than 0.5 GB")
            print(rows[-1], flush=True)
            continue
        path = CONTAINER / "xlrecalc_array_size.xlsx"
        workbook(path, formula, ref, cached)
        fields = [name]
        stop = ""
        pid = excel_pid()
        before = footprint(pid) if pid else 0
        top = 0
        for action, args in (
            ("open", (str(path),)), ("read", (path.name,)), ("calc", ()), ("read", (path.name,)),
            ("close", (path.name,)),
        ):  # fmt: skip
            result, seconds, peak, stop = step(action, *args)
            top = max(top, peak)
            fields.append(f"{action}: {result or '-'} ({seconds:.1f} s, {peak / GB:.2f} GB)")
            print(fields[-1], flush=True)
            if stop or result.startswith("error"):
                break
        path.unlink(missing_ok=True)
        if top - before > TEN_COLUMN_LIMIT and "A:J" in formula:
            costly.add(template)
        if stop:
            fields.append(stop)
        rows.append("\t".join(fields))
        print(rows[-1], flush=True)
        if stop:
            break
    if excel_pid() is not None:
        subprocess.run(["osascript", str(HERE / "array_size.applescript"), "restore"], check=False)
    header = [
        f"# suite: {'array_functions' if functions else 'array_size'} "
        f"(capture.array_size_probe); Excel {version} for Mac, en_US",
        f"# captured: {datetime.date.today()}; each step: result (seconds, Excel's peak "
        "physical footprint during the step); smallest case first, the run ends at a step "
        "past its time limit or Excel past 8 GB",
        "case\tsteps",
    ]
    output = FUNCTIONS_OUTPUT if functions else OUTPUT
    output.write_text("\n".join(header + rows) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
