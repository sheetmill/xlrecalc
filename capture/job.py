"""A capture suite as an Excel capture job, and the capture its result gives.

    uv run python -m capture.job write SUITE JOB_ID OUT_DIR
    uv run python -m capture.job read SUITE JOB_DIR RESULT_DIR

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/e3a5207c251929a8a19d713777b70ace3506fb95/docs/capture-jobs.md
``write`` makes the job folder, the manifest (``job.json``: open the suite's workbook, save a copy)
and the workbook in ``files/``, built as ``capture/pipeline.py`` builds it, with its zip entries'
times and its document properties fixed, so that writing it again gives the same bytes and the
SHA-256 the capture's header records. ``read`` takes the job folder and its result folder, checks
the manifest's and the inputs' SHA-256 against the result's, and writes
``tests/captures/windows/<suite>.tsv``: the cached value and type of each formula, read from the
saved copy with ``capture/pipeline.py``'s reader, under the result's ``header`` copied whole; the
manifest is copied beside it, byte for byte, as ``<suite>.job.json``.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from capture.pipeline import CAPTURES, COLUMNS, build, read_results, result_repr

WINDOWS = CAPTURES / "windows"
FIXED_TIME = "2026-01-01T00:00:00Z"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)


def manifest(suite: str, job_id: str) -> dict[str, object]:
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "xlrecalc",
        "purpose": f"capture suite {suite}: open the workbook, save a copy with the cached values",
        "limits": {"job_seconds": 3600, "step_seconds": 600},
        "audit": False,
        "workbooks": [
            {
                "file": f"{suite}.xlsx",
                "steps": [
                    {"op": "open"},
                    {"op": "save_as", "name": f"{suite}.saved.xlsx", "format": "xlsx"},
                ],
            }
        ],
    }


def workbook(suite: str, path: Path) -> None:
    """The suite's workbook, the same bytes every time it is written."""
    from capture.suites import SUITES

    spec = SUITES[suite]
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / "raw.xlsx"
        build(raw, spec.cases(), spec.date1904, spec.fixed)
        with zipfile.ZipFile(raw) as source, zipfile.ZipFile(path, "w") as out:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "docProps/core.xml":
                    text = data.decode("utf-8")
                    text = re.sub(
                        r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(<)",
                        rf"\g<1>{FIXED_TIME}\2",
                        text,
                    )
                    data = text.encode("utf-8")
                entry = zipfile.ZipInfo(info.filename, date_time=ZIP_TIME)
                entry.compress_type = zipfile.ZIP_DEFLATED
                entry.external_attr = 0o600 << 16
                out.writestr(entry, data)


def write(suite: str, job_id: str, out: Path) -> Path:
    """The job folder ``out/<job_id>``."""
    folder = out / job_id
    (folder / "files").mkdir(parents=True)
    (folder / "job.json").write_text(json.dumps(manifest(suite, job_id), indent=2) + "\n")
    workbook(suite, folder / "files" / f"{suite}.xlsx")
    return folder


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(suite: str, job: Path, result: Path) -> Path:
    """The capture from a job folder and its result folder (``result.json``, ``artifacts/``)."""
    from capture.suites import SUITES

    outcome = json.loads((result / "result.json").read_text(encoding="utf-8"))
    if outcome.get("status") != "ok":
        sys.exit(f"the job did not end ok: {outcome.get('status')!r}")
    header = outcome["header"]
    if header["manifest_sha256"] != _sha256(job / "job.json"):
        sys.exit("the result's manifest_sha256 is not this job's manifest")
    for name, entry in header["inputs"].items():
        if entry["sha256"] != _sha256(job / "files" / name):
            sys.exit(f"the result's input {name} is not this job's")
    cases = SUITES[suite].cases()
    results = read_results(result / "artifacts" / f"{suite}.saved.xlsx", len(cases))
    text = json.dumps(header, indent=1, sort_keys=True, ensure_ascii=False)
    lines = [
        f"# suite: {suite}",
        f"# generator: capture/suites.py:{SUITES[suite].cases.__name__}, capture/job.py",
        "# method: the job beside this file (" + suite + ".job.json) opens the workbook the",
        "#   generator writes and saves a copy; the cached value and type of A are read from the",
        "#   copy's XML as capture/pipeline.py reads them (n: number, str: text, b: logical,",
        "#   e: error)",
        "# header (the job's result, copied whole):",
        *("#   " + line for line in text.splitlines()),
        "# columns: " + " ".join(COLUMNS),
    ]
    for k, (case, got) in enumerate(zip(cases, results, strict=True), 1):
        kind, text = result_repr(got)
        lines.append(
            "\t".join([suite, case.tag, str(k), case.formula, repr(case.inputs), kind, text])
        )
    WINDOWS.mkdir(exist_ok=True)
    path = WINDOWS / f"{suite}.tsv"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    shutil.copyfile(job / "job.json", WINDOWS / f"{suite}.job.json")
    return path


def main(argv: list[str]) -> None:
    if argv[:1] == ["write"] and len(argv) == 4:
        print(write(argv[1], argv[2], Path(argv[3])))
    elif argv[:1] == ["read"] and len(argv) == 4:
        print(read(argv[1], Path(argv[2]), Path(argv[3])))
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
