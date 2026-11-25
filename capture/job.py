"""A capture suite as an Excel capture job, and the capture its result gives.

    uv run python -m capture.job write SUITE JOB_ID OUT_DIR
    uv run python -m capture.job read SUITE JOB_DIR RESULT_DIR

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/3adb865500f4d2223e5c1a758b466e081d3fdc79/docs/capture-jobs.md
``write`` makes the job folder, the manifest (``job.json``: open the suite's workbook, save a copy)
and the workbook in ``files/``, built as ``capture/pipeline.py`` builds it, with its zip entries'
times and its document properties fixed, so that writing it again gives the same bytes and the
SHA-256 the capture's header records. ``read`` takes the job folder and its result folder, checks
the manifest's and the inputs' SHA-256 against the result's, and writes
``tests/captures/windows/<suite>.tsv``: the cached value and type of each formula, read from the
saved copy with ``capture/pipeline.py``'s reader, under the result's ``header`` copied whole; the
manifest is copied beside it, byte for byte, as ``<suite>.job.json``.

A suite named ``dyn_<name>`` is one of ``capture/dynamic_suites.py``: its workbook is built as
``capture/dynamic_probe.py`` builds it, and its capture records every cell of the saved copy as
that probe does, with the copy's rich-value parts beside it.
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


REQUIRES: dict[str, object] = {
    "excel_build_min": {"windows": "16.0.20527.0"},
    "edition": "Microsoft 365",
}
"""What a new job requires: the edition decides which functions Excel has, and the build is the
first one the Windows captures were made with."""


def manifest(
    suite: str, job_id: str, requires: dict[str, object] | None = REQUIRES
) -> dict[str, object]:
    """The manifest; ``requires=None`` gives the manifest of a job written before manifests
    carried ``requires``."""
    head: dict[str, object] = {
        "schema": 1,
        "job_id": job_id,
        "package": "xlrecalc",
        "purpose": f"capture suite {suite}: open the workbook, save a copy with the cached values",
    }
    if requires is not None:
        head["requires"] = requires
    return {
        **head,
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
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / "raw.xlsx"
        if suite.startswith("dyn_"):
            from capture.dynamic_probe import build as build_dynamic

            build_dynamic(raw, suite[len("dyn_") :])
        else:
            from capture.suites import SUITES

            spec = SUITES[suite]
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


def write(
    suite: str, job_id: str, out: Path, requires: dict[str, object] | None = REQUIRES
) -> Path:
    """The job folder ``out/<job_id>``."""
    folder = out / job_id
    (folder / "files").mkdir(parents=True)
    text = json.dumps(manifest(suite, job_id, requires), indent=2) + "\n"
    (folder / "job.json").write_text(text)
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
    text = json.dumps(header, indent=1, sort_keys=True, ensure_ascii=False)
    WINDOWS.mkdir(exist_ok=True)
    if suite.startswith("dyn_"):
        return _read_dynamic(suite, job, result, text)
    cases = SUITES[suite].cases()
    results = read_results(result / "artifacts" / f"{suite}.saved.xlsx", len(cases))
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
    path = WINDOWS / f"{suite}.tsv"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    shutil.copyfile(job / "job.json", WINDOWS / f"{suite}.job.json")
    return path


def _read_dynamic(suite: str, job: Path, result: Path, header: str) -> Path:
    """A ``dyn_`` suite's capture: its inputs and every cell of the saved copy."""
    from capture.dynamic_probe import records, rich_parts
    from capture.dynamic_suites import BLOCK, layout

    saved = result / "artifacts" / f"{suite}.saved.xlsx"
    lines = [
        f"# suite: {suite}",
        "# generator: capture/dynamic_probe.py, capture/job.py",
        "# method: the job beside this file (" + suite + ".job.json) opens the workbook the",
        "#   generator writes (dynamic-array formulas, one per block of "
        f"{BLOCK} rows in column B) and saves",
        "#   a copy; every cell of the copy is recorded as capture/dynamic_probe.py records it",
        "# header (the job's result, copied whole):",
        *("#   " + line for line in header.splitlines()),
        "# input rows: input, kind, cell, text (formula in file form, or a constant's repr)",
        "# cell rows: cell, address, type, value (repr), formula (repr), f_t, f_ref, cm, vm",
        *records(layout(suite[len("dyn_") :]), saved),
    ]
    path = WINDOWS / f"{suite}.tsv"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    rich_parts(saved, WINDOWS, suite)
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
