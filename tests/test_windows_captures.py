"""The Windows captures are what their jobs record (``capture/job.py``, ``CAPTURING.md``).

Each capture made by a job keeps the manifest beside it and the job's result header in its own
header. Writing the job again from the header's ``job_id`` gives the same manifest and the same
input files, checksum for checksum, and the capture's rows are the generator's cases in order.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

WINDOWS = Path(__file__).parent / "captures" / "windows"
CAPTURES = sorted(WINDOWS.glob("*.tsv"))


def _header(path: Path) -> dict[str, object]:
    lines = path.read_text(encoding="utf-8").splitlines()
    start = lines.index("# header (the job's result, copied whole):") + 1
    body = []
    for line in lines[start:]:
        if not line.startswith("#   "):
            break
        body.append(line[4:])
    return json.loads("\n".join(body))


def test_there_is_a_windows_capture() -> None:
    assert [p.name for p in CAPTURES] == [
        "bessel_family.tsv",
        "bonds.tsv",
        "cancellation.tsv",
        "clean_controls.tsv",
        "codepage.tsv",
        "collate_chars.tsv",
        "collate_functions.tsv",
        "collate_words.tsv",
        "combin_limit.tsv",
        "compare.tsv",
        "complex_family.tsv",
        "continuous_family.tsv",
        "convert_family.tsv",
        "convert_odd.tsv",
        "convert_pairs.tsv",
        "criteria.tsv",
        "cum_large.tsv",
        "database_family.tsv",
        "dates.tsv",
        "dates_1904.tsv",
        "dates_family.tsv",
        "distributions_family.tsv",
        "empty_args.tsv",
        "encodeurl.tsv",
        "engineering_family.tsv",
        "finance.tsv",
        "finance_family.tsv",
        "function_names.tsv",
        "functions.tsv",
        "info_family.tsv",
        "iteration.tsv",
        "iteration_steps.tsv",
        "iterative.tsv",
        "literals.tsv",
        "mac_gaps.tsv",
        "math_integers.tsv",
        "math_trig.tsv",
        "matrix_family.tsv",
        "misc.tsv",
        "modern.tsv",
        "n_lookup.tsv",
        "number_text.tsv",
        "rate_method.tsv",
        "rounding.tsv",
        "statistics.tsv",
        "statistics_family.tsv",
        "sums.tsv",
        "text_family.tsv",
        "text_number.tsv",
        "textfmt.tsv",
        "whole_columns.tsv",
        "workdays.tsv",
        "xirr_method.tsv",
        "xirr_nan.tsv",
        "xirr_sign_rules.tsv",
    ]


@pytest.mark.parametrize("path", CAPTURES, ids=lambda p: p.stem)
def test_the_job_beside_the_capture_is_the_one_that_ran(path: Path) -> None:
    header = _header(path)
    manifest = path.with_suffix(".job.json")
    assert hashlib.sha256(manifest.read_bytes()).hexdigest() == header["manifest_sha256"]
    assert json.loads(manifest.read_text(encoding="utf-8"))["job_id"] == header["job_id"]
    assert header["format"] == 1


@pytest.mark.parametrize("path", CAPTURES, ids=lambda p: p.stem)
def test_writing_the_job_again_gives_the_same_files(path: Path, tmp_path: Path) -> None:
    pytest.importorskip("openpyxl")  # the generator writes the workbooks with it
    from capture.job import write

    header = _header(path)
    committed = json.loads(path.with_suffix(".job.json").read_text(encoding="utf-8"))
    folder = write(path.stem, str(header["job_id"]), tmp_path, committed.get("requires"))
    assert (folder / "job.json").read_bytes() == path.with_suffix(".job.json").read_bytes()
    inputs = header["inputs"]
    assert isinstance(inputs, dict)
    for name, entry in inputs.items():
        data = (folder / "files" / name).read_bytes()
        assert (hashlib.sha256(data).hexdigest(), len(data)) == (entry["sha256"], entry["bytes"])


@pytest.mark.parametrize("path", CAPTURES, ids=lambda p: p.stem)
def test_the_rows_are_the_generator_s_cases(path: Path) -> None:
    pytest.importorskip("openpyxl")  # capture.suites imports the pipeline, which needs it
    from capture.suites import SUITES

    rows = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [r for r in rows if r[0] and not r[0].startswith("#")]
    cases = SUITES[path.stem].cases()
    assert [(r[3], r[4]) for r in rows] == [(c.formula, repr(c.inputs)) for c in cases]
    assert all(r[5] != "none" for r in rows)


COMPARED = [
    "clean_controls",
    "codepage",
    "collate_chars",
    "collate_words",
    "collate_functions",
    "modern",
    "encodeurl",
]
"""The Windows captures the package is compared with, under ``platform="windows"``: those whose
platform-dependent behavior it implements for Excel for Windows."""


@pytest.mark.parametrize("suite", COMPARED)
def test_the_package_on_windows_gives_every_value_the_capture_holds(suite: str) -> None:
    """Each row of the Windows capture, computed on an engine for Excel for Windows: the same
    value and type, bit for bit, as Excel for Windows saved."""
    pytest.importorskip("openpyxl")  # tests.test_captures reads the suites
    from dataclasses import replace

    from tests.captures import _expected, load
    from tests.test_captures import KNOWN_DEVIATIONS, run

    lines = (WINDOWS / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
    fields = [line.split("\t") for line in lines if line and not line.startswith("#")]
    mac = load(suite)
    # the rows are the generator's cases (inputs too): test_the_rows_are_the_generator_s_cases
    assert [f[3] for f in fields] == [r.formula for r in mac.rows]
    windows = replace(
        mac,
        rows=[
            replace(r, expected=_expected(f[5], f[6]))
            for r, f in zip(mac.rows, fields, strict=True)
        ],
    )
    engine, s = run(windows, platform="windows")
    wrong = []
    for k, (row, on_mac) in enumerate(zip(windows.rows, mac.rows, strict=True)):
        got = engine.get(s, k, 0)
        same = type(got) is type(row.expected) and got == row.expected
        # A row listed as a deviation on the Mac, where Excel for Windows gives the Mac's value,
        # is the same deviation; it must still differ, as test_captures has it.
        if (suite, row.case) in KNOWN_DEVIATIONS and repr(row.expected) == repr(on_mac.expected):
            same = not same
        if not same:
            wrong.append((k + 1, row.formula, row.inputs, row.expected, got))
    assert not wrong, wrong[:10]
