"""The Windows captures are what their jobs record (``capture/job.py``, ``CAPTURING.md``).

Each capture made by a job keeps the manifest beside it and the job's result header in its own
header. Writing the job again from the header's ``job_id`` gives the same manifest and the same
input files, checksum for checksum, and the capture's rows are the generator's cases in order.
"""

from __future__ import annotations

import ast
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
        "bare_names.tsv",
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
        "dyn_cancel.tsv",
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
        "three_d.tsv",
        "ties_tiny.tsv",
        "unions.tsv",
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
    if path.stem.startswith("dyn_"):  # capture/dynamic_suites.py: the layout, then every cell
        from capture.dynamic_suites import layout

        inputs = [(r[1], r[2], ast.literal_eval(r[3])) for r in rows if r[0] == "input"]
        assert inputs == layout(path.stem.removeprefix("dyn_"))
        assert all(r[0] in ("input", "cell") for r in rows)
        return
    cases = SUITES[path.stem].cases()
    assert [(r[3], r[4]) for r in rows] == [(c.formula, repr(c.inputs)) for c in cases]
    assert all(r[5] != "none" for r in rows)


COMPARED_ELSEWHERE = {
    "mac_gaps": "a record of which names Excel knows, compared on neither platform",
    "whole_columns": "tests/test_whole_columns.py, on both platforms",
    "dyn_cancel": "tests/test_dynamic_arrays.py, on both platforms",
}
"""The Windows captures this file does not compare with the package, and why."""


@pytest.mark.parametrize(
    "suite", [p.stem for p in CAPTURES if p.stem not in COMPARED_ELSEWHERE], ids=str
)
def test_the_package_on_windows_gives_every_value_the_capture_holds(suite: str) -> None:
    """Each row of the Windows capture, computed on an engine for Excel for Windows, classified
    as ``tests/test_captures.py`` classifies the Mac's: Excel's value bit for bit or within the
    function's declared tolerance, a listed deviation (the Mac's, or ``WINDOWS_DEVIATIONS``) that
    must still differ where the Mac's must, an exempt row, or a function not computed."""
    pytest.importorskip("openpyxl")  # tests.test_captures reads the suites
    from tests.test_captures import SUITES, outcomes

    assert suite in SUITES
    _counts, _reasons, failures = outcomes(suite, "windows")
    assert not failures, failures[:10]


def test_capturing_md_says_which_windows_captures_are_compared() -> None:
    """CAPTURING.md §5 says "not compared" of no Windows capture this file compares."""
    text = (WINDOWS.parents[2] / "CAPTURING.md").read_text(encoding="utf-8")
    section = text[text.index("## 5. ") :]
    section = section[: section.index("\n## ", 1)] if "\n## " in section[1:] else section
    rows = [line for line in section.split("\n") if line.startswith("| `tests/captures/windows/")]
    compared = {p.stem for p in CAPTURES} - set(COMPARED_ELSEWHERE)
    assert rows
    for row in rows:
        name = row.split("`")[1].removeprefix("tests/captures/windows/").removesuffix(".tsv")
        if name == "<suite>" or name in compared:
            assert "not compared" not in row.lower(), row[:120]
