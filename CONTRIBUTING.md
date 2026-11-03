# Contributing

## Issues and bug reports

Bug reports are welcome. The most useful report includes a small workbook (or just the
formula and its inputs) that shows the problem, the result Excel computes, the result
xlrecalc computes, and your Excel version and platform.

## Pull requests

The project is just getting started, so pull requests aren't being accepted yet while
the test-capture workflow, release process and review practices settle. This will be
revisited after the first releases. In the meantime, if you have a change in mind,
please open an issue describing it.

When pull requests open, these rules will apply:

- AI-assisted contributions must be disclosed in the pull request description.
- The contributor must understand every line of the change and be able to explain it when asked.
- Undisclosed or unexplained AI-generated pull requests may be closed without review.

## Setting up

Python 3.11 or later. Either route installs the package in editable mode with the development
and documentation tools:

```bash
uv sync --group docs
```

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e . --group dev --group docs      # needs pip 25.1 or later
```

`uv.lock` pins the versions CI tests with, and `uv sync` installs exactly those versions.
pip and other tools ignore it and resolve the latest versions the dependency groups in
`pyproject.toml` allow, which is normally fine but is not what CI runs.

## Checks to run before proposing a change

All of these must pass. CI runs the same set on Python 3.11 to 3.14, plus a run without
openpyxl, a development dependency. Prefix each with `uv run` on the uv route.

```bash
pytest                                   # unit tests and the doc examples
ruff check .
ruff format --check .
mypy
sphinx-build -n -W docs docs/_build/html # the strict documentation build
```

The examples in `README.md` and `docs/*.md` run as doctests, so a change of behavior shows up
there too.

On a Mac with Excel 16, `pytest --tier excel` also runs the tests that drive Excel: it opens
the files the writer produces, which must open without a repair and hold Excel's values
(`tests/test_excel_live.py`, `CAPTURING.md`). The default run never starts Excel.

## Benchmarks

`bench/` is a separate uv project, so that the packages it compares against and their
dependencies never enter this package's environment. `bench/results/SUMMARY.md` lists the
commands.

## Expected results

The result of a formula is what Excel computes. Expected results are captured from Excel for
Mac by the scripts in `capture/`: Excel recalculates a workbook built from formulas and inputs,
saves a copy, and the values are read from the saved file (`CAPTURING.md`). Excel behaviors
reproduced on purpose and switchable are listed in `QUIRKS.md`. They are never written from memory
and never taken from another program's output or another file's cached values. Captured files
are never edited by hand. A capture that changes any expected result is reviewed as a change of
evidence.

## Commit messages

A commit made with the help of an AI tool ends with the trailer

```
Assisted-by: <tool>:<model ID>
```

as its last line, after a blank line, naming the tool and the exact model, for example
`Assisted-by: Claude Code:claude-opus-5-5`. One trailer per tool and model used.

## How this project's own commits are made

Most of the code is written with Claude Code, reviewed before each commit and tested against
results captured from real Excel. Every such commit carries the git trailer
`Assisted-by: Claude Code:<model ID>`, so the history records which model took part.
