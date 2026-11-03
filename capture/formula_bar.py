"""Excel's file form of formulas entered as the formula bar takes them (ARCHITECTURE.md §5.2).

    uv run python -m capture.formula_bar [--ranges]

Non-interactive. The workbook of ``capture/formula_forms.py`` (its sheets, names and table)
gets each formula typed into a cell through AppleScript (``formula_bar.applescript``: the
cell's ``formula``, as entered in the formula bar), and Excel saves a copy; the ``<f>`` text of
the copy is Excel's file form of what was typed, or ``REJECTED`` when Excel does not take the
formula (it leaves the cell as it was, without an error). The formulas are the display form of
every corpus formula (``display_form``), a few chosen cases (``EXTRA``), and calls of every
function in ``FUNCTIONS`` with 0 to 6 arguments, which show the file-form prefix each one
takes. AppleScript's ``formula`` is the legacy formula: a ``LAMBDA`` call is not recognized
there (stored as typed), while a file holding ``_xlfn.LAMBDA(_xlpm.x,...)`` opens
(``formula_forms``).
Writes ``tests/captures/formula_bar.tsv``; with ``--ranges``, calls with range arguments of the
functions that take no number-only call (``RANGE_ARGUMENTS``), to
``tests/captures/formula_bar_ranges.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

from capture.corpus_filter import SIDE_EFFECTS
from capture.formula_forms import CONTAINER, build, corpus, excel_version, saved_formulas
from capture.functions_list import FUNCTIONS
from xlrecalc._formula._fileform import display_form

HERE = Path(__file__).resolve().parent
CAPTURES = HERE.parent / "tests" / "captures"
RS, US = "\x1e", "\x1f"
# Every cell starts with this formula: Excel leaves it in place, without an error, when it does
# not take the formula typed over it.
PLACEHOLDER = "0.5+0.25"
EXTRA = [
    "=#SPILL!",
    "=#CALC!",
    '="a\tb"',
    "=SUM( )",
    "=-\t1",
    "=_xlfn.IFNA(1,2)",
    "=@A1:A3",
    "=A1#",
    "=LAMBDA(x,x+1)(2)",
]


def callable_functions() -> list[str]:
    """FUNCTIONS without those that reach outside the workbook (``corpus_filter``)."""
    return [
        name
        for name in FUNCTIONS
        if name not in SIDE_EFFECTS and not name.startswith("CUBE") and name != "HYPERLINK"
    ]


def cases() -> list[tuple[str, str]]:
    out = [("corpus", "=" + display_form(f)) for f in corpus()]
    out += [("extra", f) for f in EXTRA]
    for name in callable_functions():  # every arity up to 6, so that one of them is taken
        out += [("function", f"={name}({','.join(['1'] * n)})") for n in range(7)]
    return out


# Functions the formula bar took with no number-only call: their arguments must be ranges.
RANGE_ARGUMENTS = """AREAS AVERAGEIF AVERAGEIFS COUNTBLANK COUNTIF COUNTIFS DAVERAGE DCOUNT DCOUNTA
DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP FORMULATEXT ISFORMULA MAXIFS MINIFS
ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET PHONETIC RANK RANK.AVG RANK.EQ SUBTOTAL SUMIF
SUMIFS""".split()  # noqa: SIM905


def range_cases() -> list[tuple[str, str]]:
    out = []
    for name in RANGE_ARGUMENTS:
        for n in range(1, 7):
            out.append(("function", f"={name}({','.join(['A1:A2'] + ['1'] * (n - 1))})"))
            out.append(("function", f"={name}({','.join(['1'] + ['A1:A2'] * (n - 1))})"))
    return out


def main() -> None:
    ranges = "--ranges" in sys.argv[1:]
    rows = range_cases() if ranges else cases()
    output = CAPTURES / ("formula_bar_ranges.tsv" if ranges else "formula_bar.tsv")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_formula_bar.xlsx"
    copy = CONTAINER / "xlrecalc_formula_bar_saved.xlsx"
    listing = CONTAINER / "xlrecalc_formula_bar.txt"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src, [PLACEHOLDER] * len(rows))
    listing.write_text(RS.join(text for _, text in rows), encoding="utf-8")
    version = excel_version()
    result = subprocess.run(
        ["osascript", str(HERE / "formula_bar.applescript"), str(src), str(listing), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    rejected = {}
    for line in result.stdout.strip().splitlines():
        number, code, message = line.split(US)
        rejected[int(number)] = f"REJECTED {code}: {message}"
    saved = saved_formulas(copy)
    header = [
        f"# suite: {output.stem}",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/formula_bar.py (corpus in display form, and one call per function)",
        "# method: each formula is set as a cell's formula through AppleScript, as typed in the",
        "#   formula bar; Excel saves a copy; file is the <f> text of the copy, REJECTED when",
        "#   Excel does not take the formula",
        "# columns: tag, typed, file (Python repr)",
    ]
    body = []
    for k, (tag, typed) in enumerate(rows, 1):
        excel = rejected.get(k) or saved[k - 1]
        if excel == PLACEHOLDER:
            excel = "REJECTED"
        body.append("\t".join([tag, repr(typed), repr(excel)]))
    output.write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    print(f"{len(rows)} formulas -> {output.name}")
    for p in (src, copy, listing):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
