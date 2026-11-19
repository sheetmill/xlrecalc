"""How Excel evaluates a conditional format's formula: capture ``cf_rules``.

    uv run python -m capture.cf_probe --feasibility     step 0: is a rule's fill and bold readable?
    uv run python -m capture.cf_probe --add-rule        step 1: can AppleScript add a formula rule?
    uv run python -m capture.cf_probe                   every case: tests/captures/cf_rules.tsv

Each case is a sheet of its own: a few constants, and one rule (two in group I) over a small range,
written with openpyxl, its format a red solid fill (``fgColor`` and ``bgColor`` both, since a
conditional format's solid fill draws ``bgColor``) and bold. Forms openpyxl cannot write are
patched into the sheet's XML: a ``sqref`` of several areas in a given order (openpyxl may reorder
them), a rule in the Excel 2010 extension (``x14:conditionalFormatting``, where Excel keeps a rule
reading another sheet), and structured references, written in file form. Every workbook holds the
control sheet ``ctl`` (``=TRUE`` over A1, ``=FALSE`` over A2): a case's reading counts only when
the control reads applied and not applied. Every case has a workbook of its own (group F's
share one), so a file Excel would offer to repair (error -50 with alerts off) loses one case.

Each workbook is opened with alerts off and read through ``display format`` one cell at a time
(``cf_read.applescript``: the cell's text, its interior color, its boldness), then closed without
saving. Group G's opening cases (``volatile_open``) use the method of capture ``volatile_names``:
a copy Excel saved, its cached values planted wrong, opened, read and saved again, to see whether a
volatile function in a rule makes Excel recalculate cells on opening.

Groups: A the cell a rule's relative references are measured from; B which results apply the
format; C a result that is a range or an array; D references to another sheet; E structured
references; F defined names; G volatile functions; H ``cellIs`` and ``containsText`` operands; I a
merged range and two rules with ``stopIfTrue``.

Workbooks live in Excel's container and are removed after; the record is written to
``capture/build/cf_probe.out``.
"""

from __future__ import annotations

import re
import sys
import zipfile
from dataclasses import dataclass, field, replace
from pathlib import Path

from capture.edition import edition_line
from capture.excel_watch import osascript_or_exit, preflight
from capture.open_calc_probe import saved_values
from capture.writeback_probe import CONTAINER, rewrite

HERE = Path(__file__).resolve().parent
OUT = HERE / "build" / "cf_probe.out"
CAPTURE = HERE.parent / "tests" / "captures" / "cf_rules.tsv"
RS, US = "\x1e", "\x1f"
OTHER = "other"
"""A sheet in every workbook, A1 = 5: what group D's rules read."""


@dataclass(frozen=True)
class Rule:
    """One conditional-format rule: its range and formulas as a file stores them (no ``=``)."""

    sqref: str
    formulas: tuple[str, ...]
    kind: str = "expression"  # or cellIs, containsText
    operator: str | None = None
    text: str | None = None  # containsText's text
    stop: bool = False
    fill: bool = True
    bold: bool = True


@dataclass(frozen=True)
class Case:
    """A sheet: constants, rules, the cells read, and how the workbook is written."""

    name: str
    group: str
    rules: tuple[Rule, ...]
    cells: dict[str, object] = field(default_factory=dict)
    read: tuple[str, ...] = ()
    workbook: str = ""  # its own workbook (named after the case) unless one is named
    form: str = "plain"  # or "x14": the rule in the 2010 extension
    names: dict[str, str] = field(default_factory=dict)  # workbook names: name -> text
    local_names: dict[str, str] = field(default_factory=dict)  # names local to the case's sheet
    table: bool = False  # Table1 over A1:B4 (columns Col, Val)
    merge: str | None = None


def _expand(ref: str) -> list[str]:
    """The cells of ``A1:B2`` (or ``A1``), row by row; whole columns as their first three cells."""
    m = re.fullmatch(r"([A-Z]+)(\d*):?([A-Z]*)(\d*)", ref)
    assert m is not None, ref
    c1, r1, c2, r2 = m.groups()
    if not r1:  # a whole column, C:C
        r1, r2 = "1", "3"
    c2, r2 = c2 or c1, r2 or r1
    cols = range(_col(c1), _col(c2) + 1)
    return [f"{_letters(c)}{r}" for r in range(int(r1), int(r2) + 1) for c in cols]


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _letters(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def cells_read(case: Case) -> tuple[str, ...]:
    """The cells a case reads: those named, else every cell of its rules' ranges."""
    if case.read:
        return case.read
    seen: list[str] = []
    for rule in case.rules:
        for area in rule.sqref.split():
            seen += [c for c in _expand(area) if c not in seen]
    return tuple(seen)


def _cases() -> list[Case]:
    cases: list[Case] = []
    one_to_five = {f"A{r}": float(r) for r in range(1, 6)}
    # A: the anchor. ROW(A1)=1 holds in the anchor's row only, COLUMN(A1)=1 in its column.
    for label, sqref in (
        ("one_area", "B2:C4"),
        ("two_areas", "D2:D3 B4:C5"),
        ("two_areas_reversed", "B4:C5 D2:D3"),
        ("two_cells", "D5 B2"),
        ("whole_column", "C:C"),
    ):
        cases.append(Case(f"anchor_row_{label}", "A", (Rule(sqref, ("ROW(A1)=1",)),)))
        cases.append(Case(f"anchor_col_{label}", "A", (Rule(sqref, ("COLUMN(A1)=1",)),)))
    cases.append(Case("anchor_relative_data", "A", (Rule("B2:B5", ("A2>2",)),), one_to_five))
    # B: which results apply the format, over B1 (Z1:Z6 hold the values read)
    data = {"Z2": "TRUE", "Z3": True, "Z4": 0.0, "Z5": "x", "Z6": 2.0}
    for label, formula in (
        ("true", "TRUE"),
        ("false", "FALSE"),
        ("one", "1"),
        ("zero", "0"),
        ("minus_one", "-1"),
        ("half", "0.5"),
        ("tiny", "1E-300"),
        ("text_true", '"TRUE"'),
        ("text_x", '"x"'),
        ("text_empty", '""'),
        ("text_one", '"1"'),
        ("na", "NA()"),
        ("div0", "1/0"),
        ("value", '"a"+1'),
        ("blank_cell", "Z1"),
        ("cell_text_true", "Z2"),
        ("cell_true", "Z3"),
        ("cell_zero", "Z4"),
        ("cell_text", "Z5"),
        ("cell_two", "Z6"),
        ("empty_string_formula", 'IF(TRUE,"")'),
    ):
        cases.append(Case(f"truth_{label}", "B", (Rule("B1", (formula,)),), data))
    # C: a range or an array as the result, over C1:C5 (D1:D3 hold 0, 1, 1)
    d = {"D1": 0.0, "D2": 1.0, "D3": 1.0}
    for label, formula in (
        ("range_absolute", "$D$1:$D$3"),
        ("range_relative", "D1:D3"),
        ("array_false_first", "{FALSE,TRUE}"),
        ("array_true_first", "{TRUE,FALSE}"),
        ("range_compared", "$D$1:$D$3>0"),
        ("sequence", "_xlfn.SEQUENCE(2)>1"),
        ("sum_of_range", "SUM($D$1:$D$3)>1"),
        ("index_column", "INDEX($D$1:$D$3,0)>0"),
        ("single_at", "_xlfn.SINGLE($D$1:$D$3)>0"),
        ("row_of_range", "ROW($D$1:$D$3)=ROW()"),
    ):
        cases.append(Case(f"shape_{label}", "C", (Rule("C1:C5", (formula,)),), d))
    # D: another sheet, each in a workbook of its own (Excel may refuse the plain form)
    cases.append(
        Case("other_sheet_plain", "D", (Rule("A1", (f"{OTHER}!$A$1>2",)),), workbook="d_plain")
    )
    cases.append(
        Case(
            "other_sheet_x14",
            "D",
            (Rule("A1", (f"{OTHER}!$A$1>2",)),),
            workbook="d_x14",
            form="x14",
        )
    )
    cases.append(
        Case(
            "other_sheet_by_name",
            "D",
            (Rule("A1", ("OTHERCELL>2",)),),
            names={"OTHERCELL": f"{OTHER}!$A$1"},
            workbook="d_name",
        )
    )
    # E: structured references, each in a workbook of its own; Table1 is A1:B4 (Col 1..3)
    table = {"A1": "Col", "B1": "Val", "A2": 1.0, "A3": 2.0, "A4": 3.0}
    for label, formula in (
        ("this_row", "Table1[[#This Row],[Col]]>1"),
        ("column_sum", "SUM(Table1[Col])>5"),
    ):
        cases.append(
            Case(
                f"table_{label}",
                "E",
                (Rule("B2:B4", (formula,)),),
                table,
                table=True,
                workbook=f"e_{label}",
            )
        )
    # F: names
    names = {"LIMIT": "2", "SPOT": "ctl!$C$1"}  # ctl!C1 holds 5
    cases.append(
        Case(
            "name_constant",
            "F",
            (Rule("B1:B5", ("A1>LIMIT",)),),
            one_to_five,
            names=names,
            workbook="names",
        )
    )
    cases.append(Case("name_cell", "F", (Rule("B1", ("SPOT>2",)),), names=names, workbook="names"))
    cases.append(
        Case(
            "name_local",
            "F",
            (Rule("B1:B5", ("A1>LOCALLIMIT",)),),
            one_to_five,
            local_names={"LOCALLIMIT": "3"},
            workbook="names",
        )
    )
    cases.append(
        Case(
            "name_lambda",
            "F",
            (Rule("B1:B5", ("BIG(A1)",)),),
            one_to_five,
            names={"BIG": "_xlfn.LAMBDA(_xlpm.x,_xlpm.x>4)"},
            workbook="names",
        )
    )
    # G: volatile functions, evaluated
    for label, formula in (
        ("rand", "RAND()<2"),
        ("today", "TODAY()>DATE(2000,1,1)"),
        ("indirect", 'INDIRECT("A"&ROW())>2'),
        ("offset", "OFFSET(A1,0,0)>2"),
    ):
        cases.append(Case(f"volatile_{label}", "G", (Rule("B1:B5", (formula,)),), one_to_five))
    # H: cellIs and containsText; Z1 holds 2, B1:B5 hold 3, 3, 1, 1, 9
    h = {**one_to_five, "Z1": 2.0, "B1": 3.0, "B2": 3.0, "B3": 1.0, "B4": 1.0, "B5": 9.0}
    for label, operator, formulas in (
        ("greater_absolute", "greaterThan", ("$Z$1",)),
        ("greater_relative", "greaterThan", ("B1",)),
        ("between", "between", ("2", "4")),
        ("between_reversed", "between", ("4", "2")),
        ("between_relative", "between", ("B1", "$Z$1")),
        ("equal_formula", "equal", ("B1*0+3",)),
        ("not_equal_text", "notEqual", ('"x"',)),
    ):
        rule = Rule("A1:A5", formulas, kind="cellIs", operator=operator)
        cases.append(Case(f"cellis_{label}", "H", (rule,), h))
    words = {"A1": "abc", "A2": "xyz", "A3": "ABC", "A4": 12.0, "A5": "b"}
    cases.append(
        Case(
            "contains_text",
            "H",
            (
                Rule(
                    "A1:A5",
                    ('NOT(ISERROR(SEARCH("b",A1)))',),
                    kind="containsText",
                    operator="containsText",
                    text="b",
                ),
            ),
            words,
        )
    )
    # I: a merged range; two rules on one range, the first bold only and stopping
    cases.append(Case("merged", "I", (Rule("A1:B1", ("TRUE",)),), merge="A1:B1", read=("A1", "B1")))
    cases.append(
        Case(
            "stop_if_true",
            "I",
            (Rule("A1", ("TRUE",), stop=True, fill=False), Rule("A1", ("TRUE",), bold=False)),
        )
    )
    cases.append(Case("rule_error", "I", (Rule("A1:A2", ("1/0",)),)))
    return cases


CASES = [c if c.workbook else replace(c, workbook=c.name) for c in _cases()]
SHEET_OF = {}
for _book in sorted({c.workbook for c in CASES}):
    for _i, _case in enumerate(c for c in CASES if c.workbook == _book):
        SHEET_OF[_case.name] = f"c{_i + 1:02d}"
"""Each case's sheet name in its workbook."""


def _dxf(rule: Rule):  # type: ignore[no-untyped-def]
    from openpyxl.styles import Font, PatternFill
    from openpyxl.styles.differential import DifferentialStyle

    fill = (
        PatternFill(fill_type="solid", start_color="FFFF0000", end_color="FFFF0000")
        if rule.fill
        else None
    )
    return DifferentialStyle(font=Font(bold=True) if rule.bold else None, fill=fill)


def build(book: str, path: Path) -> list[tuple[str, list[Rule]]]:
    """Write one workbook; the sheets whose rules' ``sqref`` or form must be patched."""
    import openpyxl
    from openpyxl.formatting.rule import Rule as XlRule
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.table import Table

    wb = openpyxl.Workbook()
    ctl = wb.active
    ctl.title = "ctl"
    ctl["A1"], ctl["A2"], ctl["C1"] = "on", "off", 5.0
    for sqref, formula in (("A1", "TRUE"), ("A2", "FALSE")):
        ctl.conditional_formatting.add(
            sqref, XlRule(type="expression", formula=[formula], dxf=_dxf(Rule(sqref, ())))
        )
    wb.create_sheet(OTHER)["A1"] = 5.0
    patches: list[tuple[str, list[Rule]]] = []
    for case in (c for c in CASES if c.workbook == book):
        ws = wb.create_sheet(SHEET_OF[case.name])
        for ref, value in case.cells.items():
            ws[ref] = value
        if case.merge:
            ws.merge_cells(case.merge)
        if case.table:
            ws.add_table(Table(displayName="Table1", ref="A1:B4"))
        for name, text in case.names.items():
            wb.defined_names[name] = DefinedName(name, attr_text=text)
        for name, text in case.local_names.items():
            ws.defined_names[name] = DefinedName(name, attr_text=text)
        for rule in case.rules:
            kwargs: dict[str, object] = {
                "type": rule.kind,
                "formula": list(rule.formulas),
                "dxf": _dxf(rule),
            }
            if rule.operator:
                kwargs["operator"] = rule.operator
            if rule.text:
                kwargs["text"] = rule.text
            if rule.stop:
                kwargs["stopIfTrue"] = True
            # a placeholder range per sqref, replaced in the XML (openpyxl refuses C:C, may reorder)
            placeholder = f"ZZ{100 + sorted({r.sqref for r in case.rules}).index(rule.sqref)}"
            ws.conditional_formatting.add(placeholder, XlRule(**kwargs))
        patches.append((SHEET_OF[case.name], list(case.rules)))
    wb.save(path)
    forms = {SHEET_OF[c.name]: c.form for c in CASES if c.workbook == book}
    _patch(path, patches, forms)
    return patches


_X14 = (
    '<extLst><ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    "<x14:conditionalFormattings>"
    '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    '<x14:cfRule type="expression" priority="1" id="{{6C1B4E52-3A5D-4F2B-9C1E-2D7A8B9C0D1E}}">'
    '<xm:f>{formula}</xm:f><x14:dxf><font><b/></font><fill><patternFill patternType="solid">'
    '<fgColor rgb="FFFF0000"/><bgColor rgb="FFFF0000"/></patternFill></fill></x14:dxf></x14:cfRule>'
    "<xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>"
)


def _patch(path: Path, patches: list[tuple[str, list[Rule]]], forms: dict[str, str]) -> None:
    """Each case sheet's ``sqref`` as given; an x14 case's rule moved into the extension."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
        sheets = _sheet_parts(z)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in items:
            sheet = sheets.get(info.filename)
            rules = dict(patches).get(sheet) if sheet else None
            if rules:
                text = data.decode()
                real = {f"ZZ{100 + i}": q for i, q in enumerate(sorted({r.sqref for r in rules}))}
                text = re.sub(
                    r'(<conditionalFormatting sqref=")(ZZ1\d\d)(")',
                    lambda m, real=real: m.group(1) + real[m.group(2)] + m.group(3),
                    text,
                )
                assert "ZZ1" not in text, sheet
                if forms.get(sheet) == "x14":
                    (rule,) = rules
                    text = re.sub(
                        r"<conditionalFormatting .*?</conditionalFormatting>", "", text, flags=re.S
                    )
                    formula = (
                        rule.formulas[0]
                        .replace("&", "&amp;")
                        .replace(">", "&gt;")
                        .replace("<", "&lt;")
                    )
                    text = text.replace(
                        "</worksheet>",
                        _X14.format(formula=formula, sqref=rule.sqref) + "</worksheet>",
                    )
                data = text.encode()
            out.writestr(info, data)


def _sheet_parts(z: zipfile.ZipFile) -> dict[str, str]:
    """Worksheet part -> sheet name."""
    workbook = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="/?(?:xl/)?([^"]+)"', rels))
    targets |= {i: t for t, i in re.findall(r'Target="/?(?:xl/)?([^"]+)"[^>]*Id="([^"]+)"', rels)}
    return {
        "xl/" + targets[rid]: name
        for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', workbook)
    }


def read(path: Path, cells: list[str]) -> list[str] | str:
    """``[text, color, bold]`` per cell, or Excel's error when it refused the file."""
    result = osascript_or_exit(
        str(HERE / "cf_read.applescript"), str(path), RS.join(cells), seconds=120
    )
    if result.returncode:
        raise RuntimeError(result.stderr.strip())
    fields = result.stdout.rstrip("\n").split(US)
    if fields[0] != "open":
        return " ".join(fields[1:])
    return fields[1:]


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    OUT.parent.mkdir(exist_ok=True)
    lines: list[str] = []

    def log(text: str) -> None:
        print(text, flush=True)
        lines.append(text)

    for p in CONTAINER.glob("xlrecalc_cf_*.xlsx"):
        p.unlink()
    rows: list[str] = []
    try:
        for book in sorted({c.workbook for c in CASES}):
            path = CONTAINER / f"xlrecalc_cf_{book}.xlsx"
            build(book, path)
            cases = [c for c in CASES if c.workbook == book]
            cells = ["ctl|A1", "ctl|A2"] + [
                f"{SHEET_OF[c.name]}|{a}" for c in cases for a in cells_read(c)
            ]
            fields = read(path, cells)
            if isinstance(fields, str):
                log(f"{book}: REFUSED {fields}")
                rows += [
                    f"{c.name}\t{c.group}\t{book}\t{_rules(c)}\t-\t-\t-\t-\trefused" for c in cases
                ]
                continue
            got = [fields[i : i + 3] for i in range(0, len(fields), 3)]
            on, off = got[0], got[1]
            controls = on[2] == "true" and off[2] == "false" and on[1] != off[1]
            log(f"{book}: control on {on} off {off}: {'ok' if controls else 'NOT READABLE'}")
            k = 2
            for case in cases:
                for cell in cells_read(case):
                    text, color, bold = got[k]
                    k += 1
                    fill = color == on[1]
                    applied = (
                        (
                            "applied"
                            if fill and bold == "true"
                            else "not applied"
                            if not fill and bold == "false"
                            else "mixed"
                        )
                        if controls
                        else "unreadable"
                    )
                    rows.append(
                        "\t".join(
                            (
                                case.name,
                                case.group,
                                book,
                                _rules(case),
                                cell,
                                text,
                                color,
                                bold,
                                applied,
                            )
                        )
                    )
        rows += volatile_open(log)
    finally:
        OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
        for p in CONTAINER.glob("xlrecalc_cf_*.xlsx"):
            p.unlink()
    header = [
        "# suite: cf_rules (how Excel evaluates a conditional format's formula)",
        f"# excel: {_version()} for Mac, locale en_US",
        edition_line(),
        "# generator: capture/cf_probe.py (cases in CASES; a rule's format is a red fill and bold;",
        "#   read through display format, cell by cell, against the control sheet ctl)",
        "# columns: case, group, workbook, rules (sqref|kind|operator|formulas, joined by ' ;; '),",
        "#   cell, displayed text, interior color, bold, applied / not applied / mixed / refused",
    ]
    CAPTURE.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")


def _rules(case: Case) -> str:
    return " ;; ".join(
        f"{r.sqref}|{r.kind}|{r.operator or '-'}|{' ~ '.join(r.formulas)}"
        + ("|stop" if r.stop else "")
        for r in case.rules
    )


def _version() -> str:
    result = osascript_or_exit("-e", 'tell application "Microsoft Excel" to version', seconds=30)
    return result.stdout.strip()


VOLATILE_OPEN = {
    "open_rule_rand": "RAND()<2",
    "open_rule_offset": "OFFSET(A1,0,0)>0",
    "open_rule_plain": "A1>0",
}
"""Group G on opening: a rule over A2 with this formula; A2 = A1*3 and A3 = A1*4 planted wrong."""


def volatile_open(log) -> list[str]:  # type: ignore[no-untyped-def]
    """Whether a volatile function in a rule makes Excel recalculate cells on opening its copy."""
    import openpyxl
    from openpyxl.formatting.rule import Rule as XlRule

    from capture.volatile_name_probe import cell_xml

    rows = []
    for name, formula in VOLATILE_OPEN.items():
        src = CONTAINER / f"xlrecalc_cf_{name}_src.xlsx"
        base = CONTAINER / f"xlrecalc_cf_{name}_base.xlsx"
        path = CONTAINER / f"xlrecalc_cf_{name}.xlsx"
        copy = CONTAINER / f"xlrecalc_cf_{name}_saved.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "probe"
        ws["A1"], ws["A2"], ws["A3"] = 2, "=A1*3", "=A1*4"
        ws.conditional_formatting.add(
            "A2", XlRule(type="expression", formula=[formula], dxf=_dxf(Rule("A2", ())))
        )
        wb.save(src)
        log(
            f"{name}: base "
            + osascript_or_exit(
                str(HERE / "save_copy.applescript"), str(src), str(base), seconds=120
            ).stdout.strip()
        )
        _plant(base, path, {"A2": "999", "A3": "888"})
        log(f"  planted A2 {cell_xml(path, 'A2')} A3 {cell_xml(path, 'A3')}")
        shown = (
            osascript_or_exit(str(HERE / "open_read.applescript"), str(path), "3", seconds=120)
            .stdout.strip()
            .split(US)
        )
        if shown[0] != "open":
            log(f"  REFUSED {shown[1:]}")
            continue
        texts = [shown[2 + 3 * i] for i in range(3)]
        osascript_or_exit(str(HERE / "save_copy.applescript"), str(path), str(copy), seconds=120)
        saved = saved_values(copy)
        log(f"  shown {texts}; saved copy {saved}")
        recalculated = texts[1] != "999" or texts[2] != "888"
        if recalculated:
            log(f"  NOTE: opening recalculated cells with only the rule {formula!r} volatile")
        rows.append(
            "\t".join(
                (
                    name,
                    "G",
                    "open",
                    f"A2|expression|-|{formula}",
                    "A2:A3",
                    " ".join(texts[1:]),
                    saved,
                    "-",
                    "recalculated" if recalculated else "kept",
                )
            )
        )
    return rows


def _plant(src: Path, dst: Path, values: dict[str, str]) -> None:
    """A copy with each named cell's ``<v>`` text replaced, the rest as Excel wrote it."""

    def edit(xml: str) -> str:
        for ref, value in values.items():
            cell = re.compile(rf'(<c r="{ref}"[^>]*>.*?<v>)(.*?)(</v>)', re.S)
            xml, n = cell.subn(lambda m, v=value: m.group(1) + v + m.group(3), xml, count=1)
            if n != 1:
                raise RuntimeError(f"no cached value in {ref}: cannot plant")
        return xml

    rewrite(src, dst, sheet=edit)


def feasibility() -> None:
    """Step 0: the control sheet alone, read: does display format report a rule's fill and bold?"""
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    path = CONTAINER / "xlrecalc_cf_feasibility.xlsx"
    try:
        build("feasibility", path)  # no case uses it: the control sheet and the other sheet only
        print(read(path, ["ctl|A1", "ctl|A2"]))
    finally:
        path.unlink(missing_ok=True)


def add_rule() -> None:
    """Step 1: can AppleScript add a formula rule, and how does Excel store it?"""
    import openpyxl

    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    path = CONTAINER / "xlrecalc_cf_add_rule.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r in range(1, 6):
        ws[f"A{r}"] = r
    wb.save(path)
    try:
        result = osascript_or_exit(
            str(HERE / "cf_add_rule.applescript"), str(path), "probe", "D2:D3", "=A2>2", seconds=120
        )
        print(result.stdout.strip() or result.stderr.strip())
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode()
        print(re.findall(r"<conditionalFormatting.*?</conditionalFormatting>", sheet, re.S))
    finally:
        path.unlink(missing_ok=True)


if __name__ == "__main__":
    if sys.argv[1:2] == ["--feasibility"]:
        feasibility()
    elif sys.argv[1:2] == ["--add-rule"]:
        add_rule()
    else:
        main()
