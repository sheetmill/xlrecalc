"""Generated source never holds workbook text (ARCHITECTURE.md §5.4).

Every formula of the corpus, hostile ones included, is compiled; Python's own tokenizer then
checks the generated source: no string tokens, only allowlisted names and temporaries,
plain integers, and no nesting deeper than two levels.
"""

import io
import keyword
import re
import tokenize as pytokenize

import pytest

from tests.formula_corpus import CORPUS, HOSTILE
from xlrecalc._formula import ALLOWED_NAMES, compile_formula, parse, template_key
from xlrecalc._values import ErrorValue

_TEMPORARY = re.compile(r"[tb]\d+")  # temporaries and branch functions
_INTEGER = re.compile(r"\d+")
_PUNCTUATION = {"(", ")", "[", "]", ",", ":", "=", "+", "-", "."}


class Recorder:
    """Stands in for both the evaluation context and the operators: records each call."""

    def __getattr__(self, name):
        return lambda *args: (name, *args)


def compiled(text, row=4, col=4, functions=None):
    return compile_formula(parse(text), row, col, functions or {}, Recorder())


def audit(source: str) -> None:
    depth = max_depth = 0
    for tok in pytokenize.generate_tokens(io.StringIO(source).readline):
        kind, text = tok.type, tok.string
        if kind == pytokenize.NAME:
            assert text in ALLOWED_NAMES or _TEMPORARY.fullmatch(text), text
            assert not keyword.iskeyword(text) or text in {"def", "return", "None"}, text
        elif kind == pytokenize.NUMBER:
            assert _INTEGER.fullmatch(text), text
        elif kind == pytokenize.OP:
            assert text in _PUNCTUATION, text
            if text in "([":
                depth += 1
                max_depth = max(max_depth, depth)
            elif text in ")]":
                depth -= 1
        else:
            assert kind in {
                pytokenize.NEWLINE,
                pytokenize.NL,
                pytokenize.INDENT,
                pytokenize.DEDENT,
                pytokenize.ENDMARKER,
            }, (pytokenize.tok_name[kind], text)
    assert max_depth <= 2, max_depth


@pytest.mark.parametrize("text", CORPUS, ids=repr)
def test_generated_source_holds_no_workbook_text(text):
    audit(compiled(text, functions={"SUM": lambda ctx, *a: a}).source)


@pytest.mark.parametrize("text", [t for t in HOSTILE if t.startswith('"')], ids=repr)
def test_string_literals_come_back_unchanged(text):
    expected = text[1:-1].replace('""', '"')
    result = compiled(text).fn(Recorder(), 4, 4)
    if "&" in text:  # '"lambda: 1"&"exec"'
        assert result == ("concat", "lambda: 1", "exec")
    else:
        assert result == expected


@pytest.mark.parametrize(
    ("text", "sheet"),
    [
        ("'__import__(''os'')'!A1", "__import__('os')"),
        ("'); import os; ('!A1", "); import os; ("),
        ("'It''s'!A1", "It's"),
        ("Sheet1:Sheet3!A1", "Sheet1:Sheet3"),
        ("[1]Sheet1!A1", "[1]Sheet1"),
    ],
)
def test_sheet_names_reach_the_context_exactly(text, sheet):
    assert compiled(text, 0, 0).fn(Recorder(), 0, 0) == ("cell", sheet, 0, 0)


def test_names_and_unknown_functions_go_through_constants():
    assert compiled("__import__", 0, 0).fn(Recorder(), 0, 0) == ("name", None, "__import__")
    result = compiled("eval(1)", 0, 0).fn(Recorder(), 0, 0)
    assert result == ("unsupported", "eval", 1.0)


def test_functions_are_called_by_index():
    seen = []

    def total(ctx, *args):
        seen.append(args)
        return 6.0

    c = compiled("SUM(1,2,3)", functions={"SUM": total})
    assert "SUM" not in c.source
    assert c.fn(Recorder(), 4, 4) == 6.0
    assert seen == [(1.0, 2.0, 3.0)]


def test_no_builtins_reach_generated_code():
    fn = compiled("1").fn
    assert fn.__globals__["__builtins__"] == {}


def test_literals():
    c = compiled('{1,-2;"a",TRUE}+#N/A')
    array, error = c.constants
    assert array == ((1.0, -2.0), ("a", True))
    assert error is ErrorValue.NA


def test_relative_and_absolute_references():
    assert compiled("B3", 4, 4).fn(Recorder(), 10, 20) == ("cell", None, 8, 17)
    assert compiled("$B$3", 4, 4).fn(Recorder(), 10, 20) == ("cell", None, 2, 1)
    assert compiled("B:B", 4, 4).fn(Recorder(), 10, 20) == ("columns", None, 17, 17)
    assert compiled("$3:$3", 4, 4).fn(Recorder(), 10, 20) == ("rows", None, 2, 2)


class Arithmetic:
    """Numeric operators only, enough to evaluate a long chain."""

    @staticmethod
    def add(a, b):
        return a + b

    @staticmethod
    def mul(a, b):
        return a * b

    add_top = add  # the formula's last addition (snapping is tested elsewhere)


class Cells:
    @staticmethod
    def cell(sheet, row, col):
        return float(row + 1)


def test_long_formulas_compile_flat():
    text = "+".join(f"A{i}" for i in range(1, 3001))
    c = compile_formula(parse(text), 0, 1, {}, Arithmetic())
    audit(c.source)
    assert c.fn(Cells(), 0, 1) == sum(range(1, 3001))


def test_same_template_same_source():
    a, b = compiled("A1*2+SUM(A1:A3)", 0, 1), compiled("A2*2+SUM(A2:A4)", 1, 1)
    assert template_key(parse("A1*2+SUM(A1:A3)"), 0, 1) == template_key(
        parse("A2*2+SUM(A2:A4)"), 1, 1
    )
    assert a.source == b.source
