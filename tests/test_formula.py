"""Tokenizer, parser, exact printer and template keys (ARCHITECTURE.md §5.2 and §5.3)."""

import pytest

from tests.formula_corpus import CORPUS, INVALID
from xlrecalc._formula import (
    FormulaSyntaxError,
    Kind,
    TokenizeError,
    parse,
    template_key,
    tokenize,
)
from xlrecalc._formula._nodes import (
    ArrayConstant,
    Binary,
    Call,
    Group,
    Invoke,
    Leaf,
    Missing,
    Node,
    Postfix,
    Prefix,
    Reference,
    Union,
    tokens,
)
from xlrecalc._formula._parser import MAX_DEPTH


@pytest.mark.parametrize("text", CORPUS, ids=repr)
def test_tokens_keep_every_character(text):
    assert "".join(t.space + t.text for t in tokenize(text)) == text


@pytest.mark.parametrize("text", CORPUS, ids=repr)
def test_round_trip(text):
    assert parse(text).text() == text


@pytest.mark.parametrize("text", INVALID, ids=repr)
def test_invalid(text):
    with pytest.raises((FormulaSyntaxError, TokenizeError)):
        parse(text)


def sexp(node: Node) -> str:
    """The tree as an s-expression, with token texts, to check its shape."""
    match node:
        case Leaf(token):
            return token.text
        case Reference(token, prefix):
            return (prefix.text if prefix else "") + token.text
        case Missing():
            return "_"
        case Prefix(op, operand):
            return f"({op.text} {sexp(operand)})"
        case Postfix(operand, op):
            return f"({sexp(operand)} {op.text})"
        case Binary(left, op, right):
            return f"({op.text if op else 'isect'} {sexp(left)} {sexp(right)})"
        case Group(_, inner, _):
            return sexp(inner)
        case Union(items, _):
            return "(union " + " ".join(sexp(i) for i in items) + ")"
        case Call(name, _, args, _, _, prefix):
            head = (prefix.text if prefix else "") + name.text
            return f"[{head} " + " ".join(sexp(a) for a in args) + "]"
        case Invoke(callee, _, args):
            return f"[{sexp(callee)} " + " ".join(sexp(a) for a in args) + "]"
        case ArrayConstant(_, rows):
            return (
                "{" + ";".join(",".join("".join(t.text for t in e) for e in r) for r in rows) + "}"
            )
    raise AssertionError(node)


@pytest.mark.parametrize(
    ("text", "tree"),
    [
        ("1+2*3", "(+ 1 (* 2 3))"),
        ("1*2+3", "(+ (* 1 2) 3)"),
        ("1-2-3", "(- (- 1 2) 3)"),
        ("2^3^2", "(^ (^ 2 3) 2)"),
        ("-2^2", "(^ (- 2) 2)"),
        ("2^-2", "(^ 2 (- 2))"),
        ("-50%", "((- 50) %)"),
        ("2^3%", "(^ 2 (3 %))"),
        ("1+2&3", "(& (+ 1 2) 3)"),
        ("1&2=12", "(= (& 1 2) 12)"),
        ("1<2=TRUE", "(= (< 1 2) TRUE)"),
        ("-A1:B2", "(- (: A1 B2))"),
        ("@A1:A10", "(@ (: A1 A10))"),
        ("A1:B2 B1:C3", "(isect (: A1 B2) (: B1 C3))"),
        ("-A1 B1", "(- (isect A1 B1))"),
        ("A1#", "(A1 #)"),
        ("Sheet2!A1#*2", "(* (Sheet2!A1 #) 2)"),
        ("SUM((A1,B1),C1)", "[SUM (union A1 B1) C1]"),
        ("SUM(A1,B1)", "[SUM A1 B1]"),
        ("IF(A1,,1)", "[IF A1 _ 1]"),
        ("SUM(,)", "[SUM _ _]"),
        ("NOW()", "[NOW ]"),
        ("Sheet2!MyFunc(1)", "[Sheet2!MyFunc 1]"),
        ("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", "[[_xlfn.LAMBDA _xlpm.x (+ _xlpm.x 1)] 2]"),
        ("INDEX(A1:C3,1,1):C3", "(: [INDEX (: A1 C3) 1 1] C3)"),
        ("{1,-2;3,4}", "{1,-2;3,4}"),
        ("'It''s'!A1+[1]Sheet1!B2", "(+ 'It''s'!A1 [1]Sheet1!B2)"),
        ("Sheet1!#REF!+1", "(+ Sheet1!#REF! 1)"),
        ("Table1[[#This Row],[Price]]*2", "(* Table1[[#This Row],[Price]] 2)"),
    ],
)
def test_tree(text, tree):
    assert sexp(parse(text).root) == tree


@pytest.mark.parametrize(
    ("text", "kinds"),
    [
        ("A1", [Kind.CELL]),
        ("XFD1048576", [Kind.CELL]),
        ("XFE1", [Kind.NAME]),
        ("A1048577", [Kind.NAME]),
        ("LOG10(1)", [Kind.FUNCTION, Kind.OPEN, Kind.NUMBER, Kind.CLOSE]),
        ("A:C", [Kind.COLUMNS]),
        ("1:3", [Kind.ROWS]),
        ("TRUE", [Kind.BOOL]),
        ("TRUE()", [Kind.FUNCTION, Kind.OPEN, Kind.CLOSE]),
        ("#N/A", [Kind.ERROR]),
        ("A1#", [Kind.CELL, Kind.OPERATOR]),
        ("Sheet1!A1", [Kind.PREFIX, Kind.CELL]),
        ("[1]!Rate", [Kind.PREFIX, Kind.NAME]),
        ("[1]!'A, B'", [Kind.PREFIX, Kind.NAME]),
        ("[@Price]", [Kind.STRUCTURED]),
        ("Table1[[Col'[1']]]", [Kind.STRUCTURED]),
    ],
)
def test_token_kinds(text, kinds):
    assert [t.kind for t in tokenize(text)][:-1] == kinds


def test_nesting_limit_is_a_syntax_error():
    ok = "(" * (MAX_DEPTH - 1) + "1" + ")" * (MAX_DEPTH - 1)
    assert parse(ok).text() == ok
    deep = "(" * (MAX_DEPTH + 1) + "1" + ")" * (MAX_DEPTH + 1)
    with pytest.raises(FormulaSyntaxError, match="nested too deeply"):
        parse(deep)
    with pytest.raises(FormulaSyntaxError, match="nested too deeply"):
        parse("-" * (MAX_DEPTH + 1) + "1")


def test_long_chain_needs_no_recursion():
    text = "+".join(f"A{i}" for i in range(1, 5001))
    assert parse(text).text() == text


def test_rewriting_a_reference_changes_only_its_token():
    text = "SUM( A1:A10 ,\n  Sheet2!$B$2 )*2"
    formula = parse(text)
    for tok in tokens(formula.root):
        if tok.text == "$B$2":
            tok.text = "$C$5"
    assert formula.text() == "SUM( A1:A10 ,\n  Sheet2!$C$5 )*2"


def test_template_keys():
    # Shared formulas: the same relative shape in different cells.
    assert template_key(parse("A1*2"), 0, 1) == template_key(parse("A2*2"), 1, 1)
    assert template_key(parse("$A$1*2"), 0, 1) == template_key(parse("$A$1*2"), 5, 1)
    assert template_key(parse("$A$1*2"), 0, 1) != template_key(parse("$A$2*2"), 0, 1)
    assert template_key(parse("SUM(A1:A3)"), 3, 0) == template_key(parse("SUM(A2:A4)"), 4, 0)
    assert template_key(parse("A1*2"), 0, 1) != template_key(parse("A1*2"), 1, 1)
    # Whitespace and the case of names and functions do not matter; strings do.
    assert template_key(parse("sum( a1 )"), 0, 1) == template_key(parse("SUM(A1)"), 0, 1)
    assert template_key(parse('"a"'), 0, 0) != template_key(parse('"A"'), 0, 0)
    # Adjacent tokens stay apart.
    assert template_key(parse("1&23"), 0, 0) != template_key(parse("12&3"), 0, 0)
    # Columns and rows are relative unless anchored.
    assert template_key(parse("SUM(B:B)"), 0, 0) == template_key(parse("SUM(C:C)"), 0, 1)
    assert template_key(parse("SUM($B:$B)"), 0, 0) != template_key(parse("SUM($C:$C)"), 0, 1)
