"""Find the formula forms a corpus uses that this project's own formula corpus does not.

    uv run python tools/formula_census.py FORMULAS_FILE [--examples N]

Each formula gets a set of form features: reference shapes (``$A1``, column and row ranges),
prefix shapes (quoted, 3-D, book index, quoted external items), function name prefixes,
structured-reference shapes, literal shapes, operators, whitespace kinds and where they occur,
and tree constructs (union, intersection, calls of a call, array constants, omitted arguments,
nesting depth, length). The tool prints the features found in ``FORMULAS_FILE`` (one Python
``repr`` per line, as ``tools/formula_roundtrip.py --dump`` writes) and missing from
``tests/formula_corpus.py``, with counts and examples, for writing new corpus formulas by hand.
Its output describes someone else's formulas and stays local.
"""

from __future__ import annotations

import argparse
import ast
import collections
import re
import sys
from collections.abc import Iterable
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tests.formula_corpus import CORPUS

from xlrecalc._formula import FormulaSyntaxError, Kind, Token, TokenizeError, parse
from xlrecalc._formula._nodes import (
    ArrayConstant,
    Binary,
    Call,
    Group,
    Invoke,
    Missing,
    Node,
    Postfix,
    Prefix,
    Union,
    tokens,
)

_FUNCTION_PREFIX = re.compile(r"^((?:_[a-z]+\.)+)", re.I)


def _space(space: str) -> str:
    kinds = []
    if "\r\n" in space:
        kinds.append("crlf")
    if "\n" in space.replace("\r\n", ""):
        kinds.append("lf")
    if "\r" in space.replace("\r\n", ""):
        kinds.append("cr")
    if "\t" in space:
        kinds.append("tab")
    if " " in space:
        kinds.append("space" if space.count(" ") == 1 else "spaces")
    return "+".join(kinds)


def _number(text: str) -> str:
    shape = []
    if text.startswith("."):
        shape.append("leading-dot")
    if re.search(r"\.(?:[eE]|$)", text):
        shape.append("trailing-dot")
    if "." in text and not shape:
        shape.append("decimal")
    if re.search(r"[eE]", text):
        shape.append("exponent" + ("-signed" if re.search(r"[eE][+-]", text) else ""))
    if len(re.sub(r"\D", "", text.split("e")[0].split("E")[0])) > 15:
        shape.append("over-15-digits")
    return "+".join(shape) or "integer"


def _prefix(text: str) -> str:
    shape = []
    body = text[:-1]
    if body.startswith("'"):
        shape.append("quoted")
        inner = body[1:-1]
        if "''" in inner:
            shape.append("escaped-quote")
    else:
        inner = body
    if re.match(r"\[\d+\]", inner):
        shape.append("book")
        inner = re.sub(r"^\[\d+\]", "", inner)
    if inner == "":
        shape.append("book-only")
    if ":" in inner:
        shape.append("3d")
    if " " in inner:
        shape.append("space")
    if inner[:1].isdigit():
        shape.append("digit-first")
    if any(ord(ch) > 127 for ch in inner):
        shape.append("non-ascii")
    if re.search(r"[^\w .:'\[\]]", inner):
        shape.append("punctuation")
    return "+".join(shape) or "plain"


def _structured(text: str) -> str:
    table, _, rest = text.partition("[")
    shape = re.sub(r"\[(?!#|@)(?:[^\[\]']|'.)*\]", "[C]", "[" + rest)
    shape = re.sub(r"@(?!\[)(?:[^\[\]']|'.)+", "@C", shape)
    return ("T" if table else "") + shape + ("+escape" if "'" in rest else "")


def _depth(node: Node) -> int:
    best, stack = 0, [(node, 0)]
    while stack:
        item, depth = stack.pop()
        best = max(best, depth)
        match item:
            case Call(_, _, args) | Invoke(_, _, args):
                stack.extend((a, depth + 1) for a in args)
                if isinstance(item, Invoke):
                    stack.append((item.callee, depth))
            case Group(_, inner, _):
                stack.append((inner, depth + 1))
            case Prefix(_, operand) | Postfix(operand, _):
                stack.append((operand, depth))
            case Binary(left, _, right):
                stack.extend([(left, depth), (right, depth)])
            case Union(items, _):
                stack.extend((i, depth) for i in items)
    return best


def _bucket(n: int, edges: Iterable[int]) -> str:
    previous = 0
    for edge in edges:
        if n <= edge:
            return f"{previous + 1}-{edge}" if edge > previous + 1 else str(edge)
        previous = edge
    return f">{previous}"


def features(text: str) -> set[str]:
    try:
        formula = parse(text)
    except (TokenizeError, FormulaSyntaxError):
        return {"unparseable"}
    out: set[str] = set()
    toks: list[Token] = [*tokens(formula.root), formula.end]
    previous: Token | None = None
    for k, tok in enumerate(toks):
        kind = tok.kind
        if tok.space:
            where = (
                "leading" if k == 0 else
                "trailing" if kind is Kind.END else
                f"before-{tok.text}" if kind in (Kind.CLOSE, Kind.COMMA) else
                f"after-{previous.text}" if previous is not None and previous.text in "(,"
                else "between"
            )  # fmt: skip
            out.add(f"space:{_space(tok.space)}:{where}")
        if kind is Kind.CELL:
            m = re.fullmatch(r"(\$?)[A-Za-z]+(\$?)\d+", tok.text)
            assert m
            out.add(f"cell:{'$' if m[1] else ''}A{'$' if m[2] else ''}1")
            if tok.text != tok.text.upper():
                out.add("cell:lower-case")
        elif kind in (Kind.COLUMNS, Kind.ROWS):
            line_shape = re.sub(r"[A-Za-z]+|\d+", "X", tok.text)
            out.add(f"{kind.value}:{line_shape}")
        elif kind is Kind.PREFIX:
            out.add(f"prefix:{_prefix(tok.text)}")
        elif kind is Kind.NAME:
            shape = "quoted-item" if tok.text.startswith("'") else "plain"
            for ch, label in ((".", "dot"), ("?", "question"), ("\\", "backslash")):
                if ch in tok.text:
                    shape += f"+{label}"
            if tok.text.lower().startswith("_xlpm."):
                shape = "lambda-parameter"
            out.add(f"name:{shape}")
        elif kind is Kind.FUNCTION:
            m = _FUNCTION_PREFIX.match(tok.text)
            out.add(f"function-prefix:{m[1].lower() if m else 'none'}")
            name = tok.text[m.end() if m else 0 :]
            if name != name.upper():
                out.add("function:lower-case")
        elif kind is Kind.STRUCTURED:
            out.add(f"structured:{_structured(tok.text)}")
        elif kind is Kind.NUMBER:
            out.add(f"number:{_number(tok.text)}")
        elif kind is Kind.STRING:
            body = tok.text[1:-1]
            shape = []
            if body == "":
                shape.append("empty")
            if '""' in body:
                shape.append("escaped-quote")
            if any(ord(c) > 127 for c in body):
                shape.append("non-ascii")
            if "\n" in body or "\r" in body:
                shape.append("line-break")
            if len(body) > 255:
                shape.append("over-255")
            out.add(f"string:{'+'.join(shape) or 'plain'}")
        elif kind is Kind.BOOL:
            out.add("bool:" + ("upper" if tok.text.isupper() else "other-case"))
        elif kind is Kind.ERROR:
            out.add(f"error-literal:{tok.text}")
        elif kind is Kind.OPERATOR:
            out.add(f"operator:{tok.text}")
            if previous is not None and previous.kind is Kind.OPERATOR and tok.text in "+-":
                out.add("operator:unary-chain")
        previous = tok
    stack: list[Node] = [formula.root]
    while stack:
        node = stack.pop()
        match node:
            case Union(items, _):
                out.add("construct:union")
                stack.extend(items)
            case Binary(left, op, right):
                if op is None:
                    out.add("construct:intersection")
                elif op.text == ":" and not (
                    type(left).__name__ == "Reference" and type(right).__name__ == "Reference"
                ):
                    out.add("construct:range-of-expressions")
                stack.extend([left, right])
            case Invoke():
                out.add("construct:call-of-call")
                stack.extend([node.callee, *node.args])
            case Call(_, _, args):
                if not args:
                    out.add("construct:no-arguments")
                if any(isinstance(a, Missing) for a in args):
                    out.add("construct:omitted-argument")
                out.add(f"arguments:{_bucket(len(args), (1, 3, 10, 30, 255))}")
                stack.extend(args)
            case ArrayConstant(_, rows):
                shape = (
                    f"{'rows' if len(rows) > 1 else 'row'}x{'cols' if len(rows[0]) > 1 else 'col'}"
                )
                kinds = {e[-1].kind.value for r in rows for e in r}
                signed = any(len(e) == 2 for r in rows for e in r)
                out.add(f"array:{shape}:{'+'.join(sorted(kinds))}{'+signed' if signed else ''}")
            case Group(_, inner, _):
                stack.append(inner)
            case Prefix(_, operand) | Postfix(operand, _):
                stack.append(operand)
    out.add(f"depth:{_bucket(_depth(formula.root), (0, 1, 3, 7, 15, 31, 64))}")
    out.add(f"length:{_bucket(len(text), (20, 100, 255, 1000, 8192))}")
    return out


def census(texts: Iterable[str]) -> tuple[collections.Counter[str], dict[str, list[str]]]:
    counts: collections.Counter[str] = collections.Counter()
    examples: dict[str, list[str]] = collections.defaultdict(list)
    for text in texts:
        for feature in features(text):
            counts[feature] += 1
            if len(examples[feature]) < 3:
                examples[feature].append(text)
    return counts, examples


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("formulas", type=Path)
    parser.add_argument("--examples", type=int, default=2)
    args = parser.parse_args()
    if not args.formulas.exists():
        print(f"{args.formulas}: not found (the corpus is local and optional)")
        return 0
    with args.formulas.open(encoding="utf-8") as f:
        theirs = (ast.literal_eval(line) for line in f if not line.startswith("#"))
        their_counts, examples = census(theirs)
    our_counts, _ = census(CORPUS)
    missing = sorted(set(their_counts) - set(our_counts), key=lambda k: -their_counts[k])
    print(f"{len(their_counts)} features in the corpus, {len(our_counts)} in ours; missing:")
    for feature in missing:
        shown = " | ".join(repr(e[:100]) for e in examples[feature][: args.examples])
        print(f"{their_counts[feature]:9d}  {feature:45}  {shown}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
