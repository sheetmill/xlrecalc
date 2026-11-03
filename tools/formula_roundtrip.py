"""Round-trip every formula in a set of workbooks through the parser and printer.

    uv run python tools/formula_roundtrip.py FILE_OR_DIRECTORY ...

Reads each ``.xlsx`` (never ``.xlsm``) with the standard library: the ``<f>`` elements of every
worksheet part, as stored in the file (a shared formula's dependent cells hold no text and are
skipped). Prints, per run, how many formulas were read, how many distinct texts, how many
printed back unchanged, the failures with their reason, and the parse time. Files are only
read; nothing is evaluated and nothing is written, except that ``--dump FILE`` writes the
distinct formula texts, one Python ``repr`` per line, in first-seen order. ``--functions`` also
prints, per function name, the number of workbooks using it and the number of formula cells
calling it, ordered by workbooks, for ordering the function library.
"""

from __future__ import annotations

import argparse
import collections
import contextlib
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from xlrecalc._formula import FormulaSyntaxError, Kind, TokenizeError, parse, tokenize

_F = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f"


def formulas(path: Path) -> list[str]:
    found = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
                continue
            with archive.open(name) as part:
                head = part.read(4096)
            if b"<!DOCTYPE" in head:
                raise ValueError(f"{path}: {name} declares a DTD")
            with archive.open(name) as part:
                for _, element in ET.iterparse(part):
                    if element.tag == _F and element.text:
                        found.append(element.text)
                    element.clear()
    return found


def _function_names(texts: list[str]) -> set[str]:
    names: set[str] = set()
    for text in texts:
        with contextlib.suppress(TokenizeError):
            names.update(t.text.upper() for t in tokenize(text) if t.kind is Kind.FUNCTION)
    return names


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--functions", action="store_true", help="print function counts")
    parser.add_argument("--show", type=int, default=20, help="failures to print")
    parser.add_argument("--dump", type=Path, help="write the distinct formulas to this file")
    args = parser.parse_args()
    files = []
    for p in args.paths:
        if not p.exists():
            print(f"{p}: not found (local corpora are optional)")
            continue
        files.extend(sorted(p.rglob("*.xlsx")) if p.is_dir() else [p])
    if not files:
        return 0
    texts: list[str] = []
    unreadable = 0
    workbooks: collections.Counter[str] = collections.Counter()
    for f in files:
        try:
            found = formulas(f)
            texts.extend(found)
            if args.functions:
                workbooks.update(_function_names(found))
        except (zipfile.BadZipFile, ET.ParseError, ValueError, KeyError) as exc:
            unreadable += 1
            print(f"unreadable: {f}: {exc}", file=sys.stderr)
    distinct = list(dict.fromkeys(texts))
    if args.dump:
        args.dump.write_text("".join(f"{text!r}\n" for text in distinct), encoding="utf-8")
    failures: list[tuple[str, str]] = []
    counts: collections.Counter[str] = collections.Counter()
    start = time.perf_counter()
    for text in distinct:
        try:
            printed = parse(text).text()
        except (FormulaSyntaxError, TokenizeError) as exc:
            failures.append((text, str(exc)))
            continue
        if printed != text:
            failures.append((text, f"printed as {printed!r}"))
    elapsed = time.perf_counter() - start
    if args.functions:
        for text in texts:
            with contextlib.suppress(TokenizeError):
                counts.update(t.text.upper() for t in tokenize(text) if t.kind is Kind.FUNCTION)
    print(
        f"{len(files)} files ({unreadable} unreadable), {len(texts)} formulas, "
        f"{len(distinct)} distinct, {len(distinct) - len(failures)} round-tripped, "
        f"{len(failures)} failed; parsed in {elapsed:.3f} s "
        f"({elapsed / max(len(distinct), 1) * 1e6:.1f} µs per formula)"
    )
    for text, reason in failures[: args.show]:
        print(f"  {text!r}: {reason}")
    if args.functions:
        print(f"{'workbooks':>9} {'cells':>9} function")
        for name, n in workbooks.most_common():
            print(f"{n:9d} {counts[name]:9d} {name}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
