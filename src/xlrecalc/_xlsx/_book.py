"""The workbook: sheets, defined names, the date system, calculation settings, tables."""

from __future__ import annotations

import os
import posixpath
import re
import xml.parsers.expat
from collections.abc import Iterator
from dataclasses import dataclass
from types import TracebackType
from typing import Any, BinaryIO

from xlrecalc._xlsx._cells import Cell, read_cells, read_shared_strings
from xlrecalc._xlsx._package import Package, PackageError, ReadLimits

__all__ = ["Book", "CalcProperties", "DefinedName", "SheetInfo", "Table"]

_MAINS = (
    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}",
    "{http://purl.oclc.org/ooxml/spreadsheetml/main}",
)
_RELS = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}",
    "{http://purl.oclc.org/ooxml/officeDocument/relationships}",
)
_TRUE = ("1", "true")
# Rich-value error types, as captured (capture arrays: a blocked spill is errorType 8).
_RICH_ERRORS = {"8": "#SPILL!"}


@dataclass(frozen=True, slots=True)
class SheetInfo:
    """A sheet: name, part, ``visible``/``hidden``/``veryHidden``, and its kind.

    ``kind`` is the last segment of the relationship type: ``worksheet``, ``chartsheet``,
    ``dialogsheet`` or ``xlMacrosheet``; only worksheets have cells to read.
    """

    name: str
    part: str
    state: str
    kind: str


@dataclass(frozen=True, slots=True)
class DefinedName:
    """A defined name: its text (a formula, file form), scope (sheet index or None), hidden."""

    name: str
    text: str
    sheet: int | None
    hidden: bool


@dataclass(frozen=True, slots=True)
class CalcProperties:
    """``<calcPr>`` as written; ``None`` where the file leaves an attribute out."""

    calc_id: int | None = None
    full_calc_on_load: bool = False
    calc_mode: str | None = None
    iterate: bool = False
    iterate_count: int | None = None
    iterate_delta: float | None = None


@dataclass(frozen=True, slots=True)
class Table:
    """A table (for structured references): names, range, columns, header and totals rows."""

    name: str
    display_name: str
    ref: str
    columns: tuple[str, ...]
    header_rows: int
    totals_rows: int


def _find(element: object, local: str) -> list[object]:
    return [e for ns in _MAINS for e in element.iter(ns + local)]  # type: ignore[attr-defined]


def _local(element: object, local: str) -> list[Any]:
    """Descendants with this local name, in any namespace."""
    return [
        e
        for e in element.iter()  # type: ignore[attr-defined]
        if isinstance(e.tag, str) and e.tag.rpartition("}")[2] == local
    ]


def _attr(element: object, local: str) -> str | None:
    for ns in _RELS:
        value = element.get(ns + local)  # type: ignore[attr-defined]
        if value is not None:
            return str(value)
    return None


class Book:
    """An xlsx workbook opened for reading. Use as a context manager, or ``close`` it."""

    def __init__(
        self,
        source: str | os.PathLike[str] | BinaryIO,
        limits: ReadLimits | None = None,
    ) -> None:
        self._package = Package(source, limits or ReadLimits())
        try:
            self._load()
        except BaseException:
            self._package.close()
            raise
        self._strings: list[str] | None = None

    def __enter__(self) -> Book:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        """Close the file."""
        self._package.close()

    @property
    def package(self) -> Package:
        """The package, for the writer."""
        return self._package

    def _load(self) -> None:
        package = self._package
        root_rels = package.relationships("")
        workbook = next(
            (
                t
                for typ, t, ext in root_rels.values()
                if typ.endswith("/officeDocument") and not ext
            ),
            None,
        )
        if workbook is None or workbook not in package.names:
            raise PackageError("the package has no workbook part")
        self.workbook_part = workbook
        rels = package.relationships(workbook)
        root = package.xml(workbook)
        self.sheets: list[SheetInfo] = []
        for sheet in _find(root, "sheet"):
            rel = rels.get(_attr(sheet, "id") or "")
            if rel is None:
                continue
            kind = rel[0].rsplit("/", 1)[-1]
            self.sheets.append(
                SheetInfo(
                    sheet.get("name", ""),  # type: ignore[attr-defined]
                    rel[1],
                    sheet.get("state", "visible"),  # type: ignore[attr-defined]
                    kind,
                )
            )
        views = _find(root, "workbookView")
        tab = views[0].get("activeTab", "0") if views else "0"  # type: ignore[attr-defined]
        self.active_tab = int(tab) if tab.isdigit() else 0
        """The sheet the workbook opens on (``workbookView activeTab``), 0-based."""
        props = _find(root, "workbookPr")
        self.date1904 = bool(props) and props[0].get("date1904", "0") in _TRUE  # type: ignore[attr-defined]
        calc = _find(root, "calcPr")
        self.calc = _calc(calc[0]) if calc else CalcProperties()
        self.defined_names: list[DefinedName] = []
        for dn in _find(root, "definedName"):
            local = dn.get("localSheetId")  # type: ignore[attr-defined]
            self.defined_names.append(
                DefinedName(
                    dn.get("name", ""),  # type: ignore[attr-defined]
                    dn.text or "",  # type: ignore[attr-defined]
                    int(local) if local is not None else None,
                    dn.get("hidden", "0") in _TRUE,  # type: ignore[attr-defined]
                )
            )
        self._strings_part = next(
            (t for typ, t, ext in rels.values() if typ.endswith("/sharedStrings") and not ext),
            None,
        )
        self.dynamic_cm: frozenset[str] = frozenset()
        """The ``cm`` values that mark a dynamic-array formula (metadata type ``XLDAPR``)."""
        self.value_errors: dict[str, str] = {}
        """``vm`` value -> the error a rich value holds, for errors a file caches as
        ``#VALUE!`` (``#SPILL!``)."""
        meta = next(
            (t for typ, t, ext in rels.values() if typ.endswith("/sheetMetadata") and not ext),
            None,
        )
        if meta is not None and meta in package.names:
            self._metadata(meta, rels)
        self.external_links = [
            t for typ, t, ext in rels.values() if typ.endswith("/externalLink") and not ext
        ]

    def _metadata(self, part: str, rels: dict[str, tuple[str, str, bool]]) -> None:
        """The dynamic-array and rich-value metadata (capture ``arrays``)."""
        root = self._package.xml(part)
        types = [t.get("name", "") for t in _local(root, "metadataType")]
        future: dict[str, list[object]] = {}
        for fm in _local(root, "futureMetadata"):
            future[fm.get("name", "")] = list(_local(fm, "bk"))

        def entries(block: str) -> list[tuple[str, int]]:
            out = []
            for container in _local(root, block):
                for bk in _local(container, "bk"):
                    rc = next(iter(_local(bk, "rc")), None)
                    if rc is None:
                        out.append(("", 0))
                        continue
                    t = int(rc.get("t", "0"))
                    v = int(rc.get("v", "0"))
                    out.append((types[t - 1] if 0 < t <= len(types) else "", v))
            return out

        dynamic = set()
        for k, (name, v) in enumerate(entries("cellMetadata"), 1):
            blocks = future.get(name, [])
            if name == "XLDAPR" and v < len(blocks):
                props = next(iter(_local(blocks[v], "dynamicArrayProperties")), None)
                if props is not None and props.get("fDynamic", "0") in _TRUE:
                    dynamic.add(str(k))
        self.dynamic_cm = frozenset(dynamic)
        values = self._rich_errors(rels)
        for k, (name, v) in enumerate(entries("valueMetadata"), 1):
            blocks = future.get(name, [])
            if name == "XLRICHVALUE" and v < len(blocks):
                rvb = next(iter(_local(blocks[v], "rvb")), None)
                if rvb is not None:
                    error = values.get(int(rvb.get("i", "-1")))
                    if error is not None:
                        self.value_errors[str(k)] = error

    def _rich_errors(self, rels: dict[str, tuple[str, str, bool]]) -> dict[int, str]:
        """Rich value index -> the error it holds, for the error types known from captures."""
        parts = {typ.rsplit("/", 1)[-1].lower(): t for typ, t, ext in rels.values() if not ext}
        values_part = parts.get("rdrichvalue")
        structures_part = parts.get("rdrichvaluestructure")
        if values_part not in self._package.names or structures_part not in self._package.names:
            return {}
        assert values_part is not None
        assert structures_part is not None
        structures = []
        for st in _local(self._package.xml(structures_part), "s"):
            keys = [k.get("n", "") for k in _local(st, "k")]
            structures.append((st.get("t", ""), keys))
        out = {}
        for i, rv in enumerate(_local(self._package.xml(values_part), "rv")):
            s = int(rv.get("s", "-1"))
            if not 0 <= s < len(structures) or structures[s][0] != "_error":
                continue
            fields = dict(zip(structures[s][1], (v.text for v in _local(rv, "v")), strict=False))
            code = _RICH_ERRORS.get(fields.get("errorType") or "")
            if code is not None:
                out[i] = code
        return out

    @property
    def shared_strings(self) -> list[str]:
        """The shared string table, read once."""
        if self._strings is None:
            part = self._strings_part
            if part is None or part not in self._package.names:
                self._strings = []
            else:
                with self._package.part(part) as f:
                    try:
                        self._strings = read_shared_strings(f)
                    except xml.parsers.expat.ExpatError as exc:
                        raise PackageError(f"{part}: {exc}") from exc
        return self._strings

    def sheet(self, name: str) -> SheetInfo:
        """The sheet with this name (Excel compares sheet names without case)."""
        folded = name.casefold()
        for sheet in self.sheets:
            if sheet.name.casefold() == folded:
                return sheet
        raise KeyError(f"no sheet named {name!r}")

    def cells(self, sheet: str | SheetInfo) -> Iterator[Cell]:
        """The cells of a worksheet, streamed in file order."""
        info = sheet if isinstance(sheet, SheetInfo) else self.sheet(sheet)
        if info.kind != "worksheet":
            raise ValueError(f"{info.name!r} is a {info.kind}, not a worksheet")
        strings = self.shared_strings
        with self._package.part(info.part) as f:
            try:
                yield from read_cells(f, strings)
            except xml.parsers.expat.ExpatError as exc:
                raise PackageError(f"{info.part}: {exc}") from exc

    def active_cell(self, sheet: str | SheetInfo) -> tuple[int, int]:
        """A worksheet's active cell (``sheetView`` ``selection activeCell``), 1-based (row, col).

        The selection of the active pane when there are several; A1 when the sheet names none.
        Only the part's head is read, up to ``sheetData``.
        """
        info = sheet if isinstance(sheet, SheetInfo) else self.sheet(sheet)
        if info.kind != "worksheet" or info.part not in self._package.names:
            return 1, 1
        found: dict[str, str] = {}
        pane = ["topLeft"]
        first: list[str] = []

        class _HeadReadError(Exception):
            pass

        def start(name: str, attrs: dict[str, str]) -> None:
            local = name.rpartition(":")[2]
            if local == "sheetData":
                raise _HeadReadError
            if local == "pane" and "activePane" in attrs:
                pane[0] = attrs["activePane"]
            elif local == "selection" and "activeCell" in attrs:
                found[attrs.get("pane", "topLeft")] = attrs["activeCell"]
                if not first:
                    first.append(attrs["activeCell"])

        p = xml.parsers.expat.ParserCreate()
        p.StartElementHandler = start
        with self._package.part(info.part) as f:
            try:
                while chunk := f.read(1 << 14):
                    p.Parse(chunk, False)
            except (_HeadReadError, xml.parsers.expat.ExpatError):
                pass
        cell = found.get(pane[0]) or (first[0] if first else "A1")
        m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", cell.split(" ")[0].split(":")[0])
        if m is None:
            return 1, 1
        col = 0
        for ch in m[1].upper():
            col = col * 26 + ord(ch) - 64
        return int(m[2]), col

    def tables(self, sheet: str | SheetInfo) -> list[Table]:
        """The tables on a worksheet."""
        info = sheet if isinstance(sheet, SheetInfo) else self.sheet(sheet)
        out = []
        for typ, target, external in self._package.relationships(info.part).values():
            if external or not typ.endswith("/table") or target not in self._package.names:
                continue
            root = self._package.xml(target)
            columns = tuple(c.get("name", "") for c in _find(root, "tableColumn"))  # type: ignore[attr-defined]
            out.append(
                Table(
                    root.get("name", ""),
                    root.get("displayName", root.get("name", "")),
                    root.get("ref", ""),
                    columns,
                    int(root.get("headerRowCount", "1")),
                    int(root.get("totalsRowCount", "0")),
                )
            )
        return out

    def part_name(self, relative: str) -> str:
        """A part name relative to the workbook part's folder."""
        return posixpath.normpath(posixpath.join(posixpath.dirname(self.workbook_part), relative))


def _calc(element: object) -> CalcProperties:
    get = element.get  # type: ignore[attr-defined]
    calc_id = get("calcId")
    count = get("iterateCount")
    delta = get("iterateDelta")
    return CalcProperties(
        calc_id=int(calc_id) if calc_id is not None else None,
        full_calc_on_load=get("fullCalcOnLoad", "0") in _TRUE,
        calc_mode=get("calcMode"),
        iterate=get("iterate", "0") in _TRUE,
        iterate_count=int(count) if count is not None else None,
        iterate_delta=float(delta) if delta is not None else None,
    )
