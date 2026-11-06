"""The zip package: bounded, checked access to its parts (ARCHITECTURE.md §9.1).

Every part is read through a reader that counts decompressed bytes against ``ReadLimits`` (the
archive's own size fields are never trusted) and refuses a part that declares a DTD before
anything parses it: Open Packaging Conventions parts never declare one, so refusing it rules
out entity expansion whatever expat the Python build links. The check is expat's own reading of
the prolog, so no encoding, comment or processing instruction hides a DTD from it, and it
applies to every part, whatever its name.
"""

from __future__ import annotations

import io
import os
import posixpath
import xml.etree.ElementTree as ET
import xml.parsers.expat
import zipfile
from dataclasses import dataclass
from typing import IO, BinaryIO

__all__ = ["LimitError", "Package", "PackageError", "ReadLimits"]

_GIB = 1 << 30
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_PROLOG_BOUND = 1 << 20


class PackageError(ValueError):
    """The file is not a readable xlsx package.

    Not a zip archive, no workbook part, a part that is not well-formed XML or declares a DTD,
    a part named twice, or, when writing, a part in an encoding other than UTF-8. A subclass of
    ``ValueError``.
    """


class LimitError(PackageError):
    """The package exceeds a limit set by ``ReadLimits``.

    Too many parts, or more decompressed bytes in one part or in all of them than allowed.
    """


@dataclass(frozen=True, slots=True)
class ReadLimits:
    """Limits on what reading a package may decompress: protection against zip bombs.

    The defaults are above anything Excel can practically open, so they never refuse a
    workbook Excel opens; parts are streamed, so memory does not grow with them, and the limits
    bound the time a file can take. A service that accepts files from untrusted users should
    lower them to what it is willing to spend on one file.

    Attributes:
        max_parts: Parts (files) in the package.
        max_part_bytes: Decompressed bytes of any one part.
        max_total_bytes: Decompressed bytes of all parts read, together.
    """

    max_parts: int = 100_000
    max_part_bytes: int = 16 * _GIB
    max_total_bytes: int = 64 * _GIB


class _RootReachedError(Exception):
    pass


class _DoctypeFoundError(Exception):
    pass


def _root_reached(*_: object) -> None:
    raise _RootReachedError


def _doctype_found(*_: object) -> None:
    raise _DoctypeFoundError


class _Prolog:
    """A part's first bytes through expat, up to its root element.

    expat reads the prolog as the parsers that come later do, whatever the encoding and however
    long the comments and processing instructions before a DTD. A DTD can only come before the
    root, so nothing past it is looked at; past it, the parsers themselves reject one. A part
    expat cannot read at all (an image, say) is not XML to any parser, and passes.
    """

    def __init__(self, name: str) -> None:
        self.name = name
        self.fed = 0
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.StartDoctypeDeclHandler = _doctype_found
        self.parser.StartElementHandler = _root_reached

    def feed(self, data: bytes) -> bool:
        """Whether the part is settled; raises if it is refused."""
        self.fed += len(data)
        settled = self._parse(data, final=False)
        if not settled and self.fed >= _PROLOG_BOUND:
            raise PackageError(f"{self.name}: no root element in the first {_PROLOG_BOUND} bytes")
        return settled

    def finish(self) -> None:
        """The part ended before its root: an empty part, or not XML."""
        self._parse(b"", final=True)

    def _parse(self, data: bytes, *, final: bool) -> bool:
        try:
            self.parser.Parse(data, final)
        except _RootReachedError:
            return True
        except xml.parsers.expat.ExpatError:
            return True
        except _DoctypeFoundError:
            raise PackageError(
                f"{self.name}: declares a DTD, which a package part never does"
            ) from None
        except (ValueError, LookupError) as exc:  # a declared encoding expat cannot read
            raise PackageError(f"{self.name}: {exc}") from exc
        return False


class _Guarded(io.RawIOBase):
    """A part's decompressed bytes, counted, with the DTD check on the prolog."""

    def __init__(self, raw: IO[bytes], name: str, package: Package) -> None:
        self._raw = raw
        self._name = name
        self._package = package
        self._count = 0
        self._prolog: _Prolog | None = _Prolog(name)

    def readable(self) -> bool:
        return True

    def readinto(self, buffer: memoryview) -> int:  # type: ignore[override]
        data = self._raw.read(len(buffer))
        n = len(data)
        if n:
            self._count += n
            limits = self._package.limits
            if self._count > limits.max_part_bytes:
                raise LimitError(f"{self._name}: more than {limits.max_part_bytes} bytes")
            self._package.add_bytes(n)
            if self._prolog is not None and self._prolog.feed(data):
                self._prolog = None
        elif self._prolog is not None:
            self._prolog.finish()
            self._prolog = None
        buffer[:n] = data
        return n

    def close(self) -> None:
        self._raw.close()
        super().close()


class Package:
    """An open xlsx package. ``part`` and ``read`` give bounded access to its parts."""

    def __init__(self, source: str | os.PathLike[str] | BinaryIO, limits: ReadLimits) -> None:
        self.limits = limits
        self._total = 0
        try:
            self._zip = zipfile.ZipFile(source)
        except (zipfile.BadZipFile, OSError) as exc:
            raise PackageError(f"not a zip package: {exc}") from exc
        names = self._zip.namelist()
        if len(names) > limits.max_parts:
            self._zip.close()
            raise LimitError(f"more than {limits.max_parts} parts")
        if len(set(names)) != len(names):
            self._zip.close()
            raise PackageError("the package names a part twice")
        self.names = frozenset(names)

    def close(self) -> None:
        self._zip.close()

    def infos(self) -> list[zipfile.ZipInfo]:
        """The entries in archive order, for the writer."""
        return self._zip.infolist()

    @property
    def comment(self) -> bytes:
        """The archive's comment."""
        return self._zip.comment

    def add_bytes(self, n: int) -> None:
        self._total += n
        if self._total > self.limits.max_total_bytes:
            raise LimitError(f"more than {self.limits.max_total_bytes} bytes in all")

    def part(self, name: str) -> BinaryIO:
        """A part for streaming, buffered and bounded."""
        if name not in self.names:
            raise KeyError(name)
        return io.BufferedReader(_Guarded(self._zip.open(name), name, self), 1 << 16)

    def read(self, name: str) -> bytes:
        with self.part(name) as f:
            return f.read()

    def xml(self, name: str) -> ET.Element:
        """A small part parsed whole (workbook, relationships, tables)."""
        try:
            return ET.fromstring(self.read(name))
        except ET.ParseError as exc:
            raise PackageError(f"{name}: {exc}") from exc

    def relationships(self, part: str) -> dict[str, tuple[str, str, bool]]:
        """Relationship id → (type, target part or URL, external) for a part."""
        folder, base = posixpath.split(part)
        rels = posixpath.join(folder, "_rels", f"{base}.rels")
        if rels not in self.names:
            return {}
        out = {}
        for rel in self.xml(rels).iter(f"{_PKG_REL}Relationship"):
            target = rel.get("Target", "")
            external = rel.get("TargetMode") == "External"
            if not external:
                target = resolve(folder, target)
            out[rel.get("Id", "")] = (rel.get("Type", ""), target, external)
        return out


def resolve(folder: str, target: str) -> str:
    """A relationship target as a part name."""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(folder, target))
