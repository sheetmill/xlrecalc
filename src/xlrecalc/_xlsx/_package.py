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
import struct
import xml.etree.ElementTree as ET
import xml.parsers.expat
import zipfile
import zlib
from dataclasses import dataclass
from typing import IO, BinaryIO

__all__ = ["LimitError", "Package", "PackageError", "ReadLimits"]

_GIB = 1 << 30
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_PROLOG_BOUND = 1 << 20
_LOCAL_HEADER = struct.Struct("<4s22xHH")
_CHUNK = 1 << 16


class PackageError(ValueError):
    """The file is not a readable xlsx package.

    Not a zip archive, no workbook part, a part that is not well-formed XML or declares a DTD, a
    part named twice, encrypted, compressed by a method an xlsx package does not use, damaged, or
    not the size the archive states, content Excel refuses (a number it does not write, a row off
    its grid), or, when writing, a part in an encoding other than UTF-8. A subclass of
    ``ValueError``.
    """


class LimitError(PackageError):
    """The package exceeds a limit set by ``ReadLimits``.

    Too many parts, or more decompressed bytes in one part or in all of them than allowed.
    """


@dataclass(frozen=True, slots=True)
class ReadLimits:
    """Limits on what reading and computing one workbook may cost, against crafted files.

    The archive limits are above anything Excel can practically open, so they never refuse a
    workbook Excel opens; parts are streamed, so memory does not grow with them, and those limits
    bound the time a file can take. ``max_array_cells`` bounds the arrays a computation makes.
    A service that accepts files from untrusted users should lower them to what it is willing to
    spend on one file.

    Attributes:
        max_parts: Parts (files) in the package.
        max_part_bytes: Decompressed bytes of any one part.
        max_total_bytes: Decompressed bytes of all parts read, together.
        max_array_cells: Cells in any one array a formula makes while it is computed: a range
            an array formula fills, a reference taken as an array (``SUMPRODUCT(A:J*1)``), or
            the result of an operation between arrays. A formula that would pass it keeps the
            value the file holds and is reported (``Report.uncomputed``, and
            ``Report.unsupported`` names the limit). The default is ten full columns, which
            the package computes in about 2 GiB.
        max_iteration_evaluations: Cell evaluations one run of iterative calculation may make:
            the cells on and after circular references times the iterations made, up to the
            file's ``iterateCount`` (at most 32,767). A run that would pass it leaves those cells
            with the values they held, reported (``Report.uncomputed``, and
            ``Report.unsupported`` names the limit). The default is more than ten times what any
            workbook of the corpora the package is checked against needs.
    """

    max_parts: int = 100_000
    max_part_bytes: int = 16 * _GIB
    max_total_bytes: int = 64 * _GIB
    max_array_cells: int = 10_485_760
    max_iteration_evaluations: int = 100_000_000


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


class _Inflated(io.RawIOBase):
    """One entry's bytes, inflated from the archive as a stream, held to what the archive states.

    ``zipfile`` stops at the stated size, so a part larger than stated was cut short silently, and
    a smaller one surfaced as ``BadZipFile`` or ``zlib.error``. Here a part must inflate to exactly
    its stated size and CRC, else it is a ``PackageError``; the archive is read from this
    stream's own position, so several parts can be open at once.
    """

    def __init__(self, fp: IO[bytes], info: zipfile.ZipInfo) -> None:
        self._fp = fp
        self._info = info
        fp.seek(info.header_offset)
        header = fp.read(_LOCAL_HEADER.size)
        if len(header) != _LOCAL_HEADER.size or header[:4] != b"PK\x03\x04":
            raise PackageError(f"{info.filename}: no local header where the directory says")
        _, name_length, extra_length = _LOCAL_HEADER.unpack(header)
        self._position = info.header_offset + _LOCAL_HEADER.size + name_length + extra_length
        self._left = info.compress_size
        deflated = info.compress_type == zipfile.ZIP_DEFLATED
        self._inflate = zlib.decompressobj(-15) if deflated else None
        self._pending = b""
        self._produced = 0
        self._crc = 0
        self._done = False

    def readable(self) -> bool:
        return True

    def readinto(self, buffer: memoryview) -> int:  # type: ignore[override]
        while not self._done:
            data = self._next(len(buffer))
            if data:
                buffer[: len(data)] = data
                return len(data)
        return 0

    def _next(self, size: int) -> bytes:
        if self._left and not self._pending:
            self._fp.seek(self._position)
            self._pending = self._fp.read(min(self._left, _CHUNK))
            if not self._pending:
                raise PackageError(f"{self._info.filename}: the archive ends inside this part")
            self._position += len(self._pending)
            self._left -= len(self._pending)
        if self._inflate is None:
            data, self._pending = self._pending[:size], self._pending[size:]
            finished = not (self._left or self._pending)
        else:
            try:
                data = self._inflate.decompress(self._pending, size)
            except zlib.error as exc:
                raise PackageError(f"{self._info.filename}: damaged: {exc}") from exc
            self._pending = self._inflate.unconsumed_tail
            # Input used up, zlib can still hold output: the part ends when nothing more comes.
            finished = self._inflate.eof or not (data or self._pending or self._left)
        self._produced += len(data)
        if self._produced > self._info.file_size:
            self._false_size()
        self._crc = zlib.crc32(data, self._crc)
        if finished:
            self._done = True
            if self._produced != self._info.file_size or (
                self._inflate is not None and not self._inflate.eof and self._info.compress_size
            ):
                self._false_size()
            if self._crc != self._info.CRC:
                raise PackageError(f"{self._info.filename}: damaged: its CRC does not match")
        return data

    def _false_size(self) -> None:
        raise PackageError(
            f"{self._info.filename}: the archive states {self._info.file_size} bytes, which is "
            "not this part's size"
        )


class _Guarded(io.RawIOBase):
    """A part's decompressed bytes, counted, with the DTD check on the prolog."""

    def __init__(self, raw: io.RawIOBase, name: str, package: Package) -> None:
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
        try:
            self._check_directory()
        except BaseException:
            self._zip.close()
            raise
        self.names = frozenset(self._zip.namelist())

    def _check_directory(self) -> None:
        """What the central directory says, before anything is inflated."""
        infos = self._zip.infolist()
        if len(infos) > self.limits.max_parts:
            raise LimitError(f"more than {self.limits.max_parts} parts")
        if len({i.filename for i in infos}) != len(infos):
            raise PackageError("the package names a part twice")
        for info in infos:
            if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                raise PackageError(
                    f"{info.filename}: compression method {info.compress_type}, which an xlsx "
                    "package does not use"
                )
            if info.flag_bits & 1:
                raise PackageError(f"{info.filename}: encrypted")

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
        fp = self._zip.fp
        assert fp is not None
        raw = _Inflated(fp, self._zip.getinfo(name))
        return io.BufferedReader(_Guarded(raw, name, self), _CHUNK)

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
