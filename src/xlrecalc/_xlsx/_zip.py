"""Writing a zip package: entries copied byte for byte, new entries deflated (ARCHITECTURE.md §9.2).

``zipfile`` has no public way to copy an entry without decompressing and recompressing it, so
this writes the archive itself: a copied entry is its local header, compressed data and data
descriptor exactly as the source holds them, and a new entry is deflated while it streams, its
header completed afterwards by seeking back. The central directory is written from the
entries' fields, with Zip64 records where sizes, offsets or the entry count need them.
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from dataclasses import dataclass
from typing import BinaryIO

__all__ = ["ZipWriter"]

_LOCAL = struct.Struct("<4sHHHHHIIIHH")
_CENTRAL = struct.Struct("<4sBBBBHHHHIIIHHHHHII")
_END = struct.Struct("<4sHHHHIIH")
_END64 = struct.Struct("<4sQHHIIQQQQ")
_LOCATOR64 = struct.Struct("<4sIQI")
_MAX32 = 0xFFFFFFFF
_MAX16 = 0xFFFF
_UTF8 = 0x800
_DESCRIPTOR = 0x08
_CHUNK = 1 << 20


@dataclass(slots=True)
class _Entry:
    name: bytes
    flags: int
    method: int
    time: int
    date: int
    crc: int
    csize: int
    usize: int
    offset: int
    create_version: int
    create_system: int
    extract_version: int
    internal: int
    external: int
    comment: bytes


def _dos(info: zipfile.ZipInfo) -> tuple[int, int]:
    y, m, d, hh, mm, ss = info.date_time
    return (hh << 11) | (mm << 5) | (ss // 2), ((y - 1980) << 9) | (m << 5) | d


def _name(info: zipfile.ZipInfo) -> tuple[bytes, int]:
    """The entry name as the archive stores it, and the UTF-8 flag it needs."""
    if info.flag_bits & _UTF8:
        return info.orig_filename.encode("utf-8"), _UTF8
    try:
        return info.orig_filename.encode("cp437"), 0
    except UnicodeEncodeError:
        return info.orig_filename.encode("utf-8"), _UTF8


class ZipWriter:
    """A zip archive written to a seekable file, entry by entry, in order."""

    def __init__(self, fp: BinaryIO, comment: bytes = b"") -> None:
        self._fp = fp
        self._entries: list[_Entry] = []
        self._comment = comment
        self._open: _Entry | None = None
        self._compressor: zlib._Compress | None = None
        self._zip64 = False

    def copy(self, source: BinaryIO, info: zipfile.ZipInfo) -> None:
        """Copy an entry of the archive ``source`` without decompressing it."""
        source.seek(info.header_offset)
        header = source.read(_LOCAL.size)
        if len(header) != _LOCAL.size or header[:4] != b"PK\x03\x04":
            raise zipfile.BadZipFile(f"{info.filename}: no local header")
        fields = _LOCAL.unpack(header)
        length = fields[9] + fields[10] + info.compress_size
        offset = self._fp.tell()
        self._fp.write(header)
        while length:
            chunk = source.read(min(length, _CHUNK))
            if not chunk:
                raise zipfile.BadZipFile(f"{info.filename}: truncated")
            self._fp.write(chunk)
            length -= len(chunk)
        if info.flag_bits & _DESCRIPTOR:
            large = info.compress_size >= _MAX32 or info.file_size >= _MAX32
            head = source.read(4)
            size = (20 if large else 12) + (4 if head == b"PK\x07\x08" else 0)
            self._fp.write(head + source.read(size - 4))
        time, date = _dos(info)
        name, _ = _name(info)
        self._entries.append(
            _Entry(
                name,
                info.flag_bits,
                info.compress_type,
                time,
                date,
                info.CRC,
                info.compress_size,
                info.file_size,
                offset,
                info.create_version,
                info.create_system,
                info.extract_version,
                info.internal_attr,
                info.external_attr,
                info.comment,
            )
        )

    def start(self, info: zipfile.ZipInfo, zip64: bool = False) -> None:
        """Begin a new deflated entry with the metadata of ``info``; ``write`` its bytes next.

        ``zip64`` reserves room for sizes of 4 GiB and more; without it, a larger entry fails
        at ``finish``.
        """
        if self._open is not None:
            raise RuntimeError("an entry is already open")
        name, utf8 = _name(info)
        time, date = _dos(info)
        flags = (info.flag_bits & ~(_DESCRIPTOR | _UTF8 | 0x1)) | utf8
        entry = _Entry(
            name,
            flags,
            zipfile.ZIP_DEFLATED,
            time,
            date,
            0,
            0,
            0,
            self._fp.tell(),
            max(info.create_version, 20),
            info.create_system,
            45 if zip64 else 20,
            info.internal_attr,
            info.external_attr,
            info.comment,
        )
        extra = struct.pack("<HHQQ", 1, 16, 0, 0) if zip64 else b""
        size = _MAX32 if zip64 else 0
        self._fp.write(
            _LOCAL.pack(
                b"PK\x03\x04",
                entry.extract_version,
                flags,
                zipfile.ZIP_DEFLATED,
                time,
                date,
                0,
                size,
                size,
                len(name),
                len(extra),
            )
        )
        self._fp.write(name + extra)
        self._open = entry
        self._zip64 = zip64
        self._compressor = zlib.compressobj(6, zlib.DEFLATED, -15)

    def write(self, data: bytes) -> None:
        """Add bytes to the open entry."""
        entry, compressor = self._open, self._compressor
        assert entry is not None
        assert compressor is not None
        entry.crc = zlib.crc32(data, entry.crc)
        entry.usize += len(data)
        out = compressor.compress(data)
        if out:
            entry.csize += len(out)
            self._fp.write(out)

    def finish(self) -> None:
        """Complete the open entry: flush it and fill in its local header."""
        entry, compressor = self._open, self._compressor
        assert entry is not None
        assert compressor is not None
        out = compressor.flush()
        entry.csize += len(out)
        self._fp.write(out)
        end = self._fp.tell()
        if self._zip64:
            self._fp.seek(entry.offset + 14)
            self._fp.write(struct.pack("<I", entry.crc))
            self._fp.seek(entry.offset + _LOCAL.size + len(entry.name) + 4)
            self._fp.write(struct.pack("<QQ", entry.usize, entry.csize))
        else:
            if entry.usize >= _MAX32 or entry.csize >= _MAX32:
                raise zipfile.LargeZipFile(f"{entry.name!r}: 4 GiB or more without Zip64")
            self._fp.seek(entry.offset + 14)
            self._fp.write(struct.pack("<III", entry.crc, entry.csize, entry.usize))
        self._fp.seek(end)
        self._entries.append(entry)
        self._open = self._compressor = None

    def discard(self) -> None:
        """Abandon the open entry: the file is cut back to where it began."""
        entry = self._open
        assert entry is not None
        self._fp.seek(entry.offset)
        self._fp.truncate()
        self._open = self._compressor = None

    def close(self) -> None:
        """Write the central directory. The file itself is left open."""
        if self._open is not None:
            raise RuntimeError("an entry is still open")
        fp = self._fp
        start = fp.tell()
        for e in self._entries:
            extra = b""
            usize, csize, offset = e.usize, e.csize, e.offset
            big = [v for v in (usize, csize, offset) if v >= _MAX32]
            extract = e.extract_version
            if big:
                values = []
                if usize >= _MAX32:
                    values.append(usize)
                    usize = _MAX32
                if csize >= _MAX32:
                    values.append(csize)
                    csize = _MAX32
                if offset >= _MAX32:
                    values.append(offset)
                    offset = _MAX32
                extra = struct.pack(f"<HH{len(values)}Q", 1, 8 * len(values), *values)
                extract = max(extract, 45)
            fp.write(
                _CENTRAL.pack(
                    b"PK\x01\x02",
                    e.create_version,
                    e.create_system,
                    extract,
                    0,
                    e.flags,
                    e.method,
                    e.time,
                    e.date,
                    e.crc,
                    csize,
                    usize,
                    len(e.name),
                    len(extra),
                    len(e.comment),
                    0,
                    e.internal,
                    e.external,
                    offset,
                )
            )
            fp.write(e.name + extra + e.comment)
        end = fp.tell()
        count, size = len(self._entries), end - start
        if count >= _MAX16 or size >= _MAX32 or start >= _MAX32:
            fp.write(_END64.pack(b"PK\x06\x06", 44, 45, 45, 0, 0, count, count, size, start))
            fp.write(_LOCATOR64.pack(b"PK\x06\x07", 0, end, 1))
            count, size, start = min(count, _MAX16), min(size, _MAX32), min(start, _MAX32)
        fp.write(_END.pack(b"PK\x05\x06", 0, 0, count, count, size, start, len(self._comment)))
        fp.write(self._comment)
