"""Limits on what reading and computing one workbook may cost (ARCHITECTURE.md §9.1)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import ClassVar

import xlio

__all__ = ["ReadLimits", "StepLimitError", "io_limits"]

_GIB = 1 << 30


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
            ``Report.unsupported`` names the limit). It counts evaluations, not their cost
            (``max_array_cells`` bounds what one evaluation may make). The default is more than
            ten times what any workbook of the corpora the package is checked against needs
            (8.3 million); at about 0.9 million evaluations a second for the simplest formula,
            it allows about two minutes a run, so a service should lower it.
        max_function_steps: Steps that functions may take in one recalculation, an allowance
            shared by every cell it computes. A step is one iteration of a loop inside a
            function (today the recurrences of ``BESSELI``, ``BESSELJ``, ``BESSELK`` and
            ``BESSELY``, which take up to about 20 million steps a cell at the largest order
            Excel computes), charged before the loop runs. A cell whose function would pass the
            allowance keeps the value the file holds and is reported (``Report.uncomputed``, and
            ``Report.unsupported`` names the limit), as is every later cell that would. The
            default allows about a minute of such loops; a service should lower it.

    ``ReadLimits.UNTRUSTED`` is a preset for files from strangers, such as a service's uploads:
    stricter than the defaults on every field, it bounds what one file can cost to about a
    second of computing and a few hundred MiB. Its archive limits are ``xlio``'s own preset
    (10,000 parts, 256 MiB a part, 1 GiB in all); ``max_array_cells`` is one full column
    (1,048,576 cells); ``max_iteration_evaluations`` is 1,000,000 cell evaluations and
    ``max_function_steps`` 20,000,000 steps, about a second each. A workbook that needs more
    keeps the values the file holds where it passes them, and the report says which cells.
    """

    UNTRUSTED: ClassVar[ReadLimits]

    max_parts: int = 100_000
    max_part_bytes: int = 16 * _GIB
    max_total_bytes: int = 64 * _GIB
    max_array_cells: int = 10_485_760
    max_iteration_evaluations: int = 100_000_000
    max_function_steps: int = 1_000_000_000


ReadLimits.UNTRUSTED = ReadLimits(
    max_parts=xlio.ReadLimits.UNTRUSTED.max_parts,
    max_part_bytes=xlio.ReadLimits.UNTRUSTED.max_part_bytes,
    max_total_bytes=xlio.ReadLimits.UNTRUSTED.max_total_bytes,
    max_array_cells=1_048_576,
    max_iteration_evaluations=1_000_000,
    max_function_steps=20_000_000,
)


class StepLimitError(Exception):
    """A function would pass ``ReadLimits.max_function_steps``.

    Not an Excel error value: the engine catches it per cell, keeps the cell's cached value and
    reports it.
    """

    def __init__(self, limit: int) -> None:
        super().__init__(f"function steps past the limit of {limit:,}")
        self.limit = limit


def io_limits(limits: ReadLimits | None) -> xlio.ReadLimits | None:
    """The archive limits as ``xlio`` takes them; its other limits keep their defaults."""
    if limits is None:
        return None
    return xlio.ReadLimits(
        max_parts=limits.max_parts,
        max_part_bytes=limits.max_part_bytes,
        max_total_bytes=limits.max_total_bytes,
    )
