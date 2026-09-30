"""BIFF8 record walker and IDs. No I/O here — pure byte pushing."""
from __future__ import annotations
import struct
from typing import Iterator, NamedTuple

# BIFF record IDs we care about
BOF            = 0x0809
EOF            = 0x000A
FILEPASS       = 0x002F
PROTECT        = 0x0012
PASSWORD       = 0x0013
WINDOWPROTECT  = 0x0019
SCENPROTECT    = 0x00DD
PROT4REV       = 0x0086
SHEETPROTECTION = 0x0867  # Feat record used for ISFPROTECTION
BOUNDSHEET     = 0x0085
WINDOW1        = 0x003D
WINDOW2        = 0x023E
SCL            = 0x00A0   # zoom (num/den)
NAME           = 0x0018

# 2-byte body records whose body should be zeroed to indicate "no protection"
ZERO_BODY_PROT_RECS = {PROTECT, PASSWORD, WINDOWPROTECT, SCENPROTECT, PROT4REV}

RECORD_NAMES = {
    BOF: "BOF", EOF: "EOF", FILEPASS: "FILEPASS",
    PROTECT: "PROTECT", PASSWORD: "PASSWORD",
    WINDOWPROTECT: "WINDOWPROTECT", SCENPROTECT: "SCENPROTECT",
    PROT4REV: "PROT4REV", SHEETPROTECTION: "SheetProtection",
    BOUNDSHEET: "BOUNDSHEET", WINDOW1: "WINDOW1", WINDOW2: "WINDOW2",
    NAME: "NAME",
}


class Record(NamedTuple):
    """A single BIFF record inside the workbook stream."""
    offset: int   # byte offset of the record header
    rec_id: int
    size: int     # size of the body (bytes 4..4+size)

    @property
    def body_start(self) -> int:
        return self.offset + 4

    @property
    def body_end(self) -> int:
        return self.body_start + self.size


def walk(stream: bytes) -> Iterator[Record]:
    """Yield every top-level BIFF record in a workbook stream."""
    i = 0
    n = len(stream)
    while i + 4 <= n:
        rec_id, size = struct.unpack_from("<HH", stream, i)
        if size > n - i - 4:
            return
        yield Record(i, rec_id, size)
        i += 4 + size


def name(rec_id: int) -> str:
    return RECORD_NAMES.get(rec_id, f"0x{rec_id:04X}")
