"""Unlock a .xls (BIFF8) workbook.

Every touched record keeps its original size, so BOUNDSHEET offsets to each
sheet's BOF stay valid and the whole stream is byte-length-invariant.
"""
from __future__ import annotations
import io
import shutil
import struct
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import olefile

from . import _biff as biff
from .vba import unlock_vba_project_ole


DEFAULT_XLS_PASSWORDS = ("VelvetSweatshop", "")


@dataclass
class XlsReport:
    input_path: Path
    output_path: Path
    open_password_used: Optional[str] = None
    was_encrypted: bool = False
    vba_unlocked: bool = False
    touched: dict[int, int] = field(default_factory=dict)   # rec_id -> count
    stream_size: int = 0

    def bump(self, rec_id: int) -> None:
        self.touched[rec_id] = self.touched.get(rec_id, 0) + 1


def _decrypt_to_bytes(src: Path, passwords: tuple[str, ...]) -> tuple[bytes, Optional[str]]:
    """Return (plaintext OLE bytes, password that worked) — or the file's own
    bytes if it isn't encrypted."""
    import msoffcrypto

    with open(src, "rb") as f:
        head = f.read()

    # Fast probe: if there's no FILEPASS record near the start of Workbook,
    # it isn't encrypted; skip msoffcrypto to avoid its warnings.
    try:
        ole = olefile.OleFileIO(io.BytesIO(head))
        wb = ole.openstream("Workbook").read()
        ole.close()
        encrypted = any(r.rec_id == biff.FILEPASS for r in biff.walk(wb))
    except Exception:
        encrypted = True   # be pessimistic — let msoffcrypto try

    if not encrypted:
        return head, None

    last_err: Optional[Exception] = None
    for pw in passwords:
        try:
            of = msoffcrypto.OfficeFile(io.BytesIO(head))
            of.load_key(password=pw)
            out = io.BytesIO()
            of.decrypt(out)
            return out.getvalue(), pw
        except Exception as e:
            last_err = e
            continue

    raise RuntimeError(
        f"Could not decrypt {src.name}. Tried passwords: {passwords!r}. "
        f"Last error: {last_err!r}"
    )


def _patch_workbook_stream(data: bytearray, report: XlsReport) -> None:
    """Zero protection record bodies; fix WINDOW1/WINDOW2 display flags."""
    original_size = len(data)

    for rec in biff.walk(bytes(data)):
        b0, b1 = rec.body_start, rec.body_end

        if rec.rec_id in biff.ZERO_BODY_PROT_RECS:
            if rec.size:
                data[b0:b1] = b"\x00" * rec.size
                report.bump(rec.rec_id)

        elif rec.rec_id == biff.SHEETPROTECTION and rec.size >= 23:
            # Feat record body for ISFPROTECTION:
            #   [0..12)  FrtRefHeader (record type dup + Ref8U — leave as-is)
            #   [12..14) isf            — keep 0x0002 (ISFPROTECTION), a valid id
            #   [14..15) fHdr           — 0 (no encapsulated header)
            #   [15..19) reserved3      — 0
            #   [19..23) FeatProtection — 4-byte payload; set to all-zero means
            #                            "no password, all operations allowed"
            # This produces a well-formed record that Excel will not "repair".
            struct.pack_into("<HB", data, b0 + 12, 0x0002, 0x00)
            data[b0 + 15 : b1] = b"\x00" * (rec.size - 15)
            report.bump(rec.rec_id)

        elif rec.rec_id == biff.WINDOW1 and rec.size >= 10:
            # grbit at offset 8 (after xWn, yWn, dxWn, dyWn = 8 bytes)
            grbit = struct.unpack_from("<H", data, b0 + 8)[0]
            new_grbit = grbit
            new_grbit &= ~0x0001   # fHidden -> off
            new_grbit &= ~0x0002   # fIconic -> off
            new_grbit |=  0x0008   # fDspHScroll -> on
            new_grbit |=  0x0010   # fDspVScroll -> on
            new_grbit |=  0x0020   # fDspTabs -> on
            if new_grbit != grbit:
                struct.pack_into("<H", data, b0 + 8, new_grbit)
                report.bump(rec.rec_id)

        elif rec.rec_id == biff.WINDOW2 and rec.size >= 2:
            # Force display bits ON:
            #   0x0002 fDspGrid, 0x0004 fDspRwCol, 0x0010 fDspZeros, 0x0080 fDspGuts
            # Never touch fFrozen (0x0008), fDspFmla (0x0001), fSelected (0x0400).
            flags = struct.unpack_from("<H", data, b0)[0]
            new_flags = flags | 0x0002 | 0x0004 | 0x0010 | 0x0080
            if new_flags != flags:
                struct.pack_into("<H", data, b0, new_flags)
                report.bump(rec.rec_id)
            # BIFF8 WINDOW2 (size 18) has wScaleSLV at offset 14 and
            # wScaleNormal at offset 16. Force both to 0 = "use SCL / default 100".
            if rec.size >= 16:
                (wSLV,) = struct.unpack_from("<H", data, b0 + 14)
                if wSLV != 0:
                    struct.pack_into("<H", data, b0 + 14, 0)
                    report.bump(rec.rec_id)
            if rec.size >= 18:
                (wNorm,) = struct.unpack_from("<H", data, b0 + 16)
                if wNorm != 0:
                    struct.pack_into("<H", data, b0 + 16, 0)
                    report.bump(rec.rec_id)

        elif rec.rec_id == biff.SCL and rec.size >= 4:
            # Force zoom = 1/1 = 100%.
            num, den = struct.unpack_from("<HH", data, b0)
            if (num, den) != (1, 1):
                struct.pack_into("<HH", data, b0, 1, 1)
                report.bump(rec.rec_id)

    assert len(data) == original_size, "Workbook stream length changed — refusing to save."
    report.stream_size = len(data)


def unlock_xls(
    src: Path,
    dst: Path,
    *,
    passwords: tuple[str, ...] = DEFAULT_XLS_PASSWORDS,
    also_unlock_vba: bool = True,
) -> XlsReport:
    """Unlock an .xls workbook. Returns a report of what was changed."""
    src = Path(src)
    dst = Path(dst)
    report = XlsReport(input_path=src, output_path=dst)

    plaintext_ole_bytes, pw = _decrypt_to_bytes(src, passwords)
    report.was_encrypted = pw is not None
    report.open_password_used = pw

    # Write the decrypted OLE bytes to dst; then patch in place.
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(plaintext_ole_bytes)

    ole = olefile.OleFileIO(dst, write_mode=True)
    try:
        stream_name = "Workbook" if ole.exists("Workbook") else "Book"
        data = bytearray(ole.openstream(stream_name).read())
        _patch_workbook_stream(data, report)
        ole.write_stream(stream_name, bytes(data))
    finally:
        ole.close()

    if also_unlock_vba:
        report.vba_unlocked = unlock_vba_project_ole(dst) is True

    return report
