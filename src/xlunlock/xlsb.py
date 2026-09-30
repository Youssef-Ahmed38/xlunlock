"""Unlock a .xlsb workbook (OOXML binary, zip of BRT records)."""
from __future__ import annotations
import io
import shutil
import struct
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from .vba import unlock_vba_bytes


# BRT record IDs (see [MS-XLSB])
BRT_BEGIN_SHEET_PROTECTION       = 0x0226   # BrtSheetProtection
BRT_SHEET_PROTECTION_ISO         = 0x08F1   # BrtSheetProtectionIso
BRT_BOOK_PROTECTION              = 0x0224   # BrtBookProtection
BRT_BOOK_PROTECTION_ISO          = 0x08EF   # BrtBookProtectionIso

PROT_RECS = {
    BRT_BEGIN_SHEET_PROTECTION,
    BRT_SHEET_PROTECTION_ISO,
    BRT_BOOK_PROTECTION,
    BRT_BOOK_PROTECTION_ISO,
}


@dataclass
class XlsbReport:
    input_path: Path
    output_path: Path
    was_encrypted: bool = False
    open_password_used: Optional[str] = None
    sheets_unprotected: int = 0
    workbook_unprotected: bool = False
    vba_unlocked: bool = False


def _read_varint(buf: bytes, i: int) -> tuple[int, int]:
    """Read a BRT record ID / record-length varint (1 or 2 bytes for IDs,
    up to 4 bytes for lengths). Returns (value, new_index)."""
    result = 0
    for shift in (0, 7, 14, 21):
        b = buf[i]; i += 1
        result |= (b & 0x7F) << shift
        if not (b & 0x80):
            return result, i
    raise ValueError("varint overflow")


def _write_varint(v: int) -> bytes:
    out = bytearray()
    while True:
        b = v & 0x7F
        v >>= 7
        if v:
            out.append(b | 0x80)
        else:
            out.append(b)
            return bytes(out)


def _zero_prot_records(buf: bytes) -> tuple[bytes, int, bool]:
    """Walk a BRT stream and zero body of any protection record encountered.
    Returns (new_buf, sheet_prot_count, book_prot_touched)."""
    out = bytearray()
    i, n = 0, len(buf)
    sheet_hits = 0
    book_hit = False
    while i < n:
        start = i
        rec_id, i = _read_varint(buf, i)
        rec_len, i = _read_varint(buf, i)
        body = buf[i:i + rec_len]
        i += rec_len

        if rec_id in PROT_RECS:
            body = bytes(rec_len)   # zero body, same length
            if rec_id in (BRT_BOOK_PROTECTION, BRT_BOOK_PROTECTION_ISO):
                book_hit = True
            else:
                sheet_hits += 1

        out += _write_varint(rec_id) + _write_varint(rec_len) + body
    return bytes(out), sheet_hits, book_hit


def _decrypt_to_bytes(src: Path, passwords: tuple[str, ...]):
    import msoffcrypto
    raw = src.read_bytes()
    with open(src, "rb") as f:
        try:
            of = msoffcrypto.OfficeFile(f)
            encrypted = of.is_encrypted()
        except Exception:
            encrypted = False
    if not encrypted:
        return raw, None, False
    last_err = None
    for pw in passwords:
        try:
            of = msoffcrypto.OfficeFile(io.BytesIO(raw))
            of.load_key(password=pw)
            buf = io.BytesIO(); of.decrypt(buf)
            return buf.getvalue(), pw, True
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Could not decrypt {src.name}: {last_err!r}")


def unlock_xlsb(
    src: Path, dst: Path, *, passwords=("",), also_unlock_vba: bool = True
) -> XlsbReport:
    src = Path(src); dst = Path(dst)
    report = XlsbReport(input_path=src, output_path=dst)
    zip_bytes, pw, was_enc = _decrypt_to_bytes(src, passwords)
    report.was_encrypted = was_enc
    report.open_password_used = pw

    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
        infos = {info.filename: info for info in zin.infolist()}

    for name, data in list(entries.items()):
        low = name.lower()
        if low == "xl/workbook.bin":
            new, sc, bh = _zero_prot_records(data)
            entries[name] = new
            if bh: report.workbook_unprotected = True
            report.sheets_unprotected += sc
        elif low.startswith("xl/worksheets/") and low.endswith(".bin"):
            new, sc, bh = _zero_prot_records(data)
            entries[name] = new
            report.sheets_unprotected += sc
        elif low == "xl/vbaproject.bin" and also_unlock_vba:
            new, changed = unlock_vba_bytes(data)
            if changed:
                entries[name] = new
                report.vba_unlocked = True

    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            info = infos[name]
            new_info = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
            new_info.compress_type = info.compress_type or zipfile.ZIP_DEFLATED
            new_info.external_attr = info.external_attr
            zout.writestr(new_info, data)
    return report
