"""Unlock a .xlsx / .xlsm workbook (OOXML zip)."""
from __future__ import annotations
import io
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from .vba import unlock_vba_bytes


# <sheetProtection ...  />  or  <sheetProtection ...  ></sheetProtection>
_SHEET_PROT = re.compile(
    rb"<(?:\w+:)?sheetProtection\b[^>]*(?:/>|>\s*</(?:\w+:)?sheetProtection>)",
    re.DOTALL,
)
_BOOK_PROT = re.compile(
    rb"<(?:\w+:)?workbookProtection\b[^>]*(?:/>|>\s*</(?:\w+:)?workbookProtection>)",
    re.DOTALL,
)


@dataclass
class XlsxReport:
    input_path: Path
    output_path: Path
    was_encrypted: bool = False
    open_password_used: Optional[str] = None
    sheets_unprotected: int = 0
    workbook_unprotected: bool = False
    vba_unlocked: bool = False


def _decrypt_to_bytes(src: Path, passwords: tuple[str, ...]) -> tuple[bytes, Optional[str], bool]:
    """Return (plain zip bytes, password used, was_encrypted)."""
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

    last_err: Optional[Exception] = None
    for pw in passwords:
        try:
            of = msoffcrypto.OfficeFile(io.BytesIO(raw))
            of.load_key(password=pw)
            buf = io.BytesIO()
            of.decrypt(buf)
            return buf.getvalue(), pw, True
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(
        f"Could not decrypt {src.name}. Tried passwords: {passwords!r}. Last error: {last_err!r}"
    )


def unlock_xlsx(
    src: Path,
    dst: Path,
    *,
    passwords: tuple[str, ...] = ("",),
    also_unlock_vba: bool = True,
) -> XlsxReport:
    src = Path(src); dst = Path(dst)
    report = XlsxReport(input_path=src, output_path=dst)

    zip_bytes, pw, was_enc = _decrypt_to_bytes(src, passwords)
    report.was_encrypted = was_enc
    report.open_password_used = pw

    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
        infos = {info.filename: info for info in zin.infolist()}

    for name, data in list(entries.items()):
        low = name.lower()
        if low == "xl/workbook.xml":
            new, n = _BOOK_PROT.subn(b"", data)
            if n:
                entries[name] = new
                report.workbook_unprotected = True
        elif low.startswith("xl/worksheets/") and low.endswith(".xml"):
            new, n = _SHEET_PROT.subn(b"", data)
            if n:
                entries[name] = new
                report.sheets_unprotected += n
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
