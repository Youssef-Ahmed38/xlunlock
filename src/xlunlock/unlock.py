"""Top-level unlock dispatch by file extension."""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Union


class UnlockError(RuntimeError):
    pass


@dataclass
class UnlockResult:
    format: str                     # "xls", "xlsx", "xlsm", "xlsb"
    input_path: Path
    output_path: Path
    was_encrypted: bool
    open_password_used: Optional[str]
    sheets_unprotected: int
    workbook_unprotected: bool
    vba_unlocked: bool
    detail: object                  # the format-specific report


DEFAULT_PASSWORDS = ("VelvetSweatshop", "")


def _detect(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in (".xls", ".xlt"):
        return "xls"
    if ext == ".xlsx":
        return "xlsx"
    if ext == ".xlsm":
        return "xlsm"
    if ext == ".xlsb":
        return "xlsb"
    # Fallback: sniff magic bytes.
    head = path.read_bytes()[:8]
    if head[:4] == b"\xD0\xCF\x11\xE0":
        return "xls"
    if head[:2] == b"PK":
        return "xlsx"
    raise UnlockError(f"Unrecognized file type: {path.name}")


def _default_output(src: Path) -> Path:
    return src.with_name(f"{src.stem} (unlocked){src.suffix}")


def unlock(
    src: Union[str, Path],
    dst: Union[str, Path, None] = None,
    *,
    password: Optional[str] = None,
    also_unlock_vba: bool = False,
) -> UnlockResult:
    src = Path(src)
    if not src.exists():
        raise UnlockError(f"No such file: {src}")
    dst = Path(dst) if dst else _default_output(src)

    fmt = _detect(src)
    passwords = tuple([password]) + DEFAULT_PASSWORDS if password else DEFAULT_PASSWORDS

    if fmt == "xls":
        from .xls import unlock_xls
        from . import _biff as biff
        r = unlock_xls(src, dst, passwords=passwords, also_unlock_vba=also_unlock_vba)
        # In BIFF8, PROTECT (0x12) appears once per sheet-with-protection PLUS
        # once at the workbook level. WINDOWPROTECT (0x19) is the workbook-window
        # lock. Report them separately.
        total_protect = r.touched.get(biff.PROTECT, 0)
        # A conservative split: assume 1 workbook-level PROTECT if we saw WINDOWPROTECT
        wb_had_prot = biff.WINDOWPROTECT in r.touched or biff.PROTECT in r.touched
        sheet_count = max(0, total_protect - (1 if wb_had_prot else 0))
        return UnlockResult(
            "xls", src, dst, r.was_encrypted, r.open_password_used,
            sheet_count, wb_had_prot, r.vba_unlocked, r,
        )
    elif fmt in ("xlsx", "xlsm"):
        from .xlsx import unlock_xlsx
        r = unlock_xlsx(src, dst, passwords=passwords, also_unlock_vba=also_unlock_vba)
        return UnlockResult(
            fmt, src, dst, r.was_encrypted, r.open_password_used,
            r.sheets_unprotected, r.workbook_unprotected, r.vba_unlocked, r,
        )
    elif fmt == "xlsb":
        from .xlsb import unlock_xlsb
        r = unlock_xlsb(src, dst, passwords=passwords, also_unlock_vba=also_unlock_vba)
        return UnlockResult(
            fmt, src, dst, r.was_encrypted, r.open_password_used,
            r.sheets_unprotected, r.workbook_unprotected, r.vba_unlocked, r,
        )
    else:
        raise UnlockError(f"Unsupported format: {fmt}")
