"""VBA project password bypass.

Excel stores the VBA-project "lock for viewing" password in the plain-text
`PROJECT` stream (inside `_VBA_PROJECT_CUR` for .xls, or inside `vbaProject.bin`
for .xlsm/.xlsb). The password's DPB (Data Protection Block) entry is what
Excel checks. Renaming the key `DPB=` to `DPx=` (both 3 chars, same length)
invalidates the block: Excel then treats the project as unprotected. The
substitution is byte-length-invariant, so no other stream is affected.
"""
from __future__ import annotations
from pathlib import Path
from typing import Optional

import olefile


def _patch_project_stream_bytes(data: bytes) -> Optional[bytes]:
    if b'DPB="' not in data:
        return None
    patched = data.replace(b'DPB="', b'DPx="')
    if len(patched) != len(data):
        raise RuntimeError("DPB rename changed stream length — refusing.")
    return patched


def unlock_vba_project_ole(ole_file: Path) -> bool:
    """DPB-swap the PROJECT stream inside an OLE compound file. Returns True
    if a change was applied. No-op if no VBA project exists."""
    ole = olefile.OleFileIO(str(ole_file), write_mode=True)
    try:
        candidates = [
            ["_VBA_PROJECT_CUR", "PROJECT"],
            ["Macros", "PROJECT"],
            ["VBA", "PROJECT"],
        ]
        target = None
        for c in candidates:
            if ole.exists(c):
                target = c
                break
        if target is None:
            return False
        data = ole.openstream(target).read()
        patched = _patch_project_stream_bytes(data)
        if patched is None:
            return False
        ole.write_stream(target, patched)
        return True
    finally:
        ole.close()


def unlock_vba_project_bin(vba_project_bin_path: Path) -> bool:
    """DPB-swap the PROJECT stream inside a standalone vbaProject.bin (which
    is itself an OLE compound file — .xlsm/.xlsb store it as a zip entry)."""
    return unlock_vba_project_ole(vba_project_bin_path)


def unlock_vba_bytes(vba_project_bin_bytes: bytes) -> tuple[bytes, bool]:
    """DPB-swap a vbaProject.bin blob in memory. Returns (bytes, changed)."""
    import io as _io
    src = _io.BytesIO(vba_project_bin_bytes)
    ole = olefile.OleFileIO(src)
    try:
        for c in (["_VBA_PROJECT_CUR", "PROJECT"], ["Macros", "PROJECT"], ["VBA", "PROJECT"]):
            if ole.exists(c):
                data = ole.openstream(c).read()
                patched = _patch_project_stream_bytes(data)
                if patched is None:
                    return vba_project_bin_bytes, False
                # We can't write back to an OLE opened on BytesIO with write_mode
                # (olefile needs a file path for write). Use a temp file.
                import tempfile, shutil, os
                with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as tf:
                    tmp_path = tf.name
                    tf.write(vba_project_bin_bytes)
                ole.close()
                ok = unlock_vba_project_ole(Path(tmp_path))
                out = Path(tmp_path).read_bytes()
                os.unlink(tmp_path)
                return out, ok
        return vba_project_bin_bytes, False
    finally:
        try:
            ole.close()
        except Exception:
            pass
