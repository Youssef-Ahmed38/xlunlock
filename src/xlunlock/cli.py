"""Command-line interface: `xlunlock file.xls`."""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

from . import __version__
from .unlock import unlock, UnlockError


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="xlunlock",
        description="Strip open-passwords, sheet/workbook protection, and VBA "
                    "locks from Excel files (.xls, .xlsx, .xlsm, .xlsb) — "
                    "without changing formulas, formats, or macro bodies.",
    )
    p.add_argument("input", type=Path, nargs="+", help="One or more Excel files to unlock.")
    p.add_argument("-o", "--output", type=Path, default=None,
                   help="Output path (only valid when a single input is given). "
                        "Default: '<name> (unlocked).<ext>' next to the input.")
    p.add_argument("-p", "--password", default=None,
                   help="Open-password to try first. VelvetSweatshop and empty "
                        "string are always also attempted.")
    p.add_argument("--unlock-vba", action="store_true",
                   help="Also bypass the VBA-project 'lock for viewing' "
                        "password (DPB substitution in the PROJECT stream). "
                        "OFF by default because it causes Excel to show a "
                        "one-time 'invalid key DPx — Continue Loading Project?' "
                        "prompt. Only turn it on if you actually need to edit "
                        "the macro code — the workbook is editable either way.")
    p.add_argument("-v", "--verbose", action="store_true")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)

    if args.output and len(args.input) != 1:
        print("error: -o/--output only valid with a single input file", file=sys.stderr)
        return 2

    rc = 0
    for src in args.input:
        try:
            r = unlock(
                src,
                args.output,
                password=args.password,
                also_unlock_vba=args.unlock_vba,
            )
        except UnlockError as e:
            print(f"[FAIL] {src}: {e}", file=sys.stderr)
            rc = 1
            continue
        except Exception as e:
            print(f"[FAIL] {src}: {type(e).__name__}: {e}", file=sys.stderr)
            rc = 1
            continue

        # Concise one-line summary
        bits = []
        if r.was_encrypted:
            pw = r.open_password_used
            bits.append(f"decrypted (pw={'default' if pw in ('VelvetSweatshop','','',None) and pw != None else pw!r})")
        if r.sheets_unprotected:
            bits.append(f"{r.sheets_unprotected} sheets unprotected")
        if r.workbook_unprotected:
            bits.append("workbook unprotected")
        if r.vba_unlocked:
            bits.append("VBA project unlocked")
        summary = ", ".join(bits) or "no protection found"
        print(f"[OK]   {r.output_path}  ({r.format}: {summary})")

        if args.verbose:
            print(f"       input:  {r.input_path}")
            print(f"       detail: {r.detail}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
