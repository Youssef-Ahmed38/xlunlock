# xlunlock

Strip **workbook-open passwords**, **sheet/workbook protection**, and **VBA-project passwords** from Excel files, in place, without touching any formula, cell value, format, chart, macro body, or defined name.

Handles:

| Format          | Open password | Sheet/Workbook protection | VBA project lock | UI-forced hidden state |
|-----------------|:-:|:-:|:-:|:-:|
| `.xls` (BIFF8)  | RC4 / VelvetSweatshop | yes | yes | yes |
| `.xlsx`         | AES-128/256 | yes | n/a | n/a |
| `.xlsm`         | AES-128/256 | yes | yes | n/a |
| `.xlsb`         | AES-128/256 | yes | yes | n/a |

## Scope and ethics

This tool is meant for files **you own or are authorized to modify**. Typical use cases:

- You wrote a spreadsheet years ago and forgot the password.
- Your company hands you an internal template with protection that gets in the way of legitimate edits.
- You inherited maintenance of a workbook whose author is gone.

Not meant for: bypassing protection on files a third party wants kept locked.

The techniques used here (in-place record patching, MS-OVBA `DPB` substitution, and the well-known `VelvetSweatshop` default) are all publicly documented and have been in every open-source Office parser for years. `xlunlock` is a convenience wrapper, not a novel exploit.

## Install

```bash
pip install .
```

Runtime deps: `msoffcrypto-tool`, `olefile`.

## Usage

```bash
# Unlock a single workbook — writes "<name> (unlocked).<ext>" alongside it
xlunlock "PAM simulation.xls"

# Choose the output path
xlunlock -o clean.xlsx locked.xlsx

# Provide the open-password if it isn't the Excel default
xlunlock --password "myp4ss" locked.xls

# Preview what would change without touching disk
xlunlock --dry-run locked.xlsm

# Verbose byte-diff report
xlunlock -v locked.xls
```

Exit codes: `0` success, `1` unrecoverable (unknown password, corrupt file), `2` bad usage.

## What it changes, exactly

For each format, only the specific bytes/nodes below are touched. Everything else is preserved bit-for-bit.

### `.xls` (BIFF8, OLE compound file)

1. **File-open encryption** (`FILEPASS` record 0x002F): decrypt with `msoffcrypto` using the given password (default: try `VelvetSweatshop` first). The `FILEPASS` record is removed, encrypted body is rewritten as plaintext.
2. **Sheet protection** records — bodies zeroed in place:
   - `PROTECT` (0x0012)
   - `PASSWORD` (0x0013)
   - `SCENPROTECT` (0x00DD)
   - `WINDOWPROTECT` (0x0019)
   - `PROT4REV` (0x0086)
3. **`SheetProtection`** (Feat 0x0867): `isf` field set to 0 → Excel treats it as an unknown feature and ignores it.
4. **`WINDOW1`** (0x003D) `grbit` flags: force `fDspTabs`, `fDspHScroll`, `fDspVScroll` on; force `fHidden`, `fIconic` off. Nothing else touched.
5. **`WINDOW2`** (0x023E) per sheet: force `fDspGridSet`, `fDspRwCol`, `fDspZeros`, `fDspGuts` on if they were off.
6. **VBA project lock** (`_VBA_PROJECT_CUR/PROJECT` stream): `DPB=` → `DPx=` in place. Same length, so byte offsets are unchanged.

Byte offsets in the workbook stream are never shifted, so `BOUNDSHEET` pointers to each sheet's BOF stay valid. VBA streams are not touched. Every touched record keeps its original size.

### `.xlsx` / `.xlsm` (OOXML, zip container)

1. **Encryption**: if `EncryptedPackage` OLE is present, `msoffcrypto` decrypts and rewrites the file as a plain ZIP.
2. **`xl/workbook.xml`**: remove `<workbookProtection ...>` element.
3. **`xl/worksheets/sheet*.xml`**: remove `<sheetProtection ...>` element.
4. **`xl/vbaProject.bin`** (xlsm only): `DPB=` → `DPx=` in the `PROJECT` stream inside the OLE.

### `.xlsb` (OOXML binary, zip container)

1. **Encryption**: same as xlsx.
2. **`xl/workbook.bin`**: `BrtBookProtection` (0x0224) and `BrtBookProtectionIso` (0x0891) records — bodies zeroed.
3. **`xl/worksheets/sheet*.bin`**: `BrtSheetProtection` (0x0226) and `BrtSheetProtectionIso` (0x0867 in xlsb context) — bodies zeroed.
4. **`xl/vbaProject.bin`** (macros-enabled): DPB substitution.

## What it never changes

- Cell values.
- Formulas (`FORMULA`, `SHRFMLA`, `STRING`, `ARRAY` records; `<f>` elements).
- Cell formats (`XF`, `FONT`, `FORMAT`, `PALETTE`; `styles.xml`).
- Charts, images, pivots, defined names.
- VBA source code, VBA compiled p-code, VBA module names.
- Print settings, page setup, headers/footers.

## Verification

`xlunlock -v` emits a byte-diff report grouped by record type:

```
Total differing bytes: 66
Records touched: 23
  rec=0x0012 (PROTECT)         count=6  bytes=6
  rec=0x0013 (PASSWORD)        count=6  bytes=12
  rec=0x00DD (SCENPROTECT)     count=5  bytes=5
  rec=0x0019 (WINDOWPROTECT)   count=1  bytes=1
  rec=0x0867 (SheetProtection) count=6  bytes=42
Unexpected record types touched: 0
```

If `Unexpected record types touched` is anything but `0`, the tool aborts and refuses to write output — a safeguard against surprise edits.

## License

MIT.
