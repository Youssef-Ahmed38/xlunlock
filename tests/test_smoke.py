"""Smoke tests: build a tiny protected .xlsx on the fly and verify unlock."""
from __future__ import annotations
import zipfile
import io
from pathlib import Path

import pytest

from xlunlock.unlock import unlock


MINIMAL_WORKBOOK_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <workbookProtection workbookPassword="ABCD" lockStructure="1"/>
  <sheets><sheet name="S1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
</workbook>
"""

MINIMAL_SHEET_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetProtection password="ABCD" sheet="1" objects="1" scenarios="1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row></sheetData>
</worksheet>
"""

MINIMAL_CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""

MINIMAL_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

MINIMAL_XL_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""


def _build_protected_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", MINIMAL_CONTENT_TYPES)
        z.writestr("_rels/.rels", MINIMAL_RELS)
        z.writestr("xl/_rels/workbook.xml.rels", MINIMAL_XL_RELS)
        z.writestr("xl/workbook.xml", MINIMAL_WORKBOOK_XML)
        z.writestr("xl/worksheets/sheet1.xml", MINIMAL_SHEET_XML)


def test_xlsx_strip_protection(tmp_path: Path) -> None:
    src = tmp_path / "locked.xlsx"
    dst = tmp_path / "clean.xlsx"
    _build_protected_xlsx(src)

    r = unlock(src, dst)

    assert dst.exists(), "unlock() must write the output"
    assert r.sheets_unprotected == 1
    assert r.workbook_unprotected is True

    with zipfile.ZipFile(dst) as z:
        book = z.read("xl/workbook.xml")
        sheet = z.read("xl/worksheets/sheet1.xml")

    assert b"workbookProtection" not in book
    assert b"sheetProtection" not in sheet
    # Cell data must be preserved
    assert b"hello" in sheet


def test_default_output_path(tmp_path: Path) -> None:
    src = tmp_path / "MyBook.xlsx"
    _build_protected_xlsx(src)
    r = unlock(src)
    assert r.output_path == tmp_path / "MyBook (unlocked).xlsx"
    assert r.output_path.exists()


def test_unknown_extension_raises(tmp_path: Path) -> None:
    from xlunlock.unlock import UnlockError
    p = tmp_path / "not-excel.txt"
    p.write_text("hi")
    with pytest.raises(UnlockError):
        unlock(p)
