"""Shared fixtures: real workbooks written by openpyxl and hand-zipped XML."""

from __future__ import annotations

import datetime as dt
import shutil
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
import pytest
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from sheet_mcp import package, sandbox


@pytest.fixture(autouse=True)
def _sandbox(tmp_path: Path) -> Any:
    sandbox.configure(tmp_path)
    package._cache.clear()
    yield tmp_path
    sandbox.configure(None)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def write_zip(path: Path, files: dict[str, bytes | str]) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return path


def add_parts(src: Path, dst: Path, extra: dict[str, bytes | str], rename: dict[str, bytes | str] | None = None) -> Path:
    """Copy an xlsx, adding (or replacing) parts. Used to fake features openpyxl never writes."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        replaced = {**(rename or {})}
        for item in zin.infolist():
            if item.filename in replaced:
                zout.writestr(item.filename, replaced.pop(item.filename))
            elif item.filename in extra:
                continue
            else:
                zout.writestr(item, zin.read(item.filename))
        for name, data in {**replaced, **extra}.items():
            zout.writestr(name, data)
    return dst


def build_rich(path: Path) -> Path:
    """A workbook exercising everything the editor must preserve."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sales"
    ws.append(["Urun", "Adet", "Fiyat", "Toplam", "Tarih"])
    for i in range(2, 9):
        ws.append([f"Item {i}", i, 10.5 * i, f"=B{i}*C{i}", dt.date(2024, 3, i)])
        ws[f"E{i}"].number_format = "dd/mm/yyyy"
        ws[f"C{i}"].number_format = '#,##0.00 "TL"'
    ws["A1"].font = Font(bold=True, color="FF0000")
    ws["A1"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["B2"].border = Border(left=Side(style="thin"))
    ws["B2"].alignment = Alignment(horizontal="center")
    ws["B8"].alignment = Alignment(horizontal="right")
    ws["B8"].font = Font(italic=True)
    ws["A10"] = "Total"
    ws["D10"] = "=SUM(D2:D8)"
    ws.merge_cells("G1:H2")
    ws["G1"] = "merged"
    ws.column_dimensions["A"].width = 31
    ws.column_dimensions["D"].width = 17
    ws.row_dimensions[3].hidden = True
    dv = DataValidation(type="list", formula1='"a,b,c"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("F2:F9")
    ws.conditional_formatting.add(
        "B2:B9", CellIsRule(operator="greaterThan", formula=["5"], fill=PatternFill("solid", bgColor="FF0000"))
    )
    ws["A2"].comment = Comment("check this", "tester")
    ws.add_table(Table(displayName="Sales1", ref="A1:E8"))
    ws2 = wb.create_sheet("Other")
    ws2["A1"] = "=Sales!D2*2"
    ws2["A2"] = 'Sales total: Sales!D10 and "Sales"'  # text that merely mentions the sheet
    ws2["A3"] = "=Sales!$D$10+SUM(Sales!B2:B8)"
    wb.defined_names["Rate"] = DefinedName("Rate", attr_text="Sales!$C$2")
    wb.save(path)
    return path


@pytest.fixture
def rich(_sandbox: Path) -> Path:
    return build_rich(_sandbox / "rich.xlsx")


# ---------------------------------------------------------------- hand-built "Excel-like" file

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""
ROOT_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
WB_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
STYLES = """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
</styleSheet>"""
SHARED = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
<si><t>Ürün</t></si><si><t>Şeker</t></si><si><t>İzmir çayı</t></si><si><t>Date</t></si></sst>"""
SHEET1 = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:F6"/>
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1" t="s"><v>3</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="D2"><f t="shared" ref="D2:D4" si="0">B2*C2</f><v>6</v></c><c r="E2" s="1"><v>45356</v></c></row>
<row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3"><v>4</v></c><c r="C3"><v>5</v></c><c r="D3"><f t="shared" si="0"/><v>20</v></c></row>
<row r="4"><c r="B4"><v>1</v></c><c r="C4"><v>1</v></c><c r="D4"><f t="shared" si="0"/><v>1</v></c></row>
<row r="5"><c r="D5"><f>SUM(D2:D4)</f><v>27</v></c><c r="E5" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="F5" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>
<row r="6"><c r="A6" t="b"><v>1</v></c></row>
</sheetData></worksheet>"""


def build_excel_like(path: Path, date1904: bool = False) -> Path:
    attr = ' date1904="1"' if date1904 else ""
    wbxml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<workbookPr{attr}/>"
        '<sheets><sheet name="Veri" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    return write_zip(
        path,
        {
            "[Content_Types].xml": CT,
            "_rels/.rels": ROOT_RELS,
            "xl/workbook.xml": wbxml,
            "xl/_rels/workbook.xml.rels": WB_RELS,
            "xl/styles.xml": STYLES,
            "xl/sharedStrings.xml": SHARED,
            "xl/worksheets/sheet1.xml": SHEET1,
        },
    )


@pytest.fixture
def excel_like(_sandbox: Path) -> Path:
    return build_excel_like(_sandbox / "excel.xlsx")


def copy_to(src: Path, dst: Path) -> Path:
    shutil.copy(src, dst)
    return dst
