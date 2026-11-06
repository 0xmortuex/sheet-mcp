"""Differential test: the streaming reader must agree with openpyxl cell by cell."""

from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
import pytest
from conftest import build_excel_like, write_zip

from sheet_mcp import package, reader


def mine(path: Path, sheet: str) -> dict[str, Any]:
    package._cache.clear()
    info = package.inspect(path)
    sc = info.worksheet(sheet)
    out = {}
    for row in reader.rows(info, sheet, (1, 1, sc.max_row, sc.max_col)):
        for c in row:
            out[c.coord] = c
    return out


def theirs(path: Path, sheet: str) -> dict[str, tuple[Any, Any]]:
    wf = openpyxl.load_workbook(path, read_only=True)
    wv = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for rf, rv in zip(wf[sheet].iter_rows(), wv[sheet].iter_rows(), strict=False):
        for cf, cv in zip(rf, rv, strict=False):
            if getattr(cf, "coordinate", None):
                out[cf.coordinate] = (cf.value, cv.value)
    wf.close()
    wv.close()
    return out


def compare(path: Path, sheet: str) -> int:
    a, b = mine(path, sheet), theirs(path, sheet)
    n = 0
    for coord, (f, v) in b.items():
        c = a[coord]
        if f is not None and (isinstance(f, str) and f.startswith("=") or hasattr(f, "text")):
            assert c.formula is not None, coord
            if isinstance(f, str):
                assert c.formula == f, coord
            else:
                assert c.formula.startswith("{" + (f.text if f.text.startswith("=") else "=" + f.text)), coord
            assert c.value == v, (coord, c.value, v)
        else:
            assert c.formula is None, coord
            assert c.value == f, (coord, c.value, f)
        n += 1
    return n


def test_rich_and_excel_like(rich: Path, excel_like: Path, _sandbox: Path) -> None:
    assert compare(rich, "Sales") > 40
    assert compare(rich, "Other") >= 3
    assert compare(excel_like, "Veri") >= 20
    build_excel_like(_sandbox / "m.xlsx", date1904=True)
    assert compare(_sandbox / "m.xlsx", "Veri") >= 20


def test_kitchen_sink(_sandbox: Path) -> None:
    from openpyxl.worksheet.formula import ArrayFormula

    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "K"
    ws["A1"] = 1
    ws["A2"] = -2.5
    ws["A3"] = 1e-9
    ws["A4"] = 123456789012345
    ws["B1"] = dt.datetime(2024, 2, 29, 23, 59, 59)
    ws["B2"] = dt.date(1999, 12, 31)
    ws["B3"] = dt.time(7, 5, 3)
    ws["B4"] = dt.timedelta(hours=30, minutes=15)
    ws["B4"].number_format = "[h]:mm:ss"
    ws["C1"] = True
    ws["C2"] = False
    ws["C3"] = "#N/A"
    ws["C4"] = "Çiğ köfte ıİ şŞ ğĞ ü"
    ws["D1"] = "line1\nline2"
    ws["D2"] = "  padded  "
    ws["D3"] = "=A1+A2"
    ws["D4"] = '="a"&"b"'
    ws["E1"] = ArrayFormula("E1:E2", "=A1:A2*2")
    ws["F5"] = "far away"
    wb.save(_sandbox / "k.xlsx")
    assert compare(_sandbox / "k.xlsx", "K") >= 15


def test_inline_and_rich_shared_strings_and_missing_row_numbers(_sandbox: Path) -> None:
    base = build_excel_like(_sandbox / "b.xlsx")
    sheet = zipfile.ZipFile(base).read("xl/worksheets/sheet1.xml").decode()
    sheet = sheet.replace(
        '<row r="6"><c r="A6" t="b"><v>1</v></c></row>',
        '<row r="6"><c r="A6" t="b"><v>1</v></c><c r="B6" t="inlineStr"><is><r><t>rich </t></r><r><t>inline</t></r></is></c></row>'
        '<row r="8"><c r="C8" t="str"><f>"x"</f><v></v></c><c r="D8"><f>1+1</f></c></row>',
    ).replace('<dimension ref="A1:F6"/>', '<dimension ref="A1:F8"/>')
    sst = (
        '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<si><t>Ürün</t></si><si><r><t>Şe</t></r><r><t>ker</t></r><rPh sb=\"0\" eb=\"1\"><t>PHONETIC</t></rPh></si>"
        "<si><t>İzmir çayı</t></si><si><t>Date</t></si></sst>"
    )
    out = _sandbox / "i.xlsx"
    with zipfile.ZipFile(base) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    files["xl/worksheets/sheet1.xml"] = sheet.encode()
    files["xl/sharedStrings.xml"] = sst.encode()
    write_zip(out, files)
    m = mine(out, "Veri")
    assert m["A2"].value == "Şeker"  # rich runs joined, phonetic text ignored
    assert m["B6"].value == "rich inline"
    assert m["C8"].formula == '="x"' and m["C8"].no_cache is False  # cached empty string
    assert m["D8"].formula == "=1+1" and m["D8"].no_cache is True
    t = theirs(out, "Veri")
    assert t["A2"][0] == "Şeker"


def test_first_rows_and_windows_match(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    for i in range(1, 60):
        if i % 7:
            ws.append([i, f"r{i}", None, i * 2])
        else:
            ws.append([])
    wb.save(_sandbox / "w.xlsx")
    info = package.inspect(_sandbox / "w.xlsx")
    full = reader.read_grid(info, "Sheet", (1, 1, info.sheets[0].max_row, 4))
    window = reader.read_grid(info, "Sheet", (10, 2, 20, 3))
    assert [[c.value for c in r] for r in window] == [[c.value for c in r[1:3]] for r in full[9:20]]
    assert [c.coord for c in window[0]] == ["B10", "C10"]
    assert len(window) == 11  # blank rows are padded, not skipped


def test_cell_without_address_is_refused(_sandbox: Path) -> None:
    base = build_excel_like(_sandbox / "b.xlsx")
    with zipfile.ZipFile(base) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    files["xl/worksheets/sheet1.xml"] = files["xl/worksheets/sheet1.xml"].replace(b'<c r="A6" t="b">', b'<c t="b">')
    out = write_zip(_sandbox / "n.xlsx", files)
    info = package.inspect(out)
    with pytest.raises(Exception, match="no address"):
        reader.read_grid(info, "Veri", (1, 1, 6, 6))
