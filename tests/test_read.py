"""overview, read_range, find_cells, describe_columns on real workbooks."""

from __future__ import annotations

import csv
import datetime as dt
import io
import json
from pathlib import Path

import openpyxl
import pytest
from conftest import add_parts, build_excel_like

from sheet_mcp import server
from sheet_mcp.errors import SheetError


def test_overview_of_rich_workbook(rich: Path) -> None:
    out = server.workbook_overview("rich.xlsx")
    assert "'Sales' [visible]" in out and "'Other' [visible]" in out
    assert "table 'Sales1' A1:E8" in out
    assert "1 merged ranges" in out and "1 hidden rows (3)" in out
    assert "1 data validations" in out and "conditional-format" in out
    assert "Rate" in out and "Sales!$C$2" in out
    assert "WITHOUT cached values" in out
    assert "Safe to edit" in out
    assert "drops the cached results" in out


def test_overview_lists_what_would_be_lost(rich: Path) -> None:
    add_parts(rich, rich.with_name("c.xlsx"), {"xl/charts/chart1.xml": "<c/>", "xl/vbaProject.bin": b"x"})
    out = server.workbook_overview("c.xlsx")
    assert "would LOSE" in out and "charts" in out and "macros (VBA)" in out
    assert "allow_loss=true" in out


def test_overview_reports_error_cells_and_cached_formulas(excel_like: Path) -> None:
    out = server.workbook_overview("excel.xlsx")
    assert "#DIV/0!: F5" in out
    assert "WITHOUT cached" not in out
    assert "6 formulas" in out


def test_read_values_use_cached_results_and_turkish_text(excel_like: Path) -> None:
    out = server.read_range("excel.xlsx", "Veri", "A1:F6")
    assert "| 2 | Şeker | 2 | 3 | 6 | 2024-03-05 |" in out
    assert "İzmir çayı" in out
    assert "| 3 | İzmir çayı | 4 | 5 | 20 |" in out
    assert "#DIV/0!" in out
    assert "TRUE" in out  # boolean cell
    assert "dd/mm/yyyy" in out or "mm-dd-yy" in out  # builtin numFmt 14 is reported
    assert "xy" in out  # cached string result of a formula


def test_read_formulas_and_both(excel_like: Path) -> None:
    f = server.read_range("excel.xlsx", "Veri", "D2:D5", mode="formulas")
    assert "=B2*C2" in f and "=B3*C3" in f and "=B4*C4" in f  # shared formula expanded
    assert "=SUM(D2:D4)" in f
    b = server.read_range("excel.xlsx", "Veri", "D2:D5", mode="both")
    assert "=B3*C3 -> 20" in b and "=SUM(D2:D4) -> 27" in b


def test_no_cached_values_is_reported_not_blank(rich: Path) -> None:
    out = server.read_range("rich.xlsx", "Sales", "D2:D4")
    assert "(no cached value)" in out
    assert "never calculated" in out
    f = server.read_range("rich.xlsx", "Sales", "D2:D4", mode="formulas")
    assert "=B2*C2" in f and "never calculated" not in f


def test_1904_date_system(_sandbox: Path) -> None:
    build_excel_like(_sandbox / "a.xlsx", date1904=False)
    build_excel_like(_sandbox / "b.xlsx", date1904=True)
    a = server.read_range("a.xlsx", "Veri", "E2")
    b = server.read_range("b.xlsx", "Veri", "E2")
    assert "2024-03-05" in a
    assert "2028-03-06" in b  # same serial 45356, 1462 days later in the 1904 system
    assert "1904" in server.workbook_overview("b.xlsx")


def test_dates_show_iso_and_note_the_format(rich: Path) -> None:
    out = server.read_range("rich.xlsx", "Sales", "E2:E3")
    assert "2024-03-02" in out and "dd/mm/yyyy" in out
    assert "00:00" not in out  # midnight datetimes with a date format show as dates


def test_datetime_with_time_keeps_time(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = dt.datetime(2024, 5, 6, 13, 45, 10)
    ws["A1"].number_format = "yyyy-mm-dd hh:mm:ss"
    ws["A2"] = dt.time(8, 30)
    wb.save(_sandbox / "t.xlsx")
    out = server.read_range("t.xlsx", "Sheet", "A1:A2")
    assert "2024-05-06T13:45:10" in out and "08:30:00" in out


def test_number_formats_are_noted(rich: Path) -> None:
    out = server.read_range("rich.xlsx", "Sales", "C2:C3")
    assert '#,##0.00 "TL"' in out
    assert "21" in out  # unformatted value


def test_merged_cells_are_listed(rich: Path) -> None:
    out = server.read_range("rich.xlsx", "Sales", "G1:H2")
    assert "G1:H2 (value is in G1)" in out
    assert "| 1 | merged |  |" in out


def test_range_forms_and_clamping(rich: Path) -> None:
    assert "| 1 | Urun |" in server.read_range("rich.xlsx", "Sales", "A:B", max_cells=4)
    out = server.read_range("rich.xlsx", "Sales", "A1:ZZ500")
    assert "clamped to the used range: A1:H10" in out
    assert "| 2 |" in server.read_range("rich.xlsx", "Sales", "2:2")
    with pytest.raises(SheetError, match="cannot parse range"):
        server.read_range("rich.xlsx", "Sales", "not a range")
    with pytest.raises(SheetError, match="outside the used range"):
        server.read_range("rich.xlsx", "Sales", "A500:B600")


def test_truncation_tells_how_to_continue(rich: Path) -> None:
    out = server.read_range("rich.xlsx", "Sales", "A1:E10", max_cells=10)
    assert "truncated at 10 cells" in out and "Continue with range A3:E10" in out


def test_formats_csv_and_json(excel_like: Path) -> None:
    out = server.read_range("excel.xlsx", "Veri", "A1:D3", format="csv")
    rows = list(csv.reader(io.StringIO(out.split("\n\n#")[0])))
    assert rows[1][0] == "Şeker" and rows[1][3] == "6"
    data = json.loads(server.read_range("excel.xlsx", "Veri", "D2:E5", mode="both", format="json"))
    assert data["columns"] == ["D", "E"]
    assert data["rows"][0]["cells"][0] == {"formula": "=B2*C2", "value": 6}
    assert data["rows"][0]["cells"][1] == "2024-03-05"
    assert data["rows"][3]["cells"][0] == {"formula": "=SUM(D2:D4)", "value": 27}


def test_bad_arguments(rich: Path) -> None:
    with pytest.raises(SheetError, match="mode must be"):
        server.read_range("rich.xlsx", "Sales", mode="x")
    with pytest.raises(SheetError, match="format must be"):
        server.read_range("rich.xlsx", "Sales", format="xml")
    with pytest.raises(SheetError, match="no sheet named 'Nope'"):
        server.read_range("rich.xlsx", "Nope")
    with pytest.raises(SheetError, match="max_cells"):
        server.read_range("rich.xlsx", "Sales", max_cells=0)


def test_markdown_escapes_pipes_and_newlines(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "a|b\nc"
    wb.save(_sandbox / "p.xlsx")
    out = server.read_range("p.xlsx", "Sheet", "A1:A1")
    assert "a\\|b\\nc" in out


def test_find_text_regex_number(excel_like: Path) -> None:
    assert "Veri!A2" in server.find_cells("excel.xlsx", "seker")  # diacritic-insensitive
    assert "Veri!A3" in server.find_cells("excel.xlsx", "IZMIR")  # dotted/dotless i
    assert "Veri!A3" in server.find_cells("excel.xlsx", "İzmir")
    assert "No cells match" in server.find_cells("excel.xlsx", "izmir", case_sensitive=True)
    r = server.find_cells("excel.xlsx", r"^[ŞİÜ]", mode="regex", case_sensitive=True)
    assert "Veri!A2" in r and "Veri!A3" in r
    n = server.find_cells("excel.xlsx", "20", mode="number")
    assert "Veri!D3: 20" in n and "formula: =B3*C3" in n
    f = server.find_cells("excel.xlsx", "SUM(", search_in="formulas")
    assert "Veri!D5" in f and "1 match" in f
    assert "Veri!E2" not in server.find_cells("excel.xlsx", "2024", mode="number")


def test_find_regex_guard_and_errors(excel_like: Path) -> None:
    with pytest.raises(SheetError, match="nested quantifier"):
        server.find_cells("excel.xlsx", "(a+)+$", mode="regex")
    with pytest.raises(SheetError, match="invalid regular expression"):
        server.find_cells("excel.xlsx", "(", mode="regex")
    with pytest.raises(SheetError, match="not a number"):
        server.find_cells("excel.xlsx", "abc", mode="number")
    with pytest.raises(SheetError, match="query is empty"):
        server.find_cells("excel.xlsx", "")


def test_find_caps_results(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    for i in range(50):
        ws.append(["hit"])
    wb.save(_sandbox / "h.xlsx")
    out = server.find_cells("h.xlsx", "hit", max_results=5)
    assert "50 match(es)" in out and "45 more not shown" in out


def test_find_warns_when_values_are_uncached(rich: Path) -> None:
    out = server.find_cells("rich.xlsx", "999999")
    assert "no cached value" in out


def test_describe_columns(excel_like: Path, rich: Path) -> None:
    out = server.describe_columns("rich.xlsx", "Sales", "A1:E8")
    assert "Header: row 1 (auto-detected)" in out
    assert "B 'Adet': integer" in out and "min 2, max 8" in out
    assert "E 'Tarih': date" in out
    assert "D 'Toplam'" in out and "all 7 formulas" in out
    out2 = server.describe_columns("excel.xlsx", "Veri", "A1:F6")
    assert "Header: none detected" in out2  # row 1 mixes text and numbers
    assert "F: error" in out2 and "E: mixed" in out2
    out3 = server.describe_columns("excel.xlsx", "Veri", "D1:D5", header_row=0)
    assert "none (header_row=0)" in out3


def test_describe_columns_mixed_types_and_text_numbers(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["id", "amount", "note"])
    ws.append([1, 10, "a"])
    ws.append([2, "20", "b"])
    ws.append([3, 30.5, "a"])
    ws.append([4, "n/a", None])
    ws["B6"] = "=B2+1"
    ws["A6"] = 5
    wb.save(_sandbox / "m.xlsx")
    out = server.describe_columns("m.xlsx", "Sheet", None)
    assert "B 'amount': mixed" in out
    assert "types:" in out
    assert "text cell(s) look like numbers" in out
    assert "C 'note': text" in out and "top: 'a' x2" in out
    assert "MIXED formulas 1 / constants" in out


def test_describe_columns_header_row_validation(rich: Path) -> None:
    with pytest.raises(SheetError, match="header_row"):
        server.describe_columns("rich.xlsx", "Sales", "A1:E8", header_row=500)
    out = server.describe_columns("rich.xlsx", "Sales", "A1:E8", header_row=1)
    assert "row 1 (given)" in out and "Data starts at row 2" in out


def test_chartsheet_and_empty_sheet_errors(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    wb.save(_sandbox / "e.xlsx")
    with pytest.raises(SheetError, match="no data"):
        server.read_range("e.xlsx", "Sheet")
    assert "empty" in server.workbook_overview("e.xlsx")


def test_sandbox_blocks_traversal_and_absolute_paths(rich: Path, tmp_path: Path) -> None:
    with pytest.raises(SheetError, match="outside the allowed directory"):
        server.workbook_overview("../rich.xlsx")
    with pytest.raises(SheetError, match="outside the allowed directory"):
        server.workbook_overview(str(Path(tmp_path).parent / "x.xlsx"))
    with pytest.raises(SheetError, match="not found"):
        server.workbook_overview("missing.xlsx")
    (tmp_path / "a.txt").write_text("x")
    with pytest.raises(SheetError, match="only .xlsx and .xlsm"):
        server.workbook_overview("a.txt")
