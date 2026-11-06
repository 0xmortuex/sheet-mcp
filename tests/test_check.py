"""check_formulas on workbooks with known problems (and known-good ones)."""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from sheet_mcp import server
from sheet_mcp.formulas import r1c1, rename_in_formula


def _wb(path: Path, fill: dict[str, dict[str, object]]) -> None:
    wb = openpyxl.Workbook()
    first = True
    for name, cells in fill.items():
        ws = wb.active if first else wb.create_sheet(name)
        assert ws is not None
        ws.title = name
        first = False
        for addr, v in cells.items():
            ws[addr] = v
    wb.defined_names["Rate"] = DefinedName("Rate", attr_text="S!$Z$1")
    wb.save(path)


def test_each_problem_kind_is_found(_sandbox: Path) -> None:
    cells: dict[str, object] = {"Z1": 0.2}
    for r in range(2, 8):
        cells[f"A{r}"] = r
        cells[f"B{r}"] = f"=A{r}*2"
    cells["B5"] = "=A5*3"  # pattern outlier between identical neighbours
    cells["C1"] = "=#REF!+1"
    cells["C2"] = "=Nope!A1"
    cells["C3"] = "=NoSuchName*2"
    cells["C4"] = "=Rate*A2"  # defined: fine
    cells["C5"] = "=Missing[Col]"
    cells["D1"] = "=D2"
    cells["D2"] = "=D1"
    cells["D3"] = "=D3+1"
    cells["E1"] = "=SUM(E1:E3)"
    cells["F1"] = "=A2*1.18"
    cells["F2"] = "=ROUND(A2*2,2)+0+1"
    cells["F3"] = '="#REF! 3.14"'
    cells["G1"] = "=NOW()"
    cells["G2"] = "=TODAY()+OFFSET(A1,1,1)"
    cells["H2"] = "=A2*2"
    cells["H3"] = "=A3*2"
    cells["H4"] = 7  # hardcoded value between identical formulas
    cells["H5"] = "=A5*2"
    cells["H6"] = "=A6*2"
    _wb(_sandbox / "c.xlsx", {"S": cells})
    out = server.check_formulas("c.xlsx")
    assert "S!B5: =A5*3 differs from the formulas directly above and below" in out
    assert "S!C1: =#REF!+1 contains #REF!" in out
    assert "refers to missing sheet 'Nope'" in out
    assert "uses name 'NoSuchName'" in out
    assert "'Rate'" not in out
    assert "unknown table 'Missing'" in out
    assert "S!D1" in out and "S!D3" in out and "S!E1" in out  # circular cells
    assert out.count("cells that depend on each other") == 3  # D1<->D2, D3 itself, E1 in its own range
    assert "S!F1: =A2*1.18 has literal 1.18" in out
    assert "F2" not in out.split("Hardcoded numbers")[1].split("Volatile")[0]  # ROUND(..,2) and +0+1 are fine
    assert "F3" not in out  # strings are not formulas to analyse
    assert "S!G1: NOW()" in out and "OFFSET()" in out
    assert "S!H4: a typed number sits between formulas" in out


def test_clean_workbook_reports_no_problems(_sandbox: Path) -> None:
    cells: dict[str, object] = {}
    for r in range(1, 20):
        cells[f"A{r}"] = r
        cells[f"B{r}"] = f"=A{r}*2"
        cells[f"C{r}"] = f"=SUM($A$1:A{r})"
    cells["B20"] = "=SUM(B1:B19)"
    _wb(_sandbox / "ok.xlsx", {"S": cells})
    out = server.check_formulas("ok.xlsx")
    assert "No problems found." in out


def test_total_row_is_not_an_outlier(_sandbox: Path) -> None:
    cells: dict[str, object] = {f"A{r}": r for r in range(1, 6)}
    cells.update({f"B{r}": f"=A{r}+1" for r in range(1, 6)})
    cells["B6"] = "=SUM(B1:B5)"
    _wb(_sandbox / "t.xlsx", {"S": cells})
    assert "No problems found." in server.check_formulas("t.xlsx")


def test_cross_sheet_cycle_and_sheet_filter(_sandbox: Path) -> None:
    _wb(_sandbox / "x.xlsx", {"S": {"A1": "=T!A1+1"}, "T": {"A1": "=S!A1+1", "B1": "=1+1.5"}})
    out = server.check_formulas("x.xlsx")
    assert "S!A1, T!A1" in out
    assert "Checked 3 of 3 formulas on 2 sheet(s)" in out
    only = server.check_formulas("x.xlsx", sheet="T")
    assert "Checked 2 of 2 formulas on 1 sheet(s)" in only and "T!B1" in only


def test_cached_error_values_are_reported(excel_like: Path) -> None:
    out = server.check_formulas("excel.xlsx")
    assert "Veri!F5: shows #DIV/0!" in out


def test_max_issues_caps_output(_sandbox: Path) -> None:
    cells: dict[str, object] = {f"A{r}": f"=#REF!+{r}" for r in range(1, 40)}
    _wb(_sandbox / "m.xlsx", {"S": cells})
    out = server.check_formulas("m.xlsx", max_issues=5)
    assert "Broken references (#REF! or a missing sheet): 39" in out and "34 more" in out


def test_r1c1_patterns_are_position_independent() -> None:
    assert r1c1("=A2*2", 2, 2) == "=RC[-1]*2"
    assert r1c1("=A3*2", 3, 2) == "=RC[-1]*2"
    assert r1c1("=$A$1+B$2", 5, 3) == "=R1C1+R2C[-1]"
    assert r1c1("=SUM(Sheet2!A1:B2)", 1, 1) == "=SUM(Sheet2!RC:R[1]C[1])"
    assert r1c1("=SUM(A:A)", 1, 3) == "=SUM(C[-2]:C[-2])"


def test_rename_in_formula_unit() -> None:
    assert rename_in_formula("=A!B1+'A'!B2", "A", "Q Q") == "='Q Q'!B1+'Q Q'!B2"
    assert rename_in_formula('=IF(1,"A!B1",A!B1)', "A", "Data") == '=IF(1,"A!B1",Data!B1)'
    assert rename_in_formula("A!B1", "A", "Data") == "Data!B1"  # no leading '=' stays without one
    assert rename_in_formula("=SUM(B1:B2)", "A", "Data") == "=SUM(B1:B2)"
    assert rename_in_formula("=[1]A!B1", "A", "Data") == "=[1]A!B1"  # external workbook ref untouched
    assert rename_in_formula("=A!B1", "A", "C") == "='C'!B1"  # a sheet called C needs quotes (R1C1 ambiguity)
