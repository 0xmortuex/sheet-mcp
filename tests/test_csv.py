"""csv_to_xlsx / xlsx_to_csv with explicit encodings and delimiters."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from sheet_mcp import server
from sheet_mcp.errors import SheetError

TURKISH = "Ad;Şehir;Tutar;Tarih;Kod\nÇağrı;İstanbul;1234.5;2024-03-05;007\nÖzlem;Iğdır;10;2024-03-06;=cmd\n"


def test_csv_to_xlsx_semicolon_cp1254(_sandbox: Path) -> None:
    (_sandbox / "t.csv").write_bytes(TURKISH.encode("cp1254"))
    out = server.csv_to_xlsx("t.csv", "t.xlsx", encoding="cp1254", delimiter="semicolon")
    assert "3 rows x 5 columns" in out
    ws = openpyxl.load_workbook(_sandbox / "t.xlsx")["Sheet1"]
    assert [c.value for c in ws[2]] == ["Çağrı", "İstanbul", 1234.5, dt.datetime(2024, 3, 5), "007"]
    assert ws["E3"].value == "=cmd" and ws["E3"].data_type == "s"  # never a formula
    assert ws["D2"].number_format == "yyyy-mm-dd"
    assert ws["C3"].value == 10 and isinstance(ws["C3"].value, int)


def test_wrong_encoding_is_an_error_not_mojibake(_sandbox: Path) -> None:
    (_sandbox / "t.csv").write_bytes(TURKISH.encode("cp1254"))
    with pytest.raises(SheetError, match="not valid utf-8"):
        server.csv_to_xlsx("t.csv", "t.xlsx", delimiter=";")
    with pytest.raises(SheetError, match="unknown encoding"):
        server.csv_to_xlsx("t.csv", "t.xlsx", encoding="klingon")
    with pytest.raises(SheetError, match="delimiter must be"):
        server.csv_to_xlsx("t.csv", "t.xlsx", delimiter="ab")
    assert not (_sandbox / "t.xlsx").exists()


def test_tsv_utf8_sig_and_infer_off(_sandbox: Path) -> None:
    (_sandbox / "a.tsv").write_bytes("﻿a\tb\n1\t2.50\n".encode())
    server.csv_to_xlsx("a.tsv", "a.xlsx", encoding="utf-8-sig", delimiter="tab", sheet_name="Veri", infer_types=False)
    ws = openpyxl.load_workbook(_sandbox / "a.xlsx")["Veri"]
    assert ws["A1"].value == "a" and ws["A2"].value == "1" and ws["B2"].value == "2.50"


def test_overwrite_rules(_sandbox: Path) -> None:
    (_sandbox / "t.csv").write_text("a,b\n1,2\n", encoding="utf-8")
    server.csv_to_xlsx("t.csv", "t.xlsx")
    with pytest.raises(SheetError, match="already exists"):
        server.csv_to_xlsx("t.csv", "t.xlsx")
    server.csv_to_xlsx("t.csv", "t.xlsx", overwrite=True)
    with pytest.raises(SheetError, match="must end in .xlsx"):
        server.csv_to_xlsx("t.csv", "t.csv2")
    with pytest.raises(SheetError, match="outside the allowed directory"):
        server.csv_to_xlsx("t.csv", "../x.xlsx")


def test_giant_field_and_ragged_rows(_sandbox: Path) -> None:
    (_sandbox / "g.csv").write_text("a,b\n" + "x" * 200000 + ",1\n", encoding="utf-8")
    with pytest.raises(SheetError, match="not parseable"):
        server.csv_to_xlsx("g.csv", "g.xlsx")
    (_sandbox / "r.csv").write_text("a,b,c\n1\n1,2,3,4\n", encoding="utf-8")
    assert "3 rows x 4 columns" in server.csv_to_xlsx("r.csv", "r.xlsx")


def test_xlsx_to_csv_roundtrip(rich: Path, _sandbox: Path, excel_like: Path) -> None:
    out = server.xlsx_to_csv("excel.xlsx", "Veri", "o.csv", encoding="utf-8-sig", delimiter=";")
    assert "6 rows" in out and "Dates are ISO" in out
    text = (_sandbox / "o.csv").read_bytes().decode("utf-8-sig")
    lines = text.splitlines()
    assert lines[1].startswith("Şeker;2;3;6;2024-03-05")
    assert lines[2].startswith("İzmir çayı;4;5;20")
    assert lines[4].split(";")[3:6] == ["27", "xy", "#DIV/0!"]
    server.xlsx_to_csv("excel.xlsx", "Veri", "f.tsv", range="D2:D5", delimiter="tab", mode="formulas")
    assert (_sandbox / "f.tsv").read_text(encoding="utf-8").splitlines() == ["=B2*C2", "=B3*C3", "=B4*C4", "=SUM(D2:D4)"]


def test_xlsx_to_csv_warns_about_uncached_formulas(rich: Path, _sandbox: Path) -> None:
    out = server.xlsx_to_csv("rich.xlsx", "Sales", "o.csv", range="A1:D3")
    assert "WARNING" in out and "no cached value" in out
    assert (_sandbox / "o.csv").read_text(encoding="utf-8").splitlines()[1].startswith("Item 2,2,21,")


def test_xlsx_to_csv_encoding_errors_and_clobber(excel_like: Path, _sandbox: Path) -> None:
    with pytest.raises(SheetError, match="cannot encode"):
        server.xlsx_to_csv("excel.xlsx", "Veri", "o.csv", encoding="ascii")
    assert not (_sandbox / "o.csv").exists()
    (_sandbox / "o.csv").write_text("keep", encoding="utf-8")
    with pytest.raises(SheetError, match="already exists"):
        server.xlsx_to_csv("excel.xlsx", "Veri", "o.csv")
    assert (_sandbox / "o.csv").read_text(encoding="utf-8") == "keep"
    with pytest.raises(SheetError, match="must end in"):
        server.xlsx_to_csv("excel.xlsx", "Veri", "o.xlsx")
    assert sorted(p.name for p in _sandbox.iterdir()) == ["excel.xlsx", "o.csv"]
