"""A 200k-row sheet: reads stay capped and fast enough, limits are enforced."""

from __future__ import annotations

import time
from pathlib import Path

import openpyxl
import pytest

from sheet_mcp import package, server

ROWS = 200_000


@pytest.fixture(scope="module")
def big(tmp_path_factory: pytest.TempPathFactory) -> Path:
    d = tmp_path_factory.mktemp("big")
    p = d / "big.xlsx"
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    ws.append(["id", "name", "amount", "calc"])
    for i in range(1, ROWS):
        ws.append([i, f"name {i % 100}", i * 1.5, f"=C{i + 1}*2"])
    wb.save(p)
    return p


@pytest.fixture
def _bigbox(big: Path) -> Path:
    from sheet_mcp import sandbox

    sandbox.configure(big.parent)
    package._cache.clear()
    return big


def test_overview_and_capped_reads_on_200k_rows(_bigbox: Path) -> None:
    t = time.monotonic()
    ov = server.workbook_overview("big.xlsx")
    assert f"{ROWS} rows x 4 cols" in ov and f"{ROWS - 1} formulas" in ov
    out = server.read_range("big.xlsx", "Data", "A199990:D200000", max_cells=100)
    assert "| 199999 | 199998 |" in out or "| 199999 |" in out
    head = server.read_range("big.xlsx", "Data", None, max_cells=40)
    assert "truncated at 40 cells" in head and "Continue with range A11:D200000" in head
    assert time.monotonic() - t < 120


def test_describe_and_find_and_check_on_200k_rows(_bigbox: Path) -> None:
    d = server.describe_columns("big.xlsx", "Data")
    assert "A 'id': integer" in d and f"{ROWS - 1} data row" in d
    assert "D 'calc'" in d and "all" in d
    f = server.find_cells("big.xlsx", "name 7", max_results=3)
    assert "match(es)" in f and "more not shown" in f
    c = server.check_formulas("big.xlsx", max_issues=3)
    assert "Checked 100000 of 199999 formulas" in c
    assert "Stopped analysing after 100000 formulas" in c


def test_edit_limit_for_huge_sheets(_bigbox: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from sheet_mcp.errors import SheetError

    monkeypatch.setattr(package, "MAX_EDIT_SHEET_BYTES", 1000)
    with pytest.raises(SheetError, match="limited to"):
        server.set_cells("big.xlsx", "Data", [{"cell": "A1", "value": 1}])
