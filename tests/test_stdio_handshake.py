"""MCP stdio smoke test: spawns the real server over stdio, lists the tools, calls them,
and checks that instructive errors survive the trip to the agent (MCPServer hides any
exception that is not a ToolError)."""

import sys
from pathlib import Path

import pytest

pytest.importorskip("mcp")

from mcp import ClientSession  # noqa: E402
from mcp.client.stdio import StdioServerParameters, stdio_client  # noqa: E402

EXPECTED_TOOLS = {
    "workbook_overview",
    "read_range",
    "find_cells",
    "describe_columns",
    "check_formulas",
    "set_cells",
    "append_rows",
    "add_sheet",
    "rename_sheet",
    "csv_to_xlsx",
    "xlsx_to_csv",
}


def _params(root: Path) -> StdioServerParameters:
    return StdioServerParameters(command=sys.executable, args=["-m", "sheet_mcp", "--dir", str(root)])


@pytest.mark.anyio
async def test_tools_and_a_session_over_stdio(rich: Path) -> None:
    async with stdio_client(_params(rich.parent)) as (read, write), ClientSession(read, write) as session:
        init = await session.initialize()
        assert init.server_info.name == "sheet"
        names = {t.name for t in (await session.list_tools()).tools}
        assert names == EXPECTED_TOOLS, (names - EXPECTED_TOOLS, EXPECTED_TOOLS - names)

        res = await session.call_tool("workbook_overview", {"path": "rich.xlsx"})
        assert not res.is_error and "'Sales' [visible]" in res.content[0].text  # type: ignore[union-attr]

        res = await session.call_tool("read_range", {"path": "rich.xlsx", "sheet": "Sales", "range": "A1:B3"})
        assert "| 1 | Urun | Adet |" in res.content[0].text  # type: ignore[union-attr]

        dry = await session.call_tool(
            "set_cells", {"path": "rich.xlsx", "sheet": "Sales", "changes": [{"cell": "B2", "value": 41}]}
        )
        assert "DRY RUN" in dry.content[0].text  # type: ignore[union-attr]
        assert not (rich.parent / "rich.edited.xlsx").exists()
        done = await session.call_tool(
            "set_cells",
            {"path": "rich.xlsx", "sheet": "Sales", "changes": [{"cell": "B2", "value": 41}], "dry_run": False},
        )
        assert not done.is_error and "Saved" in done.content[0].text  # type: ignore[union-attr]
        res = await session.call_tool("read_range", {"path": "rich.edited.xlsx", "sheet": "Sales", "range": "B2"})
        assert "| 2 | 41 |" in res.content[0].text  # type: ignore[union-attr]


@pytest.mark.anyio
async def test_instructive_errors_reach_the_agent(rich: Path) -> None:
    async with stdio_client(_params(rich.parent)) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        res = await session.call_tool("workbook_overview", {"path": "../etc/passwd.xlsx"})
        assert res.is_error and "outside the allowed directory" in res.content[0].text  # type: ignore[union-attr]
        res = await session.call_tool("read_range", {"path": "rich.xlsx", "sheet": "Nope"})
        assert res.is_error and "Sheets in this workbook" in res.content[0].text  # type: ignore[union-attr]
        res = await session.call_tool(
            "set_cells",
            {"path": "rich.xlsx", "sheet": "Sales", "changes": [{"cell": "H2", "value": 1}], "dry_run": False},
        )
        assert res.is_error and "merged range G1:H2" in res.content[0].text  # type: ignore[union-attr]


@pytest.mark.anyio
async def test_refusal_to_lose_features_over_stdio(rich: Path) -> None:
    from conftest import add_parts

    add_parts(rich, rich.with_name("chart.xlsx"), {"xl/charts/chart1.xml": "<c/>"})
    async with stdio_client(_params(rich.parent)) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        res = await session.call_tool(
            "set_cells",
            {"path": "chart.xlsx", "sheet": "Sales", "changes": [{"cell": "B2", "value": 1}], "dry_run": False},
        )
        assert res.is_error and "Refusing to edit" in res.content[0].text and "allow_loss=true" in res.content[0].text  # type: ignore[union-attr]
