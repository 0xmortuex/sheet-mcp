"""sheet-mcp: read and safely edit .xlsx files."""

from __future__ import annotations

import argparse
import functools
import os
from collections.abc import Callable
from typing import Any, TypeVar

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from . import checks, csvio, editor, package, sandbox, views
from .errors import SheetError
from .package import PackageInfo

mcp = MCPServer(
    "sheet",
    instructions=(
        "Read and safely edit Excel .xlsx/.xlsm workbooks. Start with workbook_overview: it lists "
        "sheets, tables, names, hidden rows, formula counts, error cells and - importantly - any "
        "features (charts, pivots, macros, images, slicers, external links ...) that saving would "
        "destroy. Read with read_range (mode values/formulas/both), find_cells, describe_columns, "
        "check_formulas. Edit with set_cells, append_rows, add_sheet, rename_sheet: every edit is a "
        "DRY RUN by default (shows before/after), writes to a NEW file '<name>.edited.xlsx' unless "
        "output/overwrite are given, and is refused if it would lose features unless "
        "allow_loss=true. Formulas are never calculated here: values come from the cached results "
        "stored in the file, and tools say so when there are none. For CSV/TSV use csv_to_xlsx and "
        "xlsx_to_csv."
    ),
)

F = TypeVar("F", bound=Callable[..., Any])


def tool(fn: F) -> F:
    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except SheetError as e:
            raise ToolError(str(e)) from e
        except (OSError, ValueError, KeyError) as e:
            msg = e.args[0] if isinstance(e, KeyError) and e.args else str(e)
            raise ToolError(f"{type(e).__name__}: {msg}") from e

    mcp.tool()(wrapper)
    return fn


def _open(path: str) -> PackageInfo:
    return package.inspect(sandbox.resolve_workbook(path))


@tool
def workbook_overview(path: str) -> str:
    """Start here. Summarise a workbook without changing it: sheets (with hidden/very hidden
    state), used ranges, tables, named ranges, merged cells, hidden rows/columns, data
    validations, conditional formats, comments, number of formulas, formulas that have NO cached
    value (the file was never calculated), cells showing errors like #REF!, and the list of
    features that saving would LOSE (charts, images, pivot tables, macros, slicers, external
    links, ...). The file is checked for zip bombs and XML entity attacks first.

    path: .xlsx or .xlsm, relative to the server's --dir (or absolute inside it).
    """
    return views.overview(_open(path))


@tool
def read_range(
    path: str,
    sheet: str,
    range: str | None = None,
    mode: str = "values",
    format: str = "markdown",
    max_cells: int = 2000,
) -> str:
    """Read cells from one sheet.

    range: A1-style: 'A1:D20', 'B3', 'A:C' (whole columns), '2:9' (whole rows). Omit for the whole
      used range. Anything beyond the used range is clamped.
    mode: 'values' (cached results of formulas; constants as stored), 'formulas' (formula text
      instead of results), or 'both' ('=SUM(A1:A3) -> 6').
    format: 'markdown' (default, with row numbers and column letters), 'csv', or 'json'.
    max_cells: output cap (default 2000, max 20000); when exceeded the reply says which range to
      request next.

    Dates are shown as ISO 8601 and the number format is reported in a note. Formulas are NOT
    calculated: if the file has no cached values the reply says so instead of showing blanks.
    Merged ranges in view are listed (the value lives in the top-left cell).
    """
    return views.read_range(_open(path), sheet, range, mode, format, max_cells)


@tool
def find_cells(
    path: str,
    query: str,
    mode: str = "text",
    sheet: str | None = None,
    search_in: str = "both",
    case_sensitive: bool = False,
    max_results: int = 100,
) -> str:
    """Find cells across all sheets (or one) and return their addresses.

    mode: 'text' (substring), 'regex' (Python regular expression), or 'number' (numeric cells equal
      to the query, e.g. '1234.5').
    search_in: 'values' (cached results / constants), 'formulas' (formula text, e.g. find every
      cell that uses VLOOKUP), or 'both' (default).
    Case-insensitive search also folds Turkish letters (dotted/dotless i, and u o s c g with their
    Turkish marks), so 'istanbul' finds 'İstanbul' and 'urun' finds 'Ürün'. Each hit shows the sheet, address, value and formula.
    """
    return views.find_cells(_open(path), query, mode, sheet, search_in, case_sensitive, max_results)


@tool
def describe_columns(
    path: str,
    sheet: str,
    range: str | None = None,
    header_row: int | None = None,
    max_rows: int = 500000,
) -> str:
    """Profile the columns of a sheet or range like a data table: detected header row, inferred
    type per column (integer, number, text, date, datetime, boolean, mixed ...), non-null and empty
    counts, distinct count, min/max, most common text values, numbers stored as text, error cells,
    and whether the column holds formulas, constants or a MIX of both (a sign someone overwrote a
    formula with a typed number).

    header_row: omit to auto-detect; 0 means 'no header'; N means row N is the header (rows above
    it are ignored). range limits the area (default: the whole used range).
    """
    return views.describe_columns(_open(path), sheet, range, header_row, max_rows)


@tool
def check_formulas(path: str, sheet: str | None = None, max_issues: int = 30) -> str:
    """Audit formulas structurally (nothing is calculated). Reports: broken references (#REF! or a
    missing sheet), unknown names (#NAME?), cells currently showing error values, circular
    references reachable through direct cell/range references, formulas that differ from the
    pattern (R1C1) of the formulas directly above AND below them in the same column, typed numbers
    sitting inside a column of formulas, hardcoded numbers inside arithmetic (e.g. =A2*1.18), and
    volatile functions (NOW, TODAY, RAND, OFFSET, INDIRECT ...). Results are grouped per kind and
    capped by max_issues. Use sheet= to check one sheet.
    """
    return checks.check(_open(path), sheet, max_issues)


@tool
def set_cells(
    path: str,
    sheet: str,
    changes: list[dict[str, Any]],
    dry_run: bool = True,
    output: str | None = None,
    overwrite: bool = False,
    allow_loss: bool = False,
) -> str:
    """Write single cells, changing nothing else. DRY RUN BY DEFAULT: shows before/after for every
    cell and writes nothing; repeat with dry_run=false to save.

    changes: list of {"cell": "B7", "value": ...} or {"cell": "B7", "formula": "SUM(B2:B6)"}, with
    an optional per-change "sheet". Value coercion rules:
      - a string starting with '=' is a FORMULA ("=SUM(A1:A3)")
      - a string starting with a single quote is literal TEXT with the quote removed
        ("'=not a formula" stores =not a formula; "'007" stores the text 007)
      - JSON numbers are numbers, true/false are booleans, null clears the cell
      - ISO strings 'YYYY-MM-DD' and 'YYYY-MM-DDTHH:MM[:SS]' become real dates (the cell keeps its
        date format, or gets yyyy-mm-dd); use a leading quote to keep such a string as text
      - any other string is text (never auto-converted to a number)
    Existing styles, number formats, validations, conditional formats and column widths are kept.
    Writes into merged-cell interiors and array formulas are refused. Formulas referring to a
    sheet that does not exist are refused.

    The result goes to a NEW file '<name>.edited.xlsx' next to the original unless output= is given;
    an existing file is never replaced unless overwrite=true (use overwrite=true with output equal
    to path to edit in place). If the workbook has features openpyxl would drop (see
    workbook_overview) the edit is refused unless allow_loss=true. All formulas lose their cached
    results on save (Excel recalculates when opened).
    """
    return editor.set_cells(_open(path), sheet, changes, dry_run, output, overwrite, allow_loss)


@tool
def append_rows(
    path: str,
    sheet: str,
    rows: list[list[Any]],
    table: str | None = None,
    dry_run: bool = True,
    output: str | None = None,
    overwrite: bool = False,
    allow_loss: bool = False,
) -> str:
    """Append rows after the last data row (or to the bottom of an Excel table). DRY RUN BY DEFAULT.

    rows: list of rows, each a list of values using the same coercion rules as set_cells ('=...'
    formula, leading quote = literal text, ISO dates, null = empty cell).
    table: name of an Excel table on the sheet to extend (its range grows with the new rows). If the
    last data row belongs to a table and table= is omitted, the call is refused so rows are not left
    outside the table. Tables with a totals row are refused.
    Number formats and styles are copied from the row above. Rows start in the first used column.
    Refused if the space below is not empty. Warns when data validation or conditional formatting
    ranges end at the old last row and will not cover the new rows.
    output/overwrite/allow_loss work as in set_cells.
    """
    return editor.append_rows(_open(path), sheet, rows, table, dry_run, output, overwrite, allow_loss)


@tool
def add_sheet(
    path: str,
    name: str,
    position: int | None = None,
    dry_run: bool = True,
    output: str | None = None,
    overwrite: bool = False,
    allow_loss: bool = False,
) -> str:
    """Add an empty sheet. name is validated (max 31 characters, none of [ ] : * ? / \\, unique
    ignoring case). position is 1-based (default: at the end). DRY RUN BY DEFAULT; output/overwrite/
    allow_loss work as in set_cells."""
    return editor.add_sheet(_open(path), name, position, dry_run, output, overwrite, allow_loss)


@tool
def rename_sheet(
    path: str,
    sheet: str,
    new_name: str,
    dry_run: bool = True,
    output: str | None = None,
    overwrite: bool = False,
    allow_loss: bool = False,
) -> str:
    """Rename a sheet AND update every reference to it: formulas on all sheets (only real
    references, never text inside string literals), defined names, data-validation and
    conditional-format formulas, and internal hyperlinks. DRY RUN BY DEFAULT: lists each formula
    that would change. new_name is validated like add_sheet. output/overwrite/allow_loss work as in
    set_cells."""
    return editor.rename_sheet(_open(path), sheet, new_name, dry_run, output, overwrite, allow_loss)


@tool
def csv_to_xlsx(
    csv_path: str,
    output: str,
    encoding: str = "utf-8",
    delimiter: str = ",",
    sheet_name: str = "Sheet1",
    infer_types: bool = True,
    overwrite: bool = False,
) -> str:
    """Convert a CSV/TSV file into a new .xlsx.

    encoding: any Python codec: utf-8, utf-8-sig (Excel's BOM), cp1254 (Turkish), cp1252, latin-1,
      utf-16 ... A wrong encoding is an error, never silently garbled.
    delimiter: one character or one of comma, semicolon, tab, pipe (Turkish/German Excel exports
      often use semicolons).
    infer_types: plain numbers (no leading zeros, at most 15 digits) and ISO dates become real
      numbers/dates; everything else stays text. Fields starting with '=' always stay text (a CSV is
      data, not formulas). Set false to store every field as text.
    Refuses to replace an existing output unless overwrite=true.
    """
    return csvio.csv_to_xlsx(csv_path, output, encoding, delimiter, sheet_name, infer_types, overwrite)


@tool
def xlsx_to_csv(
    path: str,
    sheet: str,
    output: str,
    range: str | None = None,
    encoding: str = "utf-8",
    delimiter: str = ",",
    mode: str = "values",
    overwrite: bool = False,
) -> str:
    """Export one sheet (or a range) to CSV/TSV (.csv, .tsv or .txt).

    mode: 'values' writes cached results (formula cells with no cached value are written empty and
      counted in a warning); 'formulas' writes the formula text. Dates are ISO 8601.
    encoding/delimiter as in csv_to_xlsx; characters the encoding cannot represent are an error.
    Refuses to replace an existing output unless overwrite=true.
    """
    return csvio.xlsx_to_csv(_open(path), sheet, output, range, encoding, delimiter, mode, overwrite)


def main() -> None:
    parser = argparse.ArgumentParser(prog="sheet-mcp", description="MCP server for safe Excel edits")
    parser.add_argument(
        "--dir",
        default=os.environ.get("SHEET_MCP_DIR"),
        help="directory the tools may read and write (default: current directory; env SHEET_MCP_DIR)",
    )
    args = parser.parse_args()
    sandbox.configure(args.dir)
    mcp.run()


if __name__ == "__main__":
    main()
