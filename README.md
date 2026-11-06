# sheet-mcp

**Let your AI agent read and edit Excel files without wrecking them.**

<!-- mcp-name: io.github.0xmortuex/sheet-mcp -->

sheet-mcp is an [MCP](https://modelcontextprotocol.io) server for `.xlsx` workbooks. Its edge over
the other Excel MCP servers (for example
[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server),
[mort-lab/excel-mcp](https://github.com/mort-lab/excel-mcp),
[bassem-elsodany/mcp_excel_server](https://github.com/bassem-elsodany/mcp_excel_server)) is
**what it refuses to do**:

- It checks the file's zip structure **before** editing and lists everything a save would destroy:
  charts, images, shapes, pivot tables, macros, slicers, external links, threaded comments,
  dynamic-array formulas, sparklines, x14 conditional formats and more. If any are present, every
  edit is **refused** unless you pass `allow_loss=true`.
- Edits are **dry runs by default** (a before/after table), and they write a **new file**
  (`report.edited.xlsx`). Nothing is overwritten unless you say `overwrite=true`, and the write is
  atomic (temp file, then replace).
- Formulas stay formulas. Untouched cells, number formats, styles, column widths, data validation,
  conditional formatting, defined names, tables, comments and merged cells come out of an edit
  identical (the test suite proves it cell by cell).
- It never pretends to calculate. Values are the results Excel cached in the file; if the file was
  never calculated (written by a script) the tools say so instead of showing blanks.

## Install

Claude Code:

```bash
claude mcp add sheet -- uvx sheet-mcp --dir "/path/to/your/spreadsheets"
```

Claude Desktop (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "sheet": {
      "command": "uvx",
      "args": ["sheet-mcp", "--dir", "C:\\Users\\you\\Documents\\spreadsheets"]
    }
  }
}
```

`--dir` is the only place the server may read or write (default: the directory it was started in).
Paths you give the tools are relative to it. Windows paths work.

## Tools

| Tool | What it does |
|---|---|
| `workbook_overview` | Sheets (hidden / very hidden), used ranges, tables, named ranges, merged cells, hidden rows/columns, validations, conditional formats, comments, formula counts, formulas with no cached value, error cells, date system, **and what a save would lose** |
| `read_range` | Cells as markdown / csv / json. `mode`: `values`, `formulas` or `both`. Dates as ISO 8601 with the number format noted, merged ranges listed, capped with "continue with range X" |
| `find_cells` | Text, regex or number search across sheets, in values, formulas or both. Turkish-aware case folding (`istanbul` finds `İstanbul`) |
| `describe_columns` | Header detection, inferred type per column, nulls, distinct, min/max, mixed types, numbers stored as text, formulas vs constants |
| `check_formulas` | Broken refs, unknown names, error cells, circular refs, R1C1 pattern outliers, typed numbers inside a column of formulas, hardcoded numbers in formulas, volatile functions |
| `set_cells` | Write cells (value or formula) with explicit coercion rules. Dry run first. Refuses merged interiors and array formulas |
| `append_rows` | Append after the last data row or to an Excel table, copying styles and number formats from the row above |
| `add_sheet` / `rename_sheet` | Validated names. Rename rewrites every real reference (formulas on all sheets, names, validations, conditional formats, hyperlinks) and never touches text inside strings |
| `csv_to_xlsx` / `xlsx_to_csv` | Explicit `encoding` (utf-8, utf-8-sig, cp1254, ...) and `delimiter` (comma, semicolon, tab, pipe). A wrong encoding is an error, not garbage |

### How `set_cells` interprets values

| You send | Stored as |
|---|---|
| `"=SUM(A1:A3)"` | a formula |
| `"'=not a formula"` | the text `=not a formula` (one leading apostrophe is removed) |
| `"'007"` | the text `007` |
| `42`, `3.5` | numbers; `true` / `false` are booleans; `null` clears the cell |
| `"2024-12-31"`, `"2024-12-31T10:30"` | real dates (the cell keeps its date format, or gets `yyyy-mm-dd`) |
| any other string | text. Strings are never silently turned into numbers |

Also checked: unbalanced parentheses, formulas pointing at sheets that do not exist, control
characters, cells beyond XFD1048576. Newer functions that Excel stores as `_xlfn.NAME` (XLOOKUP,
IFS, TEXTJOIN ...) and dynamic-array functions (FILTER, SORT, UNIQUE ...) produce a warning,
because openpyxl cannot write them the way Excel expects.

## What a save does and does not preserve

openpyxl (the library underneath) cannot calculate and cannot round-trip everything:

- **Always** (every save): cached formula results are dropped. The formulas are kept, and Excel
  recalculates when the file is opened. Until then, other readers (pandas, previewers) see empty
  formula cells. The tool tells you this after each save.
- **Refused unless `allow_loss=true`**: charts, images, shapes and text boxes, pivot tables, macros
  (VBA), slicers and timelines, external links, threaded comments, form controls / ActiveX,
  embedded objects, data connections / Power Query, the data model, dynamic-array metadata,
  custom ribbon UI, sparklines, x14 conditional formatting and data validation.
- **Harmless, reported**: the calculation chain, printer settings, and comment box sizes (text
  and authors are kept).

After writing, the server re-opens the new file and checks that every sheet has exactly the number of
formulas it should; if not, nothing is written to the destination.

## Try asking

- "Give me an overview of `budget.xlsx`. Is it safe to edit?"
- "Show me `Sheet1!A1:F30` with formulas and values side by side."
- "Which cells in `forecast.xlsx` use VLOOKUP? Do any formulas look inconsistent with their neighbours?"
- "Describe the columns of the Customers sheet. Is any column a mix of formulas and typed numbers?"
- "Set `C7` to `=SUM(C2:C6)` and `D7` to the date 2026-01-31. Show me the dry run first."
- "Append these three rows to the Orders table."
- "Rename the sheet `Q1` to `Q1 2026` and tell me which formulas change."
- "Convert `export.csv` (Windows-1254, semicolon separated) to xlsx."

## Security

- **Sandbox**: every path is resolved and must stay inside `--dir`; `..` and absolute paths
  elsewhere are refused (symlinks are resolved first).
- **Zip bombs**: at most 10,000 entries, 512 MiB per entry, 1 GiB in total, a 500:1 compression
  ratio limit, entry sizes verified while reading (a header that lies is refused), duplicate or
  path-traversing entry names refused, files over 256 MiB refused. Editing additionally needs the
  sheet XML to be at most 100 MiB (it is loaded into memory); reading has no such limit.
- **XML attacks**: any XML part containing a `DOCTYPE` or `ENTITY` declaration is refused before
  parsing (Excel never writes one), non-UTF-8 XML is refused, and parsing goes through
  `defusedxml`. A billion-laughs workbook is part of the test suite.
- **Regex search**: patterns over 200 characters or with nested quantifiers like `(a+)+` are
  refused; Python's `re` has no timeout, so this is a guard rather than a guarantee.
- **CSV**: fields starting with `=` are imported as text, never as formulas. Exports are written
  as-is, so treat a CSV exported from untrusted data as untrusted when opening it in Excel.
- **Writes**: never in place by default, never over an existing file without `overwrite=true`,
  atomic replace, and temp files are cleaned up on failure.
- Macros are never executed (nothing here runs workbook code).

## Limits

- Formulas are **not calculated**. There is no recalculation engine; read cached results or open
  the file in Excel / LibreOffice.
- `.xls` (old binary) and password-protected workbooks are refused with an explanation.
- `.xlsm` can be edited only with `allow_loss=true` (the VBA project bytes are carried over; buttons
  and ribbon wiring may not be).
- Circular-reference detection follows direct cell and range references only (not `INDIRECT`,
  `OFFSET`, names or table references). Column-pattern checks flag only a cell that differs from
  the formulas directly above and below it.
- CSV/TSV are handled through the two conversion tools; the read tools take `.xlsx`/`.xlsm`.

## Development

```bash
uv sync --extra test
uv run pytest tests
uv run --with ruff==0.16.0 ruff check src tests
uv run --with mypy==2.3.0 mypy --strict src
```

Tests run against real workbooks: files written by openpyxl, hand-zipped XML for features openpyxl
never writes (chart parts, pivot caches, `vbaProject.bin`, external links ...), a billion-laughs and a
real compression-ratio zip bomb, 1900 and 1904 date systems, Turkish text, a 200,000-row sheet, and a
test that launches the installed server over real MCP stdio.

There is a from-scratch tutorial that builds a minimal version of this server step by step:
[sheet-mcp-the-hard-way](https://github.com/0xmortuex/sheet-mcp-the-hard-way).

MIT licensed.
