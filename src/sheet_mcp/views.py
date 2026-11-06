"""The read tools: overview, read_range, find_cells, describe_columns."""

from __future__ import annotations

import csv
import datetime as dt
import io
import json
import math
import re
from collections import Counter
from contextlib import closing
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from . import reader
from .cells import CellView, display, fold, jsonable
from .errors import SheetError
from .package import PackageInfo, SheetScan, describe_loss, short

MAX_COLS = 200
DEFAULT_MAX_CELLS = 2000
HARD_MAX_CELLS = 20000
NO_CACHE_NOTE = (
    "{n} formula cell(s) have no cached value: this file was never calculated by Excel or "
    "LibreOffice (typical for files written by scripts, or saved by this server). sheet-mcp "
    "cannot calculate formulas. Open and save the file in Excel/LibreOffice to get values; "
    "mode='formulas' shows the formulas themselves."
)


def _ranges(pairs: list[tuple[int, int]], letters: bool = False, cap: int = 8) -> str:
    def f(n: int) -> str:
        return get_column_letter(n) if letters else str(n)

    parts = [f(a) if a == b else f"{f(a)}-{f(b)}" for a, b in pairs[:cap]]
    if len(pairs) > cap:
        parts.append(f"... +{len(pairs) - cap} more")
    return ", ".join(parts)


def _hidden_count(pairs: list[tuple[int, int]]) -> int:
    return sum(b - a + 1 for a, b in pairs)


def overview(info: PackageInfo) -> str:
    lines = [
        f"{info.path}",
        f"  {info.file_size:,} bytes, {info.n_entries} zip entries, "
        f"date system {'1904 (Mac)' if info.date1904 else '1900'}, "
        f"{info.total_formulas} formula(s) in total",
        f"Sheets ({len(info.sheets)}):",
    ]
    for s in info.sheets:
        head = f"  {s.index + 1}. {s.name!r} [{s.state}]"
        if s.kind == "chartsheet":
            lines.append(head + " chart sheet (no cells)")
            continue
        if s.max_row == 0:
            lines.append(head + " empty")
            continue
        lines.append(
            head + f" used range A1:{get_column_letter(s.max_col)}{s.max_row} "
            f"({s.max_row} rows x {s.max_col} cols, {s.n_cells} filled cells)"
        )
        bits = []
        if s.n_formulas:
            bits.append(f"{s.n_formulas} formulas")
        if s.n_array_formulas:
            bits.append(f"{s.n_array_formulas} array formulas")
        if s.n_formulas_no_cache:
            bits.append(f"{s.n_formulas_no_cache} WITHOUT cached values")
        if s.merges:
            bits.append(f"{len(s.merges)} merged ranges")
        if s.hidden_rows:
            bits.append(f"{_hidden_count(s.hidden_rows)} hidden rows ({_ranges(s.hidden_rows)})")
        if s.hidden_cols:
            bits.append(f"hidden columns {_ranges(s.hidden_cols, letters=True)}")
        if s.n_validations:
            bits.append(f"{s.n_validations} data validations")
        if s.n_cond_formats:
            bits.append(f"{s.n_cond_formats} conditional-format blocks")
        if s.n_comments:
            bits.append(f"{s.n_comments} comments")
        if bits:
            lines.append("       " + "; ".join(bits))
        for t in s.tables:
            extra = []
            if t.totals_rows:
                extra.append("has totals row")
            lines.append(f"       table {t.name!r} {t.ref}" + (f" ({', '.join(extra)})" if extra else ""))
        if s.error_total:
            shown = ", ".join(
                f"{code}: {', '.join(addrs[:5])}{' ...' if len(addrs) > 5 else ''}"
                for code, addrs in s.errors.items()
            )
            lines.append(f"       cells showing errors ({s.error_total}): {shown}")
    visible_names = [n for n in info.names if not n.name.startswith("_xlnm.")]
    if visible_names:
        lines.append(f"Named ranges ({len(visible_names)}):")
        for n in visible_names[:30]:
            scope = f" (sheet {n.scope!r})" if n.scope else ""
            lines.append(f"  {n.name}{scope} = {short(n.text, 80)}{' [hidden]' if n.hidden else ''}")
        if len(visible_names) > 30:
            lines.append(f"  ... +{len(visible_names) - 30} more")
    no_cache = sum(s.n_formulas_no_cache for s in info.sheets)
    if no_cache:
        lines.append("Cached values: " + NO_CACHE_NOTE.format(n=no_cache))
    lines.append("")
    if info.lossy:
        lines.append("WARNING - editing and saving would LOSE (edit tools refuse unless allow_loss=true):")
        for item in describe_loss(info):
            lines.append(f"  - {item}")
    else:
        lines.append("Safe to edit: no charts, images, pivots, macros, slicers or external links found.")
    if info.minor:
        lines.append("Minor, harmless losses on any save: " + "; ".join(info.minor))
    if info.total_formulas:
        lines.append(
            "Note: any save drops the cached results of ALL formulas (openpyxl cannot calculate); "
            "Excel recalculates them when the file is opened."
        )
    return "\n".join(lines)


# --------------------------------------------------------------------- read_range


def _cell_text(c: CellView, mode: str) -> str:
    if c.formula is not None:
        if mode == "formulas":
            return c.formula
        if mode == "values":
            if c.no_cache:
                return "(no cached value)"
            return display(c.value, c.fmt)
        shown = "(no cached value)" if c.no_cache else display(c.value, c.fmt)
        return f"{c.formula} -> {shown}"
    txt = display(c.value, c.fmt)
    if isinstance(c.value, str) and txt[:1] in ("=", "'"):
        return "'" + txt  # keep literal text distinguishable from formulas
    return txt


def _md_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace("|", "\\|").replace("\r", "").replace("\n", "\\n")


def _merge_notes(scan: SheetScan, b: reader.Bounds) -> list[str]:
    out = []
    for ref in scan.merges:
        try:
            c0, r0, c1, r1 = range_boundaries(ref)
        except ValueError:
            continue
        if r0 <= b[2] and r1 >= b[0] and c0 <= b[3] and c1 >= b[1]:
            out.append(f"{ref} (value is in {get_column_letter(c0)}{r0})")
    return out


def _format_notes(grid: list[list[CellView]], b: reader.Bounds) -> list[str]:
    date_fmts: dict[int, set[str]] = {}
    num_fmts: dict[int, set[str]] = {}
    for row in grid:
        for k, c in enumerate(row):
            if c.fmt == "General":
                continue
            col = b[1] + k
            if isinstance(c.value, (dt.datetime, dt.date, dt.time, dt.timedelta)):
                date_fmts.setdefault(col, set()).add(c.fmt)
            elif isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                num_fmts.setdefault(col, set()).add(c.fmt)

    def fmt(d: dict[int, set[str]]) -> str:
        items = [f"{get_column_letter(c)}: {' | '.join(sorted(f))}" for c, f in sorted(d.items())]
        return "; ".join(items[:12]) + (" ..." if len(items) > 12 else "")

    out = []
    if date_fmts:
        out.append("Dates are shown as ISO 8601; the cell number formats are - " + fmt(date_fmts))
    if num_fmts:
        out.append("Numbers are shown unformatted; non-default number formats - " + fmt(num_fmts))
    return out


def read_range(
    info: PackageInfo,
    sheet: str,
    range_spec: str | None,
    mode: str,
    fmt: str,
    max_cells: int,
) -> str:
    if mode not in ("values", "formulas", "both"):
        raise SheetError("mode must be 'values', 'formulas' or 'both'")
    if fmt not in ("markdown", "csv", "json"):
        raise SheetError("format must be 'markdown', 'csv' or 'json'")
    if not 1 <= max_cells <= HARD_MAX_CELLS:
        raise SheetError(f"max_cells must be between 1 and {HARD_MAX_CELLS}")
    scan = info.worksheet(sheet)
    bounds, clamp_note = reader.parse_range(range_spec, scan)
    notes: list[str] = []
    if clamp_note:
        notes.append(clamp_note)
    min_row, min_col, max_row, max_col = bounds
    if max_col - min_col + 1 > MAX_COLS:
        max_col = min_col + MAX_COLS - 1
        notes.append(f"only the first {MAX_COLS} columns are shown; narrow the range")
    width = max_col - min_col + 1
    rows_allowed = max(1, max_cells // width)
    if max_row - min_row + 1 > rows_allowed:
        last = min_row + rows_allowed - 1
        notes.append(
            f"truncated at {max_cells} cells: showing rows {min_row}-{last} of {min_row}-{max_row}. "
            f"Continue with range {get_column_letter(min_col)}{last + 1}:"
            f"{get_column_letter(max_col)}{max_row}"
        )
        max_row = last
    shown: reader.Bounds = (min_row, min_col, max_row, max_col)
    grid = reader.read_grid(info, scan.name, shown)
    notes.extend(_format_notes(grid, shown))
    merges = _merge_notes(scan, shown)
    if merges:
        notes.append("merged ranges in view: " + "; ".join(merges[:20]))
    no_cache = sum(1 for r in grid for c in r if c.no_cache)
    if no_cache and mode != "formulas":
        notes.append(NO_CACHE_NOTE.format(n=no_cache))
    errs = [c.coord for r in grid for c in r if c.is_error]
    if errs:
        notes.append(f"error values at: {', '.join(errs[:10])}{' ...' if len(errs) > 10 else ''}")
    header = f"{scan.name}!{reader.a1(shown)} ({mode})"
    letters = [get_column_letter(c) for c in range(min_col, max_col + 1)]
    if fmt == "json":
        rows_json: list[dict[str, Any]] = []
        for k, row in enumerate(grid):
            cells: list[Any] = []
            for c in row:
                if c.formula is not None and mode == "both":
                    cells.append({"formula": c.formula, "value": None if c.no_cache else jsonable(c.value, c.fmt)})
                elif c.formula is not None and mode == "formulas":
                    cells.append(c.formula)
                elif c.formula is not None:
                    cells.append(None if c.no_cache else jsonable(c.value, c.fmt))
                else:
                    cells.append(jsonable(c.value, c.fmt))
            rows_json.append({"row": min_row + k, "cells": cells})
        return json.dumps(
            {"sheet": scan.name, "range": reader.a1(shown), "mode": mode, "columns": letters,
             "rows": rows_json, "notes": notes},
            ensure_ascii=False,
        )
    if fmt == "csv":
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\n")
        for row in grid:
            w.writerow([_cell_text(c, mode) for c in row])
        body = buf.getvalue().rstrip("\n")
        return body + ("\n\n# " + "\n# ".join(notes) if notes else "")
    lines = [header, "", "| row | " + " | ".join(letters) + " |", "|---|" + "---|" * len(letters)]
    for k, row in enumerate(grid):
        lines.append(f"| {min_row + k} | " + " | ".join(_md_escape(_cell_text(c, mode)) for c in row) + " |")
    if notes:
        lines.append("")
        lines.extend(f"- {n}" for n in notes)
    return "\n".join(lines)


# --------------------------------------------------------------------- find_cells

_NESTED_QUANT = re.compile(r"\((?:[^()\\]|\\.)*[+*](?:[^()\\]|\\.)*\)[+*{]")


def find_cells(
    info: PackageInfo,
    query: str,
    mode: str,
    sheet: str | None,
    search_in: str,
    case_sensitive: bool,
    max_results: int,
) -> str:
    if mode not in ("text", "regex", "number"):
        raise SheetError("mode must be 'text', 'regex' or 'number'")
    if search_in not in ("both", "values", "formulas"):
        raise SheetError("search_in must be 'both', 'values' or 'formulas'")
    if not 1 <= max_results <= 1000:
        raise SheetError("max_results must be between 1 and 1000")
    if not query:
        raise SheetError("query is empty")
    pattern: re.Pattern[str] | None = None
    number: float | None = None
    if mode == "regex":
        if len(query) > 200:
            raise SheetError("regex is longer than 200 characters")
        if _NESTED_QUANT.search(query):
            raise SheetError(
                "regex has a nested quantifier like (a+)+ which can hang the server; rewrite it"
            )
        try:
            pattern = re.compile(query, 0 if case_sensitive else re.IGNORECASE)
        except re.error as e:
            raise SheetError(f"invalid regular expression: {e}") from e
    elif mode == "number":
        try:
            number = float(query.strip())
        except ValueError as e:
            raise SheetError(f"query {query!r} is not a number (use a dot as decimal separator)") from e
    needle = query if case_sensitive else fold(query)

    def matches(text: str) -> bool:
        text = text[:2000]
        if pattern is not None:
            return pattern.search(text) is not None
        return needle in (text if case_sensitive else fold(text))

    sheets = [info.worksheet(sheet)] if sheet else [s for s in info.sheets if s.kind == "worksheet"]
    hits: list[str] = []
    total = 0
    no_cache = 0
    for sc in sheets:
        if sc.max_row == 0:
            continue
        bounds: reader.Bounds = (1, 1, sc.max_row, min(sc.max_col, 16384))
        with closing(reader.rows(info, sc.name, bounds)) as rows:
            for row in rows:
                for c in row:
                    if c.value is None and c.formula is None:
                        continue
                    if c.no_cache:
                        no_cache += 1
                    ok = False
                    if mode == "number":
                        v = c.value
                        ok = (
                            isinstance(v, (int, float))
                            and not isinstance(v, bool)
                            and math.isclose(v, number or 0.0, rel_tol=1e-9, abs_tol=1e-12)
                        )
                    else:
                        if search_in in ("both", "values") and c.value is not None:
                            ok = matches(display(c.value, c.fmt))
                        if not ok and search_in in ("both", "formulas") and c.formula is not None:
                            ok = matches(c.formula)
                    if ok:
                        total += 1
                        if len(hits) < max_results:
                            shown = display(c.value, c.fmt) if c.value is not None else "(no cached value)"
                            line = f"{sc.name}!{c.coord}: {short(shown, 80)}"
                            if c.formula is not None:
                                line += f"   formula: {short(c.formula, 120)}"
                            hits.append(line)
    if not hits:
        extra = ""
        if no_cache and search_in != "formulas":
            extra = f" ({no_cache} formula cells have no cached value, so their results cannot be searched; try search_in='formulas')"
        return f"No cells match {query!r} ({mode}).{extra}"
    out = [f"{total} match(es) for {query!r} ({mode}):"] + hits
    if total > len(hits):
        out.append(f"... {total - len(hits)} more not shown (raise max_results or narrow with sheet=)")
    return "\n".join(out)


# --------------------------------------------------------------------- describe_columns

_NUMERIC_TEXT = re.compile(r"^\s*[-+]?\d+([.,]\d+)?\s*$")


def _kind(c: CellView) -> str | None:
    v = c.value
    if c.is_error:
        return "error"
    if v is None or (isinstance(v, str) and v == ""):
        return None
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, (int, float)):
        return "number"
    if isinstance(v, dt.datetime):
        return "datetime"
    if isinstance(v, dt.date):
        return "date"
    if isinstance(v, dt.time):
        return "time"
    if isinstance(v, dt.timedelta):
        return "duration"
    return "text"


class _Col:
    def __init__(self) -> None:
        self.kinds: Counter[str] = Counter()
        self.nulls = 0
        self.distinct: set[tuple[str, Any]] = set()
        self.distinct_capped = False
        self.top: Counter[str] = Counter()
        self.min: Any = None
        self.max: Any = None
        self.formulas = 0
        self.constants = 0
        self.no_cache = 0
        self.numeric_text = 0
        self.all_int = True
        self.dates_with_time = 0


def _header_auto(first: list[CellView], rest: list[list[CellView]]) -> bool:
    kinds = [_kind(c) for c in first]
    nonnull = [k for k in kinds if k is not None]
    if not nonnull or any(k != "text" for k in nonnull):
        return False
    texts = [str(c.value) for c in first if _kind(c) == "text"]
    if len(set(texts)) != len(texts):
        return False
    for k in range(len(first)):
        below = [_kind(r[k]) for r in rest[:50] if k < len(r)]
        below = [b for b in below if b is not None]
        if below and any(b != "text" for b in below):
            return True
    # all text: header if the header strings do not recur in their own column
    for k, c in enumerate(first):
        if _kind(c) == "text" and any(r[k].value == c.value for r in rest[:200] if k < len(r)):
            return False
    return len(rest) > 0


def describe_columns(
    info: PackageInfo, sheet: str, range_spec: str | None, header_row: int | None, max_rows: int
) -> str:
    scan = info.worksheet(sheet)
    bounds, clamp_note = reader.parse_range(range_spec, scan)
    min_row, min_col, max_row, max_col = bounds
    notes: list[str] = [clamp_note] if clamp_note else []
    if max_col - min_col + 1 > MAX_COLS:
        max_col = min_col + MAX_COLS - 1
        notes.append(f"only the first {MAX_COLS} columns are described")
    if not 1 <= max_rows <= 1_000_000:
        raise SheetError("max_rows must be between 1 and 1000000")
    if max_row - min_row + 1 > max_rows:
        notes.append(f"only the first {max_rows} rows were scanned (of {max_row - min_row + 1})")
        max_row = min_row + max_rows - 1
    width = max_col - min_col + 1
    cols = [_Col() for _ in range(width)]
    header: list[CellView] | None = None
    header_note = "none detected"
    skip_n = 0
    data_rows = 0
    buf: list[list[CellView]] = []
    resolved = False
    if header_row is not None and header_row != 0 and not min_row <= header_row < min_row + 200:
        raise SheetError(
            f"header_row {header_row} must be 0 (no header) or a row within the first 200 rows of the range "
            f"(which starts at row {min_row})"
        )
    with closing(reader.rows(info, scan.name, (min_row, min_col, max_row, max_col))) as rows:
        for r_idx, row in enumerate(rows):
            if resolved:
                _accumulate(cols, row)
                data_rows += 1
                continue
            buf.append(row)
            if len(buf) < 200 and min_row + r_idx < max_row:
                continue
            resolved = True
            if header_row is None:
                if _header_auto(buf[0], buf[1:]):
                    header, header_note, skip_n = buf[0], f"row {min_row} (auto-detected)", 1
            elif header_row == 0:
                header_note = "none (header_row=0)"
            else:
                skip_n = header_row - min_row + 1
                if skip_n > len(buf):
                    raise SheetError(f"header_row {header_row} is past the end of the range")
                header, header_note = buf[skip_n - 1], f"row {header_row} (given)"
            for pr in buf[skip_n:]:
                _accumulate(cols, pr)
                data_rows += 1
    first_data_row = min_row + skip_n
    lines = [
        f"{scan.name}!{reader.a1((min_row, min_col, max_row, max_col))}: {data_rows} data row(s), "
        f"{width} column(s). Header: {header_note or 'none'}. Data starts at row {first_data_row}.",
        "",
    ]
    for k, col in enumerate(cols):
        letter = get_column_letter(min_col + k)
        title = ""
        if header is not None and header[k].value not in (None, ""):
            title = f" {display(header[k].value)!r}"
        nonnull = sum(col.kinds.values())
        ktxt = ", ".join(f"{n} {kind}" for kind, n in col.kinds.most_common())
        main = col.kinds.most_common(1)[0][0] if col.kinds else "empty"
        if main == "number" and col.all_int and len(col.kinds) == 1:
            main = "integer"
        if len(col.kinds) > 1:
            main = "mixed"
        parts = [f"{letter}{title}: {main}", f"{nonnull} non-null, {col.nulls} empty"]
        dtxt = f"{len(col.distinct)}{'+' if col.distinct_capped else ''} distinct"
        parts.append(dtxt)
        if col.min is not None:
            parts.append(f"min {display(col.min)}, max {display(col.max)}")
        if main in ("text", "boolean") and col.top:
            parts.append("top: " + ", ".join(f"{short(v, 24)!r} x{n}" for v, n in col.top.most_common(3)))
        if col.formulas or col.constants:
            if col.formulas and col.constants:
                parts.append(f"MIXED formulas {col.formulas} / constants {col.constants}")
            elif col.formulas:
                parts.append(f"all {col.formulas} formulas")
            else:
                parts.append("all constants")
        if col.no_cache:
            parts.append(f"{col.no_cache} formulas without cached value")
        lines.append("- " + "; ".join(parts))
        if len(col.kinds) > 1:
            lines.append(f"    types: {ktxt}")
        if col.numeric_text and col.kinds.get("text"):
            lines.append(
                f"    {col.numeric_text} text cell(s) look like numbers (stored as text?)"
            )
    if notes:
        lines.append("")
        lines.extend(f"- {n}" for n in notes)
    return "\n".join(lines)


def _accumulate(cols: list[_Col], row: list[CellView]) -> None:
    for k, c in enumerate(row):
        col = cols[k]
        kind = _kind(c)
        if c.formula is not None:
            col.formulas += 1
            if c.no_cache:
                col.no_cache += 1
                continue
        elif kind is not None:
            col.constants += 1
        if kind is None:
            col.nulls += 1
            continue
        col.kinds[kind] += 1
        v = c.value
        if len(col.distinct) < 5000:
            col.distinct.add((kind, v if kind != "error" else str(v)))
        else:
            key = (kind, v)
            if key not in col.distinct:
                col.distinct_capped = True
        if kind == "text":
            if _NUMERIC_TEXT.match(v):
                col.numeric_text += 1
            if len(col.top) < 2000 or v in col.top:
                col.top[v] += 1
        elif kind == "boolean":
            col.top["TRUE" if v else "FALSE"] += 1
        elif kind in ("number", "date", "datetime", "time", "duration"):
            if kind == "number" and (isinstance(v, float) and not v.is_integer()):
                col.all_int = False
            if col.min is None or v < col.min:
                col.min = v
            if col.max is None or v > col.max:
                col.max = v
