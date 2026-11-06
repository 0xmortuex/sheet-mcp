"""CSV/TSV <-> xlsx with explicit encoding and delimiter."""

from __future__ import annotations

import codecs
import csv
import datetime as dt
import os
import re
import tempfile
from contextlib import closing
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell import WriteOnlyCell

from . import reader, sandbox
from .cells import MAX_TEXT, display
from .editor import validate_sheet_name
from .errors import SheetError
from .package import PackageInfo

MAX_CSV_BYTES = 200 * 1024 * 1024
_DELIMS = {"comma": ",", "semicolon": ";", "tab": "\t", "pipe": "|", "\\t": "\t"}
_NUMBER = re.compile(r"^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$")
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_ISO_DT = re.compile(r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?$")
_ERROR_CODES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


def parse_delimiter(d: str) -> str:
    d = _DELIMS.get(d.lower(), d)
    if len(d) != 1 or d in "\r\n\"":
        raise SheetError(
            f"delimiter must be one character (or comma/semicolon/tab/pipe), got {d!r}"
        )
    return d


def check_encoding(name: str) -> str:
    try:
        return codecs.lookup(name).name
    except LookupError as e:
        raise SheetError(
            f"unknown encoding {name!r}. Common: utf-8, utf-8-sig (Excel's BOM), cp1254 (Turkish), "
            "cp1252, latin-1, utf-16."
        ) from e


def _infer(s: str) -> Any:
    if _NUMBER.match(s) and len(s.replace("-", "").replace(".", "")) <= 15 and "e" not in s.lower():
        return float(s) if "." in s else int(s)
    if _ISO_DATE.match(s):
        try:
            return dt.date.fromisoformat(s)
        except ValueError:
            return s
    if _ISO_DT.match(s):
        try:
            return dt.datetime.fromisoformat(s)
        except ValueError:
            return s
    return s


def _cell(ws: Any, v: Any) -> Any:
    if isinstance(v, str) and (v.startswith("=") or v in _ERROR_CODES):
        c = WriteOnlyCell(ws, value=v)
        c.data_type = "s"
        return c
    if isinstance(v, dt.datetime):
        c = WriteOnlyCell(ws, value=v)
        c.number_format = "yyyy-mm-dd hh:mm:ss"
        return c
    if isinstance(v, dt.date):
        c = WriteOnlyCell(ws, value=v)
        c.number_format = "yyyy-mm-dd"
        return c
    return v


def _tmp_beside(out: Path) -> Path:
    fd, name = tempfile.mkstemp(dir=out.parent, prefix=".sheetmcp-", suffix=out.suffix)
    os.close(fd)
    return Path(name)


def _check_out(out: Path, overwrite: bool) -> None:
    if out.exists() and not overwrite:
        raise SheetError(
            f"{out.name} already exists. Nothing was written. Pass overwrite=true or choose another output."
        )
    if not out.parent.is_dir():
        raise SheetError(f"output directory {out.parent} does not exist")


def csv_to_xlsx(
    csv_path: str,
    output: str,
    encoding: str,
    delimiter: str,
    sheet_name: str,
    infer_types: bool,
    overwrite: bool,
) -> str:
    src = sandbox.resolve(csv_path, must_exist=True, what="csv")
    out = sandbox.resolve(output, must_exist=False, what="output")
    if out.suffix.lower() != ".xlsx":
        raise SheetError("output must end in .xlsx")
    _check_out(out, overwrite)
    enc = check_encoding(encoding)
    delim = parse_delimiter(delimiter)
    validate_sheet_name(sheet_name, [])
    size = src.stat().st_size
    if size > MAX_CSV_BYTES:
        raise SheetError(f"{src.name} is {size} bytes (limit {MAX_CSV_BYTES})")
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title=sheet_name)
    n_rows = max_cols = 0
    tmp = _tmp_beside(out)
    try:
        try:
            with open(src, encoding=enc, newline="") as fh:
                for n_rows, rec in enumerate(csv.reader(fh, delimiter=delim), start=1):
                    if n_rows > 1048576:
                        raise SheetError("more than 1,048,576 rows: Excel cannot hold this file")
                    if len(rec) > 16384:
                        raise SheetError(f"row {n_rows} has {len(rec)} fields; Excel's limit is 16384")
                    for v in rec:
                        if len(v) > MAX_TEXT:
                            raise SheetError(f"row {n_rows}: a field has {len(v)} characters (Excel limit {MAX_TEXT})")
                    max_cols = max(max_cols, len(rec))
                    ws.append([_cell(ws, _infer(v) if infer_types else v) for v in rec])
        except UnicodeError as e:
            raise SheetError(
                f"{src.name} is not valid {enc}: {e}. Try another encoding (utf-8-sig, cp1254, cp1252, latin-1)."
            ) from e
        except csv.Error as e:
            raise SheetError(
                f"{src.name} is not parseable with delimiter {delim!r}: {e}. Check the delimiter."
            ) from e
        wb.save(tmp)
        os.replace(tmp, out)
    finally:
        if not ws.closed:  # an aborted write-only sheet must be finished or it warns at garbage collection
            ws.close()
        wb.close()
        if tmp.exists():
            tmp.unlink()
    return (
        f"Wrote {out} from {src.name}: sheet {sheet_name!r}, {n_rows} rows x {max_cols} columns "
        f"(encoding {enc}, delimiter {delim!r}). "
        + (
            "Types inferred: plain numbers (no leading zeros, at most 15 digits) and ISO dates became "
            "real values; everything else, including text starting with '=', stays text."
            if infer_types
            else "Every field was stored as text."
        )
    )


def _csv_value(c: Any, mode: str) -> str:
    if c.formula is not None and mode == "formulas":
        return str(c.formula)
    if isinstance(c.value, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        return display(c.value, c.fmt)
    if isinstance(c.value, float) and c.value.is_integer() and abs(c.value) < 1e15:
        return str(int(c.value))
    return display(c.value, c.fmt)


def xlsx_to_csv(
    info: PackageInfo,
    sheet: str,
    output: str,
    range_spec: str | None,
    encoding: str,
    delimiter: str,
    mode: str,
    overwrite: bool,
) -> str:
    if mode not in ("values", "formulas"):
        raise SheetError("mode must be 'values' or 'formulas'")
    out = sandbox.resolve(output, must_exist=False, what="output")
    if out.suffix.lower() not in (".csv", ".tsv", ".txt"):
        raise SheetError("output must end in .csv, .tsv or .txt")
    _check_out(out, overwrite)
    enc = check_encoding(encoding)
    delim = parse_delimiter(delimiter)
    scan = info.worksheet(sheet)
    bounds, note = reader.parse_range(range_spec, scan)
    tmp = _tmp_beside(out)
    n_rows = no_cache = 0
    try:
        try:
            with open(tmp, "w", encoding=enc, newline="") as fh:
                w = csv.writer(fh, delimiter=delim)
                with closing(reader.rows(info, scan.name, bounds)) as rows:
                    for row in rows:
                        no_cache += sum(1 for c in row if c.no_cache)
                        w.writerow([_csv_value(c, mode) for c in row])
                        n_rows += 1
        except UnicodeEncodeError as e:
            raise SheetError(
                f"the data contains characters {enc} cannot encode ({e.reason} at {e.object[e.start:e.end]!r}). "
                "Use encoding utf-8 or utf-8-sig."
            ) from e
        os.replace(tmp, out)
    finally:
        if tmp.exists():
            tmp.unlink()
    msg = (
        f"Wrote {out}: {scan.name}!{reader.a1(bounds)} as {n_rows} rows "
        f"(encoding {enc}, delimiter {delim!r}, mode {mode}). Dates are ISO 8601."
    )
    if note:
        msg += f" Note: {note}."
    if no_cache and mode == "values":
        msg += (
            f" WARNING: {no_cache} formula cell(s) have no cached value and were written empty "
            "(the file was never calculated by Excel/LibreOffice)."
        )
    return msg


