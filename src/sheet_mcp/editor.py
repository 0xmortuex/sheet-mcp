"""Editing: load with openpyxl, apply a plan, verify, and write atomically to a new file."""

from __future__ import annotations

import os
import re
import tempfile
import uuid
import warnings
from copy import copy
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.cell_range import CellRange

from . import package, sandbox
from .cells import Coerced, apply, before_text, coerce, display
from .errors import SheetError
from .formulas import rename_in_formula
from .package import PackageInfo

MAX_CHANGES = 5000
MAX_APPEND_ROWS = 10000
_ADDR = re.compile(r"^([A-Za-z]{1,3})(\d+)$")
_BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_NOTE_AFTER_SAVE = (
    "Formulas are kept as formulas, but this save dropped their cached results (openpyxl cannot "
    "calculate): open the file in Excel/LibreOffice once to refresh them. Untouched cells, styles, "
    "validations, conditional formats, names and column widths are preserved."
)


# ---------------------------------------------------------------- paths


def output_path(src: Path, output: str | None, overwrite: bool) -> Path:
    if output:
        out = sandbox.resolve(output, must_exist=False, what="output")
    else:
        out = src.with_name(f"{src.stem}.edited{src.suffix}")
        sandbox.resolve(str(out), must_exist=False, what="output")
    if out.suffix.lower() != src.suffix.lower():
        raise SheetError(
            f"output must keep the extension {src.suffix} (got {out.suffix}); "
            "changing xlsx/xlsm would produce a file Excel refuses to open"
        )
    if out.exists():
        if not out.is_file():
            raise SheetError(f"output {out} exists and is not a file")
        if not overwrite:
            raise SheetError(
                f"{out.name} already exists. Nothing was written. Pass overwrite=true to replace it, "
                "or choose another output= name."
            )
    elif not out.parent.is_dir():
        raise SheetError(f"output directory {out.parent} does not exist")
    return out


# ---------------------------------------------------------------- load / save


def count_formulas(wb: Any) -> dict[str, int]:
    return {
        ws.title: sum(1 for c in ws._cells.values() if c.data_type == "f") for ws in wb.worksheets
    }


def refusal_text(info: PackageInfo) -> str:
    items = "\n".join(f"  - {t}" for t in package.describe_loss(info))
    return (
        "Refusing to edit: saving this workbook with openpyxl would LOSE:\n"
        f"{items}\n"
        "Nothing was written. If you accept that loss, repeat the call with allow_loss=true "
        "(the result still goes to a NEW file unless overwrite=true). Otherwise edit it in Excel."
    )


def load_for_edit(info: PackageInfo, allow_loss: bool, dry_run: bool) -> tuple[Any, list[str]]:
    """Load a workbook for editing. Returns (workbook, notes)."""
    notes: list[str] = []
    if info.lossy and not allow_loss:
        if not dry_run:
            raise SheetError(refusal_text(info))
        notes.append(
            "A real run would be REFUSED: saving would lose "
            + "; ".join(sorted(info.lossy))
            + ". Pass allow_loss=true to override (see workbook_overview)."
        )
    biggest = sum(s.xml_bytes for s in info.sheets)
    if biggest > package.MAX_EDIT_SHEET_BYTES:
        raise SheetError(
            f"the sheets hold {biggest} bytes of XML; editing loads everything into memory and is "
            f"limited to {package.MAX_EDIT_SHEET_BYTES}. Reading tools still work on this file."
        )
    if info.has_vba and not Path(info.path).suffix.lower() == ".xlsm":
        raise SheetError("macro-enabled workbook must be named .xlsm")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            wb = openpyxl.load_workbook(
                info.path, rich_text=True, keep_vba=info.has_vba, keep_links=True
            )
        except Exception as e:
            raise SheetError(f"openpyxl could not load the workbook: {type(e).__name__}: {e}") from e
    dropped = [str(w.message) for w in caught if "removed" in str(w.message).lower()]
    if dropped and not allow_loss:
        if dry_run:
            notes.append("A real run would be refused because openpyxl reports: " + "; ".join(dropped))
        else:
            raise SheetError(
                "openpyxl reports it will drop parts of this workbook: "
                + "; ".join(dropped)
                + ". Nothing was written. Pass allow_loss=true to accept."
            )
    mem = count_formulas(wb)
    disk = {s.name: s.n_formulas for s in info.sheets if s.kind == "worksheet"}
    if mem != disk and not allow_loss:
        diff = {k: (disk.get(k), mem.get(k)) for k in set(mem) | set(disk) if mem.get(k) != disk.get(k)}
        msg = f"openpyxl loaded a different number of formulas than the file contains {diff}"
        if dry_run:
            notes.append("A real run would be refused: " + msg)
        else:
            raise SheetError(msg + ". Nothing was written; pass allow_loss=true to accept.")
    return wb, notes


def save_atomic(wb: Any, out: Path, overwrite: bool) -> str:
    """Write to a temp file beside `out`, verify the formula counts survived, then os.replace."""
    expected = count_formulas(wb)
    if out.exists() and not overwrite:
        raise SheetError(f"{out.name} already exists; pass overwrite=true")
    fd, tmp_name = tempfile.mkstemp(dir=out.parent, prefix=".sheetmcp-", suffix=out.suffix)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        try:
            wb.save(tmp)
        except Exception as e:
            raise SheetError(f"openpyxl could not save the workbook: {type(e).__name__}: {e}") from e
        written = package._inspect_uncached(str(tmp))
        got = {s.name: s.n_formulas for s in written.sheets if s.kind == "worksheet"}
        if got != expected:
            raise SheetError(
                f"the saved file has different formula counts than expected {got} vs {expected}; "
                "nothing was written to the destination (this is a bug or an unsupported feature)."
            )
        try:
            os.replace(tmp, out)
        except PermissionError as e:
            raise SheetError(
                f"cannot replace {out.name}: {e}. Is the file open in Excel? Close it and retry."
            ) from e
    finally:
        if tmp.exists():
            tmp.unlink()
    n = sum(expected.values())
    return f"Saved {out} ({out.stat().st_size:,} bytes). {n} formula(s) kept. " + (
        _NOTE_AFTER_SAVE if n else "Untouched cells, styles and validations are preserved."
    )


# ---------------------------------------------------------------- helpers


def get_ws(wb: Any, name: str) -> Any:
    if name in wb.sheetnames:
        return wb[name]
    low = [n for n in wb.sheetnames if n.casefold() == name.casefold()]
    if len(low) == 1:
        return wb[low[0]]
    raise SheetError(f"no sheet named {name!r}. Sheets: {', '.join(repr(n) for n in wb.sheetnames)}")


def parse_addr(text: Any) -> tuple[int, int]:
    if not isinstance(text, str):
        raise SheetError(f"cell must be a string like 'B7', got {text!r}")
    m = _ADDR.match(text.strip())
    if not m:
        raise SheetError(
            f"cell {text!r} is not a single cell address like 'B7' (ranges are not allowed here)"
        )
    letters, row = m.group(1).upper(), int(m.group(2))
    col = 0
    for ch in letters:
        col = col * 26 + ord(ch) - 64
    if not (1 <= row <= 1048576 and 1 <= col <= 16384):
        raise SheetError(f"cell {text!r} is outside the sheet limits (XFD1048576)")
    return row, col


def _protected_cells(ws: Any) -> tuple[list[CellRange], list[CellRange]]:
    merged = list(ws.merged_cells.ranges)
    arrays = [CellRange(ref) for ref in ws.array_formulae.values()]
    return merged, arrays


def _guard_cell(ws: Any, row: int, col: int, merged: list[CellRange], arrays: list[CellRange]) -> None:
    addr = f"{get_column_letter(col)}{row}"
    for m in merged:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            if (row, col) != (m.min_row, m.min_col):
                raise SheetError(
                    f"{ws.title}!{addr} is inside the merged range {m.coord}; only its top-left cell "
                    f"{get_column_letter(m.min_col)}{m.min_row} holds a value. Nothing was written."
                )
    for a in arrays:
        if a.min_row <= row <= a.max_row and a.min_col <= col <= a.max_col:
            raise SheetError(
                f"{ws.title}!{addr} belongs to the array formula over {a.coord}; writing into part of "
                "an array formula would corrupt it. Nothing was written."
            )
    cell = ws._cells.get((row, col))
    if cell is not None and type(cell.value).__name__ == "DataTableFormula":
        raise SheetError(f"{ws.title}!{addr} is part of a what-if data table and cannot be edited")


def _table(plan: list[tuple[str, ...]], head: tuple[str, ...], cap: int = 60) -> list[str]:
    lines = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    for row in plan[:cap]:
        lines.append("| " + " | ".join(c.replace("|", "\\|").replace("\n", "\\n") for c in row) + " |")
    if len(plan) > cap:
        lines.append(f"... {len(plan) - cap} more rows not shown")
    return lines


def _finish(
    header: str,
    body: list[str],
    notes: list[str],
    warns: list[str],
    wb: Any,
    info: PackageInfo,
    dry_run: bool,
    out: Path,
    overwrite: bool,
) -> str:
    lines = [header, ""] + body
    for w in warns:
        lines.append(f"warning: {w}")
    for n in notes:
        lines.append(f"note: {n}")
    lines.append("")
    if dry_run:
        lines.append(
            f"DRY RUN: nothing was written. Repeat with dry_run=false to save to {out}"
            + (" (replacing the existing file)." if out.exists() else ".")
        )
        if info.total_formulas:
            lines.append("A real save drops cached formula results (openpyxl cannot calculate).")
    else:
        lines.append(save_atomic(wb, out, overwrite))
        if out.resolve() == Path(info.path).resolve():
            lines.append("The original file was replaced in place (overwrite=true).")
        else:
            lines.append(f"The original {Path(info.path).name} was not modified.")
    return "\n".join(lines)


# ---------------------------------------------------------------- set_cells


def set_cells(
    info: PackageInfo,
    sheet: str,
    changes: list[dict[str, Any]],
    dry_run: bool,
    output: str | None,
    overwrite: bool,
    allow_loss: bool,
) -> str:
    if not changes:
        raise SheetError("changes is empty")
    if len(changes) > MAX_CHANGES:
        raise SheetError(f"at most {MAX_CHANGES} changes per call")
    src = Path(info.path)
    out = output_path(src, output, overwrite)
    wb, notes = load_for_edit(info, allow_loss, dry_run)
    names = set(wb.sheetnames)
    plan: list[tuple[str, ...]] = []
    warns: list[str] = []
    guards: dict[str, tuple[list[CellRange], list[CellRange]]] = {}
    work: list[tuple[Any, int, int, Coerced]] = []
    seen: set[tuple[str, int, int]] = set()
    for i, ch in enumerate(changes):
        if not isinstance(ch, dict):
            raise SheetError(f"changes[{i}] must be an object like {{\"cell\": \"B2\", \"value\": 5}}")
        extra = set(ch) - {"cell", "value", "formula", "sheet"}
        if extra:
            raise SheetError(f"changes[{i}] has unknown key(s) {sorted(extra)}; allowed: cell, value, formula, sheet")
        if "value" in ch and "formula" in ch:
            raise SheetError(f"changes[{i}] has both value and formula; give one")
        if "value" not in ch and "formula" not in ch:
            raise SheetError(f"changes[{i}] needs 'value' (use null to clear the cell) or 'formula'")
        ws = get_ws(wb, ch.get("sheet") or sheet)
        row, col = parse_addr(ch.get("cell"))
        key = (ws.title, row, col)
        if key in seen:
            raise SheetError(f"changes[{i}]: {ws.title}!{get_column_letter(col)}{row} appears twice")
        seen.add(key)
        if ws.title not in guards:
            guards[ws.title] = _protected_cells(ws)
        _guard_cell(ws, row, col, *guards[ws.title])
        if "formula" in ch:
            f = ch["formula"]
            if not isinstance(f, str) or not f.strip():
                raise SheetError(f"changes[{i}].formula must be a non-empty string")
            c = coerce(f if f.startswith("=") else "=" + f, names)
        else:
            c = coerce(ch["value"], names)
        cell = ws._cells.get((row, col))
        before = before_text(cell.value if cell else None, cell.data_type if cell else "n")
        note = []
        if cell is not None and cell.data_type == "f" and c.kind != "formula":
            note.append("REPLACES A FORMULA")
        note.extend(c.warnings)
        warns.extend(f"{ws.title}!{get_column_letter(col)}{row}: {w}" for w in c.warnings)
        plan.append((f"{ws.title}!{get_column_letter(col)}{row}", before, c.describe(), "; ".join(
            n for n in note if n == "REPLACES A FORMULA")))
        work.append((ws, row, col, c))
    if not dry_run:
        for ws, row, col, c in work:
            apply(ws.cell(row=row, column=col), c)
    body = _table(plan, ("cell", "before", "after", "flag"))
    verb = "Would change" if dry_run else "Changed"
    return _finish(
        f"{verb} {len(work)} cell(s) in {info.path}", body, notes, sorted(set(warns)), wb, info, dry_run, out, overwrite
    )


# ---------------------------------------------------------------- append_rows


def append_rows(
    info: PackageInfo,
    sheet: str,
    rows: list[list[Any]],
    table: str | None,
    dry_run: bool,
    output: str | None,
    overwrite: bool,
    allow_loss: bool,
) -> str:
    if not rows:
        raise SheetError("rows is empty")
    if len(rows) > MAX_APPEND_ROWS:
        raise SheetError(f"at most {MAX_APPEND_ROWS} rows per call")
    for i, r in enumerate(rows):
        if not isinstance(r, list):
            raise SheetError(f"rows[{i}] must be a list of cell values")
    src = Path(info.path)
    out = output_path(src, output, overwrite)
    wb, notes = load_for_edit(info, allow_loss, dry_run)
    ws = get_ws(wb, sheet)
    names = set(wb.sheetnames)
    filled = [(r, c) for (r, c), cell in ws._cells.items() if cell.value is not None]
    last_row = max((r for r, _ in filled), default=0)
    min_col = min((c for _, c in filled), default=1)
    tbl = None
    if table:
        found = [t for t in ws.tables.values() if t.name.casefold() == table.casefold() or t.displayName.casefold() == table.casefold()]
        if not found:
            avail = ", ".join(repr(t.name) for t in ws.tables.values()) or "none on this sheet"
            raise SheetError(f"no table named {table!r} on sheet {ws.title!r} (tables: {avail})")
        tbl = found[0]
        c0, r0, c1, r1 = range_boundaries(tbl.ref)
        if tbl.totalsRowCount:
            raise SheetError(
                f"table {tbl.name!r} has a totals row; appending would have to move it. "
                "Use set_cells to write the row, or remove the totals row in Excel first."
            )
        start_row, min_col, width_limit, style_row = r1 + 1, c0, c1 - c0 + 1, r1
    else:
        for t in ws.tables.values():
            if range_boundaries(t.ref)[3] == last_row:
                raise SheetError(
                    f"the last data row belongs to table {t.name!r}. Pass table={t.name!r} to extend the "
                    "table, otherwise Excel will not treat the new rows as part of it."
                )
        start_row, width_limit, style_row = last_row + 1, 16384, last_row
    if start_row + len(rows) - 1 > 1048576:
        raise SheetError("not enough room: the sheet would pass row 1048576")
    width = max(len(r) for r in rows)
    if width > width_limit:
        raise SheetError(
            f"a row has {width} values but the destination is only {width_limit} columns wide"
            + (f" (table {tbl.name!r})" if tbl else "")
        )
    if min_col + width - 1 > 16384:
        raise SheetError("rows are too wide: they would pass column XFD")
    merged = list(ws.merged_cells.ranges)
    for k in range(len(rows)):
        for j in range(width):
            row, col = start_row + k, min_col + j
            if ws._cells.get((row, col)) is not None and ws._cells[(row, col)].value is not None:
                raise SheetError(
                    f"{ws.title}!{get_column_letter(col)}{row} is not empty; the space after the last "
                    "data row is already used. Nothing was written."
                )
            for m in merged:
                if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
                    raise SheetError(f"{ws.title}!{get_column_letter(col)}{row} is inside merged range {m.coord}")
    coerced: list[list[Coerced]] = []
    warns: list[str] = []
    for k, r in enumerate(rows):
        cr = [coerce(v, names) for v in r]
        coerced.append(cr)
        for j, c in enumerate(cr):
            warns.extend(f"row {start_row + k} col {get_column_letter(min_col + j)}: {w}" for w in c.warnings)
    last_new = start_row + len(rows) - 1
    # DV / CF ranges that end exactly at the old last row will not cover the new rows
    for dv in ws.data_validations.dataValidation:
        for rng in dv.sqref.ranges:
            if rng.max_row == style_row and style_row:
                warns.append(f"data validation on {rng.coord} does not extend to the new rows")
    for cf in ws.conditional_formatting:
        for rng in cf.sqref.ranges:
            if rng.max_row == style_row and style_row:
                warns.append(f"conditional formatting on {rng.coord} does not extend to the new rows")
    if not dry_run:
        for k, cr in enumerate(coerced):
            for j in range(width):
                row, col = start_row + k, min_col + j
                cell = ws.cell(row=row, column=col)
                src_cell = ws._cells.get((style_row, col)) if style_row else None
                if src_cell is not None and src_cell.has_style:
                    cell._style = copy(src_cell._style)
                if j < len(cr):
                    apply(cell, cr[j])
        if tbl is not None:
            c0, r0, c1, r1 = range_boundaries(tbl.ref)
            tbl.ref = f"{get_column_letter(c0)}{r0}:{get_column_letter(c1)}{last_new}"
            if tbl.autoFilter is not None:
                tbl.autoFilter.ref = tbl.ref
    preview = [
        (f"{start_row + k}", *[display(c.value) if c.kind != "formula" else str(c.value) for c in cr[:6]])
        for k, cr in enumerate(coerced[:5])
    ]
    body = [
        f"rows {start_row}-{last_new} starting at column {get_column_letter(min_col)}"
        + (f", extending table {tbl.name!r} to {get_column_letter(range_boundaries(tbl.ref)[0])}{range_boundaries(tbl.ref)[1]}:"
           f"{get_column_letter(range_boundaries(tbl.ref)[2])}{last_new}" if tbl is not None else ""),
        f"styles and number formats are copied from row {style_row}" if style_row else "no row above to copy styles from",
        "",
    ] + _table([tuple(p) for p in preview], ("row", *[f"col {i + 1}" for i in range(max(len(p) for p in preview) - 1)]), 5)
    verb = "Would append" if dry_run else "Appended"
    return _finish(
        f"{verb} {len(rows)} row(s) to {ws.title!r} in {info.path}", body, notes, sorted(set(warns)), wb, info, dry_run, out, overwrite
    )


# ---------------------------------------------------------------- add / rename sheet


def validate_sheet_name(name: str, existing: list[str], ignore: str | None = None) -> None:
    if not name or not name.strip():
        raise SheetError("sheet name is empty")
    if len(name) > 31:
        raise SheetError(f"sheet name {name!r} is {len(name)} characters; Excel allows at most 31")
    if _BAD_SHEET_CHARS.search(name):
        raise SheetError(f"sheet name {name!r} contains one of [ ] : * ? / \\ which Excel forbids")
    if name.startswith("'") or name.endswith("'"):
        raise SheetError("sheet name cannot start or end with an apostrophe")
    if name.casefold() == "history":
        raise SheetError("'History' is reserved by Excel")
    for e in existing:
        if e != ignore and e.casefold() == name.casefold():
            raise SheetError(f"a sheet named {e!r} already exists (names are case-insensitive)")


def add_sheet(
    info: PackageInfo,
    name: str,
    position: int | None,
    dry_run: bool,
    output: str | None,
    overwrite: bool,
    allow_loss: bool,
) -> str:
    src = Path(info.path)
    out = output_path(src, output, overwrite)
    wb, notes = load_for_edit(info, allow_loss, dry_run)
    validate_sheet_name(name, wb.sheetnames)
    if position is not None and not 1 <= position <= len(wb.sheetnames) + 1:
        raise SheetError(f"position must be between 1 and {len(wb.sheetnames) + 1}")
    if not dry_run:
        wb.create_sheet(title=name, index=None if position is None else position - 1)
    where = f"position {position}" if position else "the end"
    verb = "Would add" if dry_run else "Added"
    return _finish(
        f"{verb} an empty sheet {name!r} at {where} of {info.path}", [], notes, [], wb, info, dry_run, out, overwrite
    )


def _rename_dv_cf(ws: Any, old: str, new: str) -> int:
    n = 0
    for dv in ws.data_validations.dataValidation:
        for attr in ("formula1", "formula2"):
            v = getattr(dv, attr)
            if isinstance(v, str) and old.casefold() in v.casefold():
                nv = rename_in_formula(v, old, new)
                if nv != v:
                    setattr(dv, attr, nv)
                    n += 1
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            if rule.formula:
                for idx, v in enumerate(rule.formula):
                    if old.casefold() in v.casefold():
                        nv = rename_in_formula(v, old, new)
                        if nv != v:
                            rule.formula[idx] = nv
                            n += 1
    return n


def rename_sheet(
    info: PackageInfo,
    sheet: str,
    new_name: str,
    dry_run: bool,
    output: str | None,
    overwrite: bool,
    allow_loss: bool,
) -> str:
    src = Path(info.path)
    out = output_path(src, output, overwrite)
    wb, notes = load_for_edit(info, allow_loss, dry_run)
    target = get_ws(wb, sheet)
    old = target.title
    validate_sheet_name(new_name, wb.sheetnames, ignore=old)
    changed: list[tuple[str, str, str]] = []
    for ws in wb.worksheets:
        for (row, col), cell in ws._cells.items():
            if cell.data_type != "f":
                continue
            v = cell.value
            if isinstance(v, str):
                if old.casefold() not in v.casefold():
                    continue
                nv = rename_in_formula(v, old, new_name)
                if nv != v:
                    changed.append((f"{ws.title}!{get_column_letter(col)}{row}", v, nv))
                    cell.value = nv
            elif getattr(v, "text", None) and old.casefold() in v.text.casefold():
                nv = rename_in_formula(v.text, old, new_name)
                if nv != v.text:
                    changed.append((f"{ws.title}!{get_column_letter(col)}{row}", v.text, nv))
                    v.text = nv
    n_names = 0
    all_names = list(wb.defined_names.values())
    for ws in wb.worksheets:
        all_names.extend(ws.defined_names.values())
    for dn in all_names:
        txt = dn.attr_text
        if txt and old.casefold() in txt.casefold():
            nv = rename_in_formula(txt, old, new_name)
            if nv != txt:
                dn.attr_text = nv
                n_names += 1
    n_rules = sum(_rename_dv_cf(ws, old, new_name) for ws in wb.worksheets)
    n_links = 0
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            h = cell.hyperlink
            loc = getattr(h, "location", None)
            if loc and old.casefold() in loc.casefold():
                nv = rename_in_formula(loc, old, new_name)
                if nv != loc:
                    h.location = nv
                    n_links += 1
    if old.casefold() == new_name.casefold():
        # openpyxl de-duplicates a case-only rename against the sheet itself ('SALES1'); go via a temp name
        target.title = f"tmp{uuid.uuid4().hex[:20]}"
    target.title = new_name
    body = [
        f"{len(changed)} formula(s), {n_names} defined name(s), {n_rules} validation/conditional-format "
        f"formula(s) and {n_links} internal hyperlink(s) refer to the old name and are updated.",
        "",
    ] + _table([(a, b, c) for a, b, c in changed], ("cell", "before", "after"), 25)
    verb = "Would rename" if dry_run else "Renamed"
    return _finish(
        f"{verb} sheet {old!r} to {new_name!r} in {info.path}", body, notes, [], wb, info, dry_run, out, overwrite
    )

