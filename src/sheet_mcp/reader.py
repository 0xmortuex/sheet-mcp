"""Read cells straight from the worksheet XML in one streaming pass.

openpyxl's read-only mode needs two full passes to give both a formula and its cached value,
and re-parses every row above the range you ask for. This reader keeps formula text *and* the
cached result together, stops as soon as it is past the requested rows, and distinguishes
"formula with no cached value" from "empty". Its output is checked against openpyxl in the tests.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from collections.abc import Generator
from typing import Any

from defusedxml.ElementTree import iterparse
from openpyxl.formula.translate import Translator
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, from_excel

from .cells import CellView
from .errors import SheetError
from .package import PackageInfo, SheetScan, col_to_num

Bounds = tuple[int, int, int, int]  # min_row, min_col, max_row, max_col
_ADDR = re.compile(r"^([A-Z]+)(\d+)$")


def parse_range(spec: str | None, scan: SheetScan) -> tuple[Bounds, str | None]:
    """Parse A1 / A1:C9 / A:C / 2:5 and clamp to the used range.

    Returns (bounds, note). An empty or None spec means the whole used range.
    """
    if scan.max_row == 0 or scan.max_col == 0:
        raise SheetError(f"sheet {scan.name!r} has no data")
    used: Bounds = (1, 1, scan.max_row, scan.max_col)
    if spec is None or not spec.strip():
        return used, None
    try:
        c0, r0, c1, r1 = range_boundaries(spec.strip().upper())
    except (ValueError, TypeError) as e:
        raise SheetError(
            f"cannot parse range {spec!r}. Use A1, A1:D20, A:C (whole columns) or 3:9 (whole rows)."
        ) from e
    min_row, max_row = r0 or 1, r1 or 1048576
    min_col, max_col = c0 or 1, c1 or 16384
    if min_row > max_row or min_col > max_col:
        raise SheetError(f"range {spec!r} is inverted")
    if min_row > used[2] or min_col > used[3]:
        raise SheetError(
            f"range {spec!r} is entirely outside the used range "
            f"A1:{get_column_letter(used[3])}{used[2]} of sheet {scan.name!r}"
        )
    clamped = (min_row, min_col, min(max_row, used[2]), min(max_col, used[3]))
    note = None
    if (clamped[2], clamped[3]) != (max_row, max_col):
        note = (
            f"range clamped to the used range: "
            f"{get_column_letter(clamped[1])}{clamped[0]}:{get_column_letter(clamped[3])}{clamped[2]}"
        )
    return clamped, note


def a1(b: Bounds) -> str:
    return f"{get_column_letter(b[1])}{b[0]}:{get_column_letter(b[3])}{b[2]}"


def _local(tag: str) -> str:
    return tag.rpartition("}")[2]


def _si_text(el: Any) -> str:
    """Text of a shared/inline string item: plain and rich runs, never phonetic (rPh) text."""
    parts = []
    for child in el:
        ct = _local(child.tag)
        if ct == "t":
            parts.append(child.text or "")
        elif ct == "r":
            for t in child:
                if _local(t.tag) == "t":
                    parts.append(t.text or "")
    return "".join(parts)


def _shared_strings(zf: zipfile.ZipFile, part: str | None) -> list[str]:
    if part is None:
        return []
    out: list[str] = []
    with zf.open(part) as fh:
        for _ev, el in iterparse(fh, events=("end",)):
            if _local(el.tag) == "si":
                out.append(_si_text(el))
                el.clear()
    return out


def _styles(zf: zipfile.ZipFile, part: str | None) -> list[str]:
    """Number format code for each cellXfs index."""
    if part is None:
        return ["General"]
    custom: dict[int, str] = {}
    formats: list[int] = []
    in_cell_xfs = False
    with zf.open(part) as fh:
        for ev, el in iterparse(fh, events=("start", "end")):
            tag = _local(el.tag)
            if ev == "start" and tag == "cellXfs":
                in_cell_xfs = True
            elif ev == "end":
                if tag == "numFmt":
                    custom[int(el.get("numFmtId", "0"))] = el.get("formatCode", "General")
                elif tag == "cellXfs":
                    in_cell_xfs = False
                elif tag == "xf" and in_cell_xfs:
                    formats.append(int(el.get("numFmtId", "0")))
    return [custom.get(i) or BUILTIN_FORMATS.get(i) or "General" for i in formats] or ["General"]


def _cast_number(text: str) -> int | float:
    if "." in text or "E" in text or "e" in text:
        return float(text)
    return int(text)


class _Ctx:
    def __init__(self, zf: zipfile.ZipFile, info: PackageInfo) -> None:
        self.sst = _shared_strings(zf, info.shared_strings_part)
        self.fmts = _styles(zf, info.styles_part)
        self.is_date = [is_date_format(f) for f in self.fmts]
        self.is_delta = [is_timedelta_format(f) for f in self.fmts]
        self.epoch = CALENDAR_MAC_1904 if info.date1904 else CALENDAR_WINDOWS_1900
        self.masters: dict[str, tuple[str, str]] = {}


def rows(info: PackageInfo, sheet: str, bounds: Bounds) -> Generator[list[CellView], None, None]:
    """Yield one list of CellView per row of `bounds`, padded to a full rectangle.

    Close the generator (contextlib.closing) if you stop early.
    """
    scan = info.worksheet(sheet)
    min_row, min_col, max_row, max_col = bounds
    try:
        zf = zipfile.ZipFile(info.path)
    except (OSError, zipfile.BadZipFile) as e:
        raise SheetError(f"cannot reopen {info.path}: {e}") from e
    with zf:
        ctx = _Ctx(zf, info)
        cur: dict[int, CellView] = {}
        cur_row = min_row
        next_emit = min_row

        def pad(r: int) -> list[CellView]:
            return [cur.get(c) or CellView(f"{get_column_letter(c)}{r}") for c in range(min_col, max_col + 1)]

        with zf.open(scan.part) as fh:
            for _ev, el in iterparse(fh, events=("end",)):
                tag = _local(el.tag)
                if tag == "row":
                    el.clear()
                    continue
                if tag != "c":
                    continue
                ref = el.get("r")
                m = _ADDR.match(ref) if ref else None
                if m:
                    row, col = int(m.group(2)), col_to_num(m.group(1))
                else:
                    raise SheetError(f"{scan.name}: a cell has no address (r=); this file is not supported")
                if row < min_row:
                    continue
                if row > max_row:
                    break
                if col < min_col or col > max_col:
                    continue
                if row != cur_row:
                    while next_emit < row:
                        yield pad(next_emit)
                        cur = {}
                        next_emit += 1
                    cur_row = row
                cur[col] = _cell(el, f"{get_column_letter(col)}{row}", ctx)
        while next_emit <= max_row:
            yield pad(next_emit)
            cur = {}
            next_emit += 1


def _cell(el: Any, coord: str, ctx: _Ctx) -> CellView:
    t = el.get("t", "n")
    s = el.get("s")
    sid = int(s) if s else 0
    view = CellView(coord, fmt=ctx.fmts[sid] if sid < len(ctx.fmts) else "General")
    f_el = v_el = is_el = None
    for ch in el:
        lt = _local(ch.tag)
        if lt == "f":
            f_el = ch
        elif lt == "v":
            v_el = ch
        elif lt == "is":
            is_el = ch
    raw = v_el.text if v_el is not None and v_el.text else None
    is_formula = f_el is not None
    if f_el is not None:
        view.dtype = "f"
        view.formula = _formula(f_el, coord, ctx.masters)
        if raw is None and t != "str":
            view.no_cache = True

    def settle(dtype: str) -> None:
        if not is_formula:
            view.dtype = dtype

    if t == "inlineStr":
        if is_el is not None:
            view.value = _si_text(is_el)
            settle("s")
        return view
    if raw is None:
        return view
    if t == "n":
        num = _cast_number(raw)
        if sid < len(ctx.is_date) and ctx.is_date[sid]:
            try:
                view.value = from_excel(num, ctx.epoch, timedelta=ctx.is_delta[sid])
                settle("d")
            except (OverflowError, ValueError):
                view.value = "#VALUE!"
                if is_formula:
                    view.cached_error = True
                else:
                    view.dtype = "e"
        else:
            view.value = num
            settle("n")
    elif t == "s":
        view.value = ctx.sst[int(raw)]
        settle("s")
    elif t == "b":
        view.value = bool(int(raw))
        settle("b")
    elif t == "str":
        view.value = raw
        settle("s")
    elif t == "e":
        view.value = raw
        if is_formula:
            view.cached_error = True
        else:
            view.dtype = "e"
    elif t == "d":
        try:
            view.value = dt.datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            view.value = raw
        settle("d")
    else:
        view.value = raw
    return view


def _formula(f_el: Any, coord: str, masters: dict[str, tuple[str, str]]) -> str:
    kind = f_el.get("t")
    text = f_el.text or ""
    if kind == "shared":
        si = f_el.get("si", "")
        if text:
            masters[si] = (coord, text)
        elif si in masters:
            origin, base = masters[si]
            try:
                return str(Translator("=" + base, origin=origin).translate_formula(coord))
            except Exception as e:  # malformed shared formula
                raise SheetError(f"cannot expand the shared formula in {coord}: {e}") from e
        return "=" + text
    if kind == "array":
        return "{=" + text + "}" + f" [array {f_el.get('ref')}]"
    if kind == "dataTable":
        return f"{{=TABLE(row={f_el.get('r1')}, col={f_el.get('r2')})}} [data table {f_el.get('ref')}]"
    return "=" + text


def read_grid(info: PackageInfo, sheet: str, bounds: Bounds) -> list[list[CellView]]:
    return list(rows(info, sheet, bounds))
