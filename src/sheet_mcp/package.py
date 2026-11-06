"""Inspect an .xlsx/.xlsm package without trusting it.

An xlsx is a zip of XML parts. Before openpyxl (or anything else) touches one we:

* cap entry count, per-entry and total uncompressed size, and compression ratio,
* stream-read every part and count the bytes that really come out,
* refuse any XML part that declares a DOCTYPE or ENTITY (billion-laughs and friends;
  Excel never writes one) or that is not UTF-8,
* inventory the features openpyxl cannot round-trip (charts, pivots, macros, ...),
* scan each worksheet once for the facts the overview and checks need.
"""

from __future__ import annotations

import os
import posixpath
import re
import zipfile
import zlib
from dataclasses import dataclass, field
from typing import Any

from defusedxml.ElementTree import fromstring, iterparse

from .errors import SheetError

MAX_ENTRIES = 10_000
MAX_ENTRY_BYTES = 512 * 1024 * 1024
MAX_TOTAL_BYTES = 1024 * 1024 * 1024
MAX_RATIO = 500  # deflate tops out near 1032:1 on zeros; real sheets are 5-30:1
RATIO_FLOOR = 1024 * 1024  # ratio only matters for entries larger than this
MAX_FILE_BYTES = 256 * 1024 * 1024
MAX_EDIT_SHEET_BYTES = 100 * 1024 * 1024  # sheet XML we are willing to load fully into memory

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

_XML_SUFFIXES = (".xml", ".rels", ".vml")
_BAD_XML = (b"<!doctype", b"<!entity")
_SHEET_MARKERS: dict[bytes, str] = {
    b"<x14:sparklinegroup": "sparklines",
    b"{05c60535-1f16-4fd2-b633-f4f36f0b64e0}": "sparklines",
    b"<x14:conditionalformatting": "extended conditional formatting (data bars / icon sets, x14)",
    b"{78c0d931-6437-407d-a8ee-f0aad7539e65}": "extended conditional formatting (data bars / icon sets, x14)",
    b"<x14:datavalidation": "extended data validation (lists from other sheets, x14)",
    b"{cce6a557-97bc-4b89-adb6-d9c93caab3df}": "extended data validation (lists from other sheets, x14)",
}
_SHAPE_RE = re.compile(rb"<(?:[\w.-]+:)?(?:sp|cxnSp|grpSp)[\s>/]")
_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")

# (label, why it matters) - surfaced verbatim to the agent.
LOSS_EXPLANATIONS: dict[str, str] = {
    "charts": "charts are re-created by openpyxl with reduced fidelity or dropped",
    "images": "pictures are re-encoded and lose cropping, effects and alt text (or are dropped without Pillow)",
    "shapes / text boxes": "drawing shapes and text boxes are dropped",
    "pivot tables": "pivot tables and their caches are not preserved faithfully",
    "macros (VBA)": "the VBA project may survive but buttons, ribbon and form wiring can be lost",
    "slicers": "slicers are dropped",
    "timelines": "timelines are dropped",
    "external links": "links to other workbooks may be rewritten or dropped",
    "threaded comments": "modern threaded comments are dropped (legacy notes remain)",
    "form controls / ActiveX": "form controls and ActiveX controls are dropped",
    "embedded objects": "embedded OLE objects are dropped",
    "data connections / query tables": "Power Query, connections and query tables are dropped",
    "data model": "the Power Pivot data model is dropped",
    "dynamic-array metadata": "dynamic array formulas (FILTER, SORT, UNIQUE, ...) degrade to implicit-intersection formulas",
    "custom ribbon UI": "custom ribbon XML is dropped",
    "ink annotations": "ink annotations are dropped",
    "rich data (images in cells, data types)": "rich data values are dropped",
    "chart sheets": "chart sheets are dropped",
    "sparklines": "sparklines are dropped",
    "extended conditional formatting (data bars / icon sets, x14)": "these x14 conditional formats are dropped",
    "extended data validation (lists from other sheets, x14)": "these x14 data validations are dropped",
}

_PREFIX_RULES: tuple[tuple[str, str], ...] = (
    ("xl/charts/", "charts"),
    ("xl/chartsheets/", "chart sheets"),
    ("xl/media/", "images"),
    ("xl/pivottables/", "pivot tables"),
    ("xl/pivotcache/", "pivot tables"),
    ("xl/slicers/", "slicers"),
    ("xl/slicercaches/", "slicers"),
    ("xl/timelines/", "timelines"),
    ("xl/timelinecaches/", "timelines"),
    ("xl/externallinks/", "external links"),
    ("xl/threadedcomments/", "threaded comments"),
    ("xl/persons/", "threaded comments"),
    ("xl/ctrlprops/", "form controls / ActiveX"),
    ("xl/activex/", "form controls / ActiveX"),
    ("xl/embeddings/", "embedded objects"),
    ("xl/querytables/", "data connections / query tables"),
    ("xl/model/", "data model"),
    ("xl/ink/", "ink annotations"),
    ("xl/richdata/", "rich data (images in cells, data types)"),
    ("customui/", "custom ribbon UI"),
)
_EXACT_RULES: dict[str, str] = {
    "xl/vbaproject.bin": "macros (VBA)",
    "xl/connections.xml": "data connections / query tables",
    "xl/metadata.xml": "dynamic-array metadata",
}
_MINOR_PREFIXES: tuple[tuple[str, str], ...] = (
    ("xl/calcchain.xml", "calculation chain (Excel rebuilds it)"),
    ("xl/printersettings/", "printer settings"),
    ("customxml/", "custom XML parts"),
    ("docprops/thumbnail", "file thumbnail"),
)


@dataclass
class SheetScan:
    name: str
    index: int
    state: str  # visible | hidden | veryHidden
    kind: str  # worksheet | chartsheet
    part: str
    declared_dim: str | None = None
    max_row: int = 0
    max_col: int = 0
    n_cells: int = 0
    n_formulas: int = 0
    n_array_formulas: int = 0
    n_formulas_no_cache: int = 0
    errors: dict[str, list[str]] = field(default_factory=dict)
    error_total: int = 0
    hidden_rows: list[tuple[int, int]] = field(default_factory=list)
    hidden_cols: list[tuple[int, int]] = field(default_factory=list)
    merges: list[str] = field(default_factory=list)
    n_validations: int = 0
    n_cond_formats: int = 0
    n_comments: int = 0
    tables: list[TableInfo] = field(default_factory=list)
    xml_bytes: int = 0


@dataclass
class TableInfo:
    name: str
    ref: str
    header_rows: int
    totals_rows: int
    sheet: str


@dataclass
class DefinedName:
    name: str
    text: str
    scope: str | None  # sheet name, or None for workbook scope
    hidden: bool


@dataclass
class PackageInfo:
    path: str
    file_size: int
    n_entries: int
    uncompressed: int
    date1904: bool
    sheets: list[SheetScan]
    names: list[DefinedName]
    lossy: dict[str, list[str]]  # label -> evidence (part names / sheet names)
    minor: list[str]
    has_vba: bool
    styles_part: str | None = None
    shared_strings_part: str | None = None

    @property
    def total_formulas(self) -> int:
        return sum(s.n_formulas for s in self.sheets)

    def sheet(self, name: str) -> SheetScan:
        for s in self.sheets:
            if s.name == name:
                return s
        low = [s for s in self.sheets if s.name.lower() == name.lower()]
        if len(low) == 1:
            return low[0]
        names = ", ".join(repr(s.name) for s in self.sheets)
        raise SheetError(f"no sheet named {name!r}. Sheets in this workbook: {names}")

    def worksheet(self, name: str) -> SheetScan:
        s = self.sheet(name)
        if s.kind != "worksheet":
            raise SheetError(f"{s.name!r} is a chart sheet, not a worksheet; it has no cells")
        return s


def col_to_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _rel_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def _xml_err(part: str, why: str) -> SheetError:
    return SheetError(
        f"unsafe or corrupt workbook: XML part {part!r} {why}. Refusing to open it."
    )


def _validate_zip(zf: zipfile.ZipFile, size: int) -> list[zipfile.ZipInfo]:
    infos = zf.infolist()
    if len(infos) > MAX_ENTRIES:
        raise SheetError(
            f"workbook has {len(infos)} zip entries (limit {MAX_ENTRIES}); refusing to open it"
        )
    seen: set[str] = set()
    total = 0
    for zi in infos:
        name = zi.filename
        low = name.lower()
        if low in seen:
            raise SheetError(f"workbook zip has a duplicate entry {name!r}; refusing to open it")
        seen.add(low)
        if name.startswith(("/", "\\")) or ".." in name.replace("\\", "/").split("/"):
            raise SheetError(f"workbook zip has an unsafe entry name {name!r}; refusing to open it")
        if zi.file_size > MAX_ENTRY_BYTES:
            raise SheetError(
                f"zip entry {name!r} expands to {zi.file_size} bytes (limit {MAX_ENTRY_BYTES}); "
                "refusing to open it (possible zip bomb)"
            )
        if zi.file_size > RATIO_FLOOR and zi.file_size > MAX_RATIO * max(zi.compress_size, 1):
            raise SheetError(
                f"zip entry {name!r} has a compression ratio above {MAX_RATIO}:1 "
                f"({zi.compress_size} -> {zi.file_size} bytes); refusing to open it (possible zip bomb)"
            )
        total += zi.file_size
    if total > MAX_TOTAL_BYTES:
        raise SheetError(
            f"workbook expands to {total} bytes in total (limit {MAX_TOTAL_BYTES}); "
            "refusing to open it (possible zip bomb)"
        )
    return infos


def _stream_check(zf: zipfile.ZipFile, zi: zipfile.ZipInfo) -> set[str]:
    """Read a whole part, verify its size and XML safety, return sheet feature labels."""
    is_xml = zi.filename.lower().endswith(_XML_SUFFIXES)
    is_sheet = zi.filename.lower().startswith("xl/worksheets/") and is_xml
    found: set[str] = set()
    read = 0
    tail = b""
    first = True
    with zf.open(zi) as fh:
        while True:
            chunk = fh.read(1024 * 1024)
            if not chunk:
                break
            read += len(chunk)
            if read > MAX_ENTRY_BYTES or read > zi.file_size:
                raise SheetError(
                    f"zip entry {zi.filename!r} is larger than its header claims; "
                    "refusing to open it (possible zip bomb)"
                )
            if not is_xml:
                continue
            if first:
                first = False
                if chunk[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in chunk[:64]:
                    raise _xml_err(zi.filename, "is not UTF-8 (UTF-16/32 XML is refused)")
            low = tail + chunk.lower()
            for bad in _BAD_XML:
                if bad in low:
                    raise _xml_err(zi.filename, "declares a DOCTYPE/ENTITY (XML entity expansion)")
            if is_sheet:
                for marker, label in _SHEET_MARKERS.items():
                    if marker in low:
                        found.add(label)
            tail = low[-48:]
    if read != zi.file_size:
        raise SheetError(f"zip entry {zi.filename!r} is shorter than its header claims; file is corrupt")
    return found


def _classify(names: list[str]) -> tuple[dict[str, list[str]], list[str]]:
    lossy: dict[str, list[str]] = {}
    minor: dict[str, None] = {}
    for n in names:
        low = n.lower()
        if low.endswith("/"):
            continue
        label = _EXACT_RULES.get(low)
        if label is None:
            for prefix, lab in _PREFIX_RULES:
                if low.startswith(prefix):
                    label = lab
                    break
        if label is not None:
            lossy.setdefault(label, []).append(n)
            continue
        for prefix, lab in _MINOR_PREFIXES:
            if low.startswith(prefix):
                minor[lab] = None
    return lossy, list(minor)


def _parse_rels(zf: zipfile.ZipFile, part: str) -> list[tuple[str, str, str]]:
    """(type, target-resolved, id) for the rels file of `part`."""
    d, base = posixpath.split(part)
    rels = posixpath.join(d, "_rels", base + ".rels")
    try:
        root = fromstring(zf.read(rels))
    except KeyError:
        return []
    out = []
    for r in root.iter(f"{{{NS_PKG_REL}}}Relationship"):
        if r.get("TargetMode") == "External":
            continue
        out.append((r.get("Type", ""), _rel_target(d, r.get("Target", "")), r.get("Id", "")))
    return out


def _scan_sheet(zf: zipfile.ZipFile, sc: SheetScan) -> None:
    hidden_rows: list[tuple[int, int]] = []
    with zf.open(sc.part) as fh:
        for _ev, el in iterparse(fh, events=("end",)):
            tag = el.tag.rpartition("}")[2]
            if tag == "c":
                has_f = has_v = has_is = False
                f_type = None
                v_text = None
                for ch in el:
                    ct = ch.tag.rpartition("}")[2]
                    if ct == "f":
                        has_f = True
                        f_type = ch.get("t")
                    elif ct == "v":
                        has_v = True
                        v_text = ch.text
                    elif ct == "is":
                        has_is = True
                if not (has_f or has_is or (has_v and v_text not in (None, ""))):
                    continue
                ref = el.get("r", "")
                m = _CELL_RE.match(ref)
                if m:
                    c, r = col_to_num(m.group(1)), int(m.group(2))
                    sc.max_row = max(sc.max_row, r)
                    sc.max_col = max(sc.max_col, c)
                sc.n_cells += 1
                t = el.get("t")
                if has_f:
                    sc.n_formulas += 1
                    if f_type == "array":
                        sc.n_array_formulas += 1
                    if (not has_v or not v_text) and t != "str":
                        sc.n_formulas_no_cache += 1
                if t == "e":
                    code = v_text or "#ERROR"
                    sc.error_total += 1
                    lst = sc.errors.setdefault(code, [])
                    if len(lst) < 50:
                        lst.append(ref)
            elif tag == "row":
                if el.get("hidden") in ("1", "true"):
                    r = int(el.get("r", "0"))
                    if hidden_rows and hidden_rows[-1][1] == r - 1:
                        hidden_rows[-1] = (hidden_rows[-1][0], r)
                    else:
                        hidden_rows.append((r, r))
                el.clear()
            elif tag == "col":
                if el.get("hidden") in ("1", "true"):
                    sc.hidden_cols.append((int(el.get("min", "0")), int(el.get("max", "0"))))
            elif tag == "mergeCell":
                ref = el.get("ref")
                if ref:
                    sc.merges.append(ref)
            elif tag == "dimension":
                sc.declared_dim = el.get("ref")
            elif tag == "dataValidation":
                sc.n_validations += 1
            elif tag == "conditionalFormatting":
                sc.n_cond_formats += 1
    sc.hidden_rows = hidden_rows


def _count_children(zf: zipfile.ZipFile, part: str, local: str) -> int:
    n = 0
    with zf.open(part) as fh:
        for _ev, el in iterparse(fh, events=("end",)):
            if el.tag.rpartition("}")[2] == local:
                n += 1
            el.clear()
    return n


def _inspect_uncached(path: str) -> PackageInfo:
    try:
        size = os.path.getsize(path)
    except OSError as e:
        raise SheetError(f"cannot stat {path}: {e}") from e
    if size > MAX_FILE_BYTES:
        raise SheetError(f"file is {size} bytes (limit {MAX_FILE_BYTES}); refusing to open it")
    with open(path, "rb") as fh:
        head = fh.read(8)
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        raise SheetError(
            "this is an OLE2 file: an old .xls, or an encrypted/password-protected xlsx. "
            "Save an unprotected .xlsx from Excel and try again."
        )
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        raise SheetError(f"not a valid .xlsx (zip) file: {e}") from e
    try:
        with zf:
            infos = _validate_zip(zf, size)
            sheet_features: set[str] = set()
            for zi in infos:
                try:
                    sheet_features |= _stream_check(zf, zi)
                except (zipfile.BadZipFile, zlib.error, EOFError, NotImplementedError, RuntimeError) as e:
                    raise SheetError(f"corrupt or unsupported zip entry {zi.filename!r}: {e}") from e
            return _build_info(zf, path, size, infos, sheet_features)
    except SheetError:
        raise
    except Exception as e:  # malformed XML structure in an otherwise safe package
        raise SheetError(f"could not read the workbook structure: {type(e).__name__}: {e}") from e


def _build_info(
    zf: zipfile.ZipFile,
    path: str,
    size: int,
    infos: list[zipfile.ZipInfo],
    sheet_features: set[str],
) -> PackageInfo:
    names = [zi.filename for zi in infos]
    lossy, minor = _classify(names)
    for label in sheet_features:
        lossy.setdefault(label, []).append("worksheet XML")

    has_vba = "macros (VBA)" in lossy
    if "[Content_Types].xml" in names:
        ct = zf.read("[Content_Types].xml")
        if b"macroEnabled" in ct:
            has_vba = True
            lossy.setdefault("macros (VBA)", []).append("[Content_Types].xml (macro-enabled)")

    # drawings: shapes and text boxes (pictures/charts are caught by their own parts)
    for n in names:
        ln = n.lower()
        if ln.startswith("xl/drawings/") and ln.endswith(".xml") and "/_rels/" not in ln:
            zi = zf.getinfo(n)
            if zi.file_size <= 20 * 1024 * 1024 and _SHAPE_RE.search(zf.read(n)):
                lossy.setdefault("shapes / text boxes", []).append(n)

    root_rels = _parse_rels(zf, "")  # "" -> "_rels/.rels"
    wb_part = next(
        (t for typ, t, _ in root_rels if typ.endswith("/officeDocument")), "xl/workbook.xml"
    )
    if wb_part not in names:
        raise SheetError("not an Excel workbook: no workbook part found")
    wb_root = fromstring(zf.read(wb_part))
    wb_rels = {rid: (typ, t) for typ, t, rid in _parse_rels(zf, wb_part)}

    styles_part = next((t for typ, t in wb_rels.values() if typ.endswith("/styles")), None)
    sst_part = next((t for typ, t in wb_rels.values() if typ.endswith("/sharedStrings")), None)
    date1904 = False
    pr = wb_root.find(f"{{{NS_MAIN}}}workbookPr")
    if pr is not None and pr.get("date1904") in ("1", "true"):
        date1904 = True

    scans: list[SheetScan] = []
    sheets_el = wb_root.find(f"{{{NS_MAIN}}}sheets")
    if sheets_el is not None:
        for i, s in enumerate(sheets_el):
            rid = s.get(f"{{{NS_REL}}}id", "")
            typ, target = wb_rels.get(rid, ("", ""))
            if target not in names:
                raise SheetError(f"sheet {s.get('name')!r} points at a missing part {target!r}")
            kind = "chartsheet" if typ.endswith("/chartsheet") else "worksheet"
            sc = SheetScan(
                name=s.get("name", ""),
                index=i,
                state=s.get("state", "visible"),
                kind=kind,
                part=target,
                xml_bytes=zf.getinfo(target).file_size,
            )
            if kind == "worksheet":
                _scan_sheet(zf, sc)
                for ref in sc.merges:  # Excel's used range includes merged areas
                    m = re.match(r"^[A-Z]+\d+(?::([A-Z]+)(\d+))?$", ref)
                    if m and m.group(1):
                        sc.max_col = max(sc.max_col, col_to_num(m.group(1)))
                        sc.max_row = max(sc.max_row, int(m.group(2)))
                for typ2, t2, _ in _parse_rels(zf, target):
                    if t2 not in names:
                        continue
                    if typ2.endswith("/table"):
                        troot = fromstring(zf.read(t2))
                        sc.tables.append(
                            TableInfo(
                                name=troot.get("displayName") or troot.get("name", ""),
                                ref=troot.get("ref", ""),
                                header_rows=int(troot.get("headerRowCount", "1")),
                                totals_rows=int(troot.get("totalsRowCount", "0")),
                                sheet=sc.name,
                            )
                        )
                    elif typ2.endswith("/comments"):
                        sc.n_comments += _count_children(zf, t2, "comment")
            scans.append(sc)

    if any(s.n_comments for s in scans):
        minor.append("comment box sizes and positions (text and authors are kept)")

    dn: list[DefinedName] = []
    dn_el = wb_root.find(f"{{{NS_MAIN}}}definedNames")
    if dn_el is not None:
        for d in dn_el:
            local = d.get("localSheetId")
            scope = None
            if local is not None and local.isdigit() and int(local) < len(scans):
                scope = scans[int(local)].name
            dn.append(
                DefinedName(
                    name=d.get("name", ""),
                    text=d.text or "",
                    scope=scope,
                    hidden=d.get("hidden") in ("1", "true"),
                )
            )
    return PackageInfo(
        path=path,
        file_size=size,
        n_entries=len(infos),
        uncompressed=sum(zi.file_size for zi in infos),
        date1904=date1904,
        sheets=scans,
        names=dn,
        lossy=lossy,
        minor=minor,
        has_vba=has_vba,
        styles_part=styles_part if styles_part in names else None,
        shared_strings_part=sst_part if sst_part in names else None,
    )


_cache: dict[tuple[str, int, int], PackageInfo] = {}


def inspect(path: str | os.PathLike[str]) -> PackageInfo:
    """Inspect (and safety-check) a workbook; results are cached per file version."""
    p = os.fspath(path)
    try:
        st = os.stat(p)
    except OSError as e:
        raise SheetError(f"cannot stat {p}: {e}") from e
    key = (os.path.abspath(p), st.st_mtime_ns, st.st_size)
    hit = _cache.get(key)
    if hit is not None:
        return hit
    info = _inspect_uncached(p)
    if len(_cache) >= 8:
        _cache.pop(next(iter(_cache)))
    _cache[key] = info
    return info


def describe_loss(info: PackageInfo) -> list[str]:
    out = []
    for label, evidence in info.lossy.items():
        why = LOSS_EXPLANATIONS.get(label, "would not survive a save")
        out.append(f"{label} ({why}); found in: {', '.join(sorted(set(evidence))[:3])}")
    return out


def short(v: Any, n: int = 60) -> str:
    s = str(v)
    return s if len(s) <= n else s[: n - 3] + "..."
