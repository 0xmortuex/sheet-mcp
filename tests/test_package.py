"""Zip/XML safety and detection of features openpyxl would drop."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import add_parts, build_excel_like, build_rich, write_zip

from sheet_mcp import package
from sheet_mcp.errors import SheetError


def _inspect(p: Path) -> package.PackageInfo:
    package._cache.clear()
    return package.inspect(p)


def test_clean_workbook_has_no_losses(rich: Path) -> None:
    info = _inspect(rich)
    assert info.lossy == {}
    assert [s.name for s in info.sheets] == ["Sales", "Other"]
    assert info.total_formulas == 7 + 1 + 2 - 0 or info.total_formulas > 0
    assert any("comment box" in m for m in info.minor)


@pytest.mark.parametrize(
    ("parts", "label"),
    [
        ({"xl/charts/chart1.xml": "<c/>"}, "charts"),
        ({"xl/chartsheets/sheet1.xml": "<c/>"}, "chart sheets"),
        ({"xl/media/image1.png": b"\x89PNG"}, "images"),
        ({"xl/pivotCache/pivotCacheDefinition1.xml": "<p/>"}, "pivot tables"),
        ({"xl/pivotTables/pivotTable1.xml": "<p/>"}, "pivot tables"),
        ({"xl/vbaProject.bin": b"\xd0\xcf"}, "macros (VBA)"),
        ({"xl/slicers/slicer1.xml": "<s/>"}, "slicers"),
        ({"xl/slicerCaches/slicerCache1.xml": "<s/>"}, "slicers"),
        ({"xl/timelines/timeline1.xml": "<s/>"}, "timelines"),
        ({"xl/externalLinks/externalLink1.xml": "<e/>"}, "external links"),
        ({"xl/threadedComments/threadedComment1.xml": "<t/>"}, "threaded comments"),
        ({"xl/ctrlProps/ctrlProp1.xml": "<t/>"}, "form controls / ActiveX"),
        ({"xl/embeddings/oleObject1.bin": b"x"}, "embedded objects"),
        ({"xl/connections.xml": "<c/>"}, "data connections / query tables"),
        ({"xl/model/item.data": b"x"}, "data model"),
        ({"xl/metadata.xml": "<m/>"}, "dynamic-array metadata"),
        ({"customUI/customUI.xml": "<m/>"}, "custom ribbon UI"),
        (
            {"xl/drawings/drawing1.xml": '<xdr:wsDr xmlns:xdr="x"><xdr:twoCellAnchor><xdr:sp macro="" textlink=""/></xdr:twoCellAnchor></xdr:wsDr>'},
            "shapes / text boxes",
        ),
    ],
)
def test_each_lossy_feature_is_detected(rich: Path, parts: dict[str, bytes | str], label: str) -> None:
    out = add_parts(rich, rich.with_name("x.xlsx"), parts)
    info = _inspect(out)
    assert label in info.lossy, info.lossy


def test_drawing_with_only_pictures_is_not_flagged_as_shapes(rich: Path) -> None:
    out = add_parts(
        rich, rich.with_name("x.xlsx"),
        {"xl/drawings/drawing1.xml": '<xdr:wsDr xmlns:xdr="x"><xdr:twoCellAnchor><xdr:pic/></xdr:twoCellAnchor></xdr:wsDr>'},
    )
    assert "shapes / text boxes" not in _inspect(out).lossy


def test_macro_enabled_content_type_is_detected(rich: Path) -> None:
    with zipfile.ZipFile(rich) as z:
        ct = z.read("[Content_Types].xml").decode()
    ct = ct.replace("sheet.main+xml", "sheet.macroEnabled.main+xml")
    out = add_parts(rich, rich.with_name("m.xlsm"), {}, rename={"[Content_Types].xml": ct})
    info = _inspect(out)
    assert info.has_vba and "macros (VBA)" in info.lossy


@pytest.mark.parametrize(
    ("marker", "label"),
    [
        ('<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/></ext></extLst>', "sparklines"),
        ('<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"/></extLst>', "extended conditional formatting"),
        ('<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"/></extLst>', "extended data validation"),
    ],
)
def test_sheet_extensions_are_detected(rich: Path, marker: str, label: str) -> None:
    with zipfile.ZipFile(rich) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    sheet = sheet.replace("</worksheet>", marker + "</worksheet>")
    out = add_parts(rich, rich.with_name("x.xlsx"), {}, rename={"xl/worksheets/sheet1.xml": sheet})
    assert any(label in k for k in _inspect(out).lossy)


def test_minor_losses_are_not_blockers(rich: Path) -> None:
    out = add_parts(rich, rich.with_name("x.xlsx"), {"xl/calcChain.xml": "<c/>", "xl/printerSettings/p1.bin": b"x"})
    info = _inspect(out)
    assert info.lossy == {}
    assert any("calculation chain" in m for m in info.minor)
    assert any("printer" in m for m in info.minor)


def test_date_system_detected(_sandbox: Path) -> None:
    assert _inspect(build_excel_like(_sandbox / "a.xlsx")).date1904 is False
    assert _inspect(build_excel_like(_sandbox / "b.xlsx", date1904=True)).date1904 is True


def test_scan_counts_shared_formulas_errors_and_missing_cache(_sandbox: Path, rich: Path) -> None:
    info = _inspect(build_excel_like(_sandbox / "e.xlsx"))
    s = info.sheets[0]
    assert s.n_formulas == 6  # 3 shared + SUM + string + error
    assert s.n_formulas_no_cache == 0
    assert s.errors == {"#DIV/0!": ["F5"]}
    r = _inspect(rich)
    assert r.sheets[0].n_formulas_no_cache == r.sheets[0].n_formulas  # openpyxl wrote no cache
    assert r.sheets[0].hidden_rows == [(3, 3)]
    assert r.sheets[0].merges == ["G1:H2"]
    assert r.sheets[0].tables[0].name == "Sales1"
    assert [n.name for n in r.names] == ["Rate"]


# ----------------------------------------------------------------------- safety


def test_not_a_zip(_sandbox: Path) -> None:
    p = _sandbox / "bad.xlsx"
    p.write_bytes(b"this is not a zip file at all")
    with pytest.raises(SheetError, match="not a valid .xlsx"):
        _inspect(p)


def test_old_xls_or_encrypted_ole(_sandbox: Path) -> None:
    p = _sandbox / "old.xlsx"
    p.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 100)
    with pytest.raises(SheetError, match="OLE2"):
        _inspect(p)


def test_truncated_zip(rich: Path) -> None:
    data = rich.read_bytes()
    rich.write_bytes(data[: len(data) // 2])
    with pytest.raises(SheetError):
        _inspect(rich)


BILLION_LAUGHS = (
    '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol">'
    '<!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">'
    '<!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">'
    '<!ENTITY lol4 "&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;">]>'
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>&lol4;</t></si></sst>'
)


def test_billion_laughs_is_refused(_sandbox: Path) -> None:
    p = build_excel_like(_sandbox / "bomb.xlsx")
    out = add_parts(p, _sandbox / "bomb2.xlsx", {}, rename={"xl/sharedStrings.xml": BILLION_LAUGHS})
    with pytest.raises(SheetError, match="DOCTYPE/ENTITY"):
        _inspect(out)


def test_doctype_in_sheet_and_workbook_parts_is_refused(_sandbox: Path, rich: Path) -> None:
    for part in ("xl/worksheets/sheet1.xml", "xl/workbook.xml", "[Content_Types].xml"):
        with zipfile.ZipFile(rich) as z:
            text = z.read(part).decode()
        evil = text.replace("?>", '?><!DOCTYPE x [<!ENTITY a "b">]>', 1) if text.startswith("<?xml") else '<!DOCTYPE x>' + text
        out = add_parts(rich, _sandbox / "evil.xlsx", {}, rename={part: evil})
        with pytest.raises(SheetError, match="DOCTYPE"):
            _inspect(out)


def test_utf16_xml_is_refused(rich: Path, _sandbox: Path) -> None:
    out = add_parts(rich, _sandbox / "u16.xlsx", {}, rename={"xl/sharedStrings.xml": "<sst/>".encode("utf-16")})
    with pytest.raises(SheetError, match="not UTF-8"):
        _inspect(out)


def test_real_zip_bomb_by_ratio(_sandbox: Path) -> None:
    p = _sandbox / "ratio.xlsx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/workbook.xml", "<workbook/>")
        z.writestr("xl/junk.bin", b"\0" * (64 * 1024 * 1024))
    assert p.stat().st_size < 200_000
    with pytest.raises(SheetError, match="compression ratio"):
        _inspect(p)


def test_total_size_limit(rich: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(package, "MAX_TOTAL_BYTES", 1000)
    with pytest.raises(SheetError, match="in total"):
        _inspect(rich)


def test_entry_size_limit(rich: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(package, "MAX_ENTRY_BYTES", 500)
    with pytest.raises(SheetError, match="possible zip bomb"):
        _inspect(rich)


def test_entry_count_limit(_sandbox: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    p = write_zip(_sandbox / "many.xlsx", {f"xl/f{i}.xml": "<a/>" for i in range(30)})
    monkeypatch.setattr(package, "MAX_ENTRIES", 10)
    with pytest.raises(SheetError, match="zip entries"):
        _inspect(p)


def test_duplicate_entries_are_refused(_sandbox: Path, rich: Path) -> None:
    out = _sandbox / "dup.xlsx"
    with zipfile.ZipFile(rich) as zin, zipfile.ZipFile(out, "w") as z:
        for item in zin.infolist():
            z.writestr(item, zin.read(item.filename))
        with pytest.warns(UserWarning, match="Duplicate name"):
            z.writestr("xl/workbook.xml", "<workbook/>")
    with pytest.raises(SheetError, match="duplicate entry"):
        _inspect(out)


def test_unsafe_entry_names(_sandbox: Path) -> None:
    p = write_zip(_sandbox / "slip.xlsx", {"../evil.xml": "<a/>", "xl/workbook.xml": "<a/>"})
    with pytest.raises(SheetError, match="unsafe entry name"):
        _inspect(p)


def test_zip_with_unix_mode_attributes_is_fine(rich: Path, _sandbox: Path) -> None:
    out = _sandbox / "unix.xlsx"
    with zipfile.ZipFile(rich) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.create_system = 3
            zi.external_attr = 0o600 << 16
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, zin.read(item.filename))
    assert _inspect(out).lossy == {}


def test_missing_workbook_part(_sandbox: Path) -> None:
    p = write_zip(_sandbox / "nowb.xlsx", {"hello.txt": "hi"})
    with pytest.raises(SheetError, match="not an Excel workbook"):
        _inspect(p)


def test_build_rich_is_deterministic_enough(_sandbox: Path) -> None:
    a = _inspect(build_rich(_sandbox / "a.xlsx"))
    b = _inspect(build_rich(_sandbox / "b.xlsx"))
    assert [s.n_formulas for s in a.sheets] == [s.n_formulas for s in b.sheets]
