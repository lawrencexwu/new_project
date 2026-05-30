"""Post-save XML injection of native Excel line sparklines.

openpyxl 3.x does not support the <extLst><sparklineGroups> Excel
extension (it warns and drops it on load). To get native line sparklines
in our workbooks, we let openpyxl save the file as usual, then crack
open the .xlsx (a zip), patch the worksheet XML to add sparkline groups,
and zip it back up.

The function below takes a workbook path and a dict mapping sheet name
to a list of sparkline specs: {data_range, target_cell}. Data ranges
reference cells on the same sheet (hidden columns written by the
populator).
"""
from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

# The "ext" uri Excel uses to identify sparkline extensions.
_SPARKLINE_EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"


def inject_line_sparklines(xlsx_path: Path,
                           sheet_sparklines: dict[str, list[dict]]) -> Path:
    """Add native line sparklines to a .xlsx file.

    sheet_sparklines: {sheet_name: [{"data_range": str, "target_cell": str}, ...]}
        data_range: e.g. "Daily Plan!T4:AS4" — must include the sheet name
        target_cell: e.g. "P4" — cell where the sparkline renders
    """
    xlsx_path = Path(xlsx_path)
    if not sheet_sparklines:
        return xlsx_path

    sheet_to_file = _resolve_sheet_files(xlsx_path)

    tmp = Path(tempfile.mkstemp(suffix=".xlsx")[1])
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                # Find which sheet this file belongs to
                sheet_name = next(
                    (sn for sn, fp in sheet_to_file.items() if fp == item),
                    None,
                )
                if sheet_name and sheet_name in sheet_sparklines:
                    sparks = sheet_sparklines[sheet_name]
                    if sparks:
                        data = _inject_into_sheet_xml(data, sparks)
                zout.writestr(item, data)
        shutil.move(str(tmp), str(xlsx_path))
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise

    return xlsx_path


def _resolve_sheet_files(xlsx_path: Path) -> dict[str, str]:
    """Map sheet name → workbook-relative XML file path."""
    with zipfile.ZipFile(xlsx_path) as z:
        wb_xml = etree.fromstring(z.read("xl/workbook.xml"))
        rels_xml = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))

    rid_to_target = {
        rel.get("Id"): rel.get("Target")
        for rel in rels_xml.findall(f"{{{_NS_PKG}}}Relationship")
    }

    out: dict[str, str] = {}
    sheets_el = wb_xml.find(f"{{{_NS_MAIN}}}sheets")
    if sheets_el is None:
        return out
    for sheet in sheets_el.findall(f"{{{_NS_MAIN}}}sheet"):
        name = sheet.get("name")
        rid = sheet.get(f"{{{_NS_REL}}}id")
        target = rid_to_target.get(rid)
        if target is None:
            continue
        # Targets in workbook.xml.rels are relative to xl/ (e.g.
        # 'worksheets/sheet1.xml') or absolute ('/xl/worksheets/sheet1.xml').
        # Normalize to the zip-internal path.
        t = target.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[name] = t
    return out


def _inject_into_sheet_xml(sheet_xml_bytes: bytes,
                           sparklines: list[dict]) -> bytes:
    """Return modified worksheet XML with an extLst → sparklineGroups appended.

    Matches xlsxwriter's pattern exactly:
    - Worksheet root: left untouched (openpyxl already produces a valid root).
    - <ext>: declares xmlns:x14 inline.
    - <x14:sparklineGroups>: declares xmlns:xm inline.

    Critical: we *do not* iterate root.children while appending — lxml's
    append() moves elements, which silently corrupts the iteration order
    when used together.
    """
    parser = etree.XMLParser(remove_blank_text=False)
    old_root = etree.fromstring(sheet_xml_bytes, parser)

    # openpyxl omits xmlns:r on the worksheet root; xlsxwriter (whose
    # sparkline files Excel accepts) always declares it. Rebuild the root
    # with xmlns:r added. extend(list(...)) materializes the children
    # eagerly so lxml's move-on-append doesn't corrupt iteration.
    if old_root.nsmap.get("r") == _NS_REL:
        root = old_root
    else:
        root = etree.Element(f"{{{_NS_MAIN}}}worksheet",
                             nsmap={None: _NS_MAIN, "r": _NS_REL})
        for k, v in old_root.attrib.items():
            root.set(k, v)
        root.extend(list(old_root))

    # Build the new extension element with namespaces declared inline.
    ext = etree.SubElement(
        _ensure_extLst(root),
        f"{{{_NS_MAIN}}}ext",
        nsmap={"x14": _NS_X14},
    )
    ext.set("uri", _SPARKLINE_EXT_URI)

    spark_groups = etree.SubElement(
        ext, f"{{{_NS_X14}}}sparklineGroups",
        nsmap={"xm": _NS_XM},
    )
    spark_group = etree.SubElement(spark_groups, f"{{{_NS_X14}}}sparklineGroup")
    spark_group.set("displayEmptyCellsAs", "gap")
    spark_group.set("type", "line")

    # Color attributes — order matches the OOXML XSD definition for
    # CT_SparklineGroup. Excel can be strict about element order.
    color_series = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorSeries")
    color_series.set("rgb", "FF1F4E79")
    color_negative = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorNegative")
    color_negative.set("rgb", "FFEF4444")
    color_axis = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorAxis")
    color_axis.set("rgb", "FF000000")
    color_markers = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorMarkers")
    color_markers.set("rgb", "FF1F4E79")
    color_first = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorFirst")
    color_first.set("rgb", "FF1F4E79")
    color_last = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorLast")
    color_last.set("rgb", "FF1F4E79")
    color_high = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorHigh")
    color_high.set("rgb", "FF10B981")
    color_low = etree.SubElement(spark_group, f"{{{_NS_X14}}}colorLow")
    color_low.set("rgb", "FFEF4444")

    sparklines_el = etree.SubElement(spark_group, f"{{{_NS_X14}}}sparklines")
    for sp in sparklines:
        sl = etree.SubElement(sparklines_el, f"{{{_NS_X14}}}sparkline")
        f_el = etree.SubElement(sl, f"{{{_NS_XM}}}f")
        f_el.text = sp["data_range"]
        sqref_el = etree.SubElement(sl, f"{{{_NS_XM}}}sqref")
        sqref_el.text = sp["target_cell"]

    # Write XML declaration manually — lxml emits single quotes there
    # which Excel's XmlLite parser rejects with HRESULT 0x808c0002.
    body = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + body
    )


def _ensure_extLst(root) -> etree._Element:
    """Return the worksheet's <extLst> element, creating it if absent."""
    extLst = root.find(f"{{{_NS_MAIN}}}extLst")
    if extLst is not None:
        return extLst
    extLst = etree.SubElement(root, f"{{{_NS_MAIN}}}extLst")
    return extLst
