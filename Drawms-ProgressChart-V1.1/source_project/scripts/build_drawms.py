"""Build the standalone drawms milestone text chart workbook."""
from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as ET

from vba_project import rebuild_vba_modules

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_NAME = "drawms.xlsm"
DRAWMS_COLUMN_WIDTHS = (3.5, 40, 28)
EXAMPLES = [
    ("First samples\nprocessing", "2026/8/20"),
    ("Appearance / size\ninspection", "2026/8/22"),
    ("Probe electrical\nvalidation", "2026/8/25"),
    ("Wire bond\nvalidation", "2026/8/28"),
    ("Three-lot stability\nvalidation", "2026/9/5"),
    ("Production release", "2026/9/10"),
]
NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def q(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def sub(parent, prefix: str, tag: str, **attrs):
    return ET.SubElement(parent, q(prefix, tag), {k: str(v) for k, v in attrs.items()})


def xml(root) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True, standalone=True)


def emu(points: float) -> str:
    return str(round(points * 12700))


def transform(parent, x: float, y: float, width: float, height: float):
    xf = sub(parent, "a", "xfrm")
    sub(xf, "a", "off", x=emu(x), y=emu(y))
    sub(xf, "a", "ext", cx=emu(width), cy=emu(height))
    return xf


def text_body(parent, value: str, *, bold: bool = False, color: str = "000000",
              size: int = 1000, center: bool = True):
    body = sub(parent, "xdr", "txBody")
    sub(body, "a", "bodyPr", wrap="square", lIns=0, rIns=0, tIns=0, bIns=0, anchor="t")
    sub(body, "a", "lstStyle")
    for line in value.split("\n"):
        para = sub(body, "a", "p")
        sub(para, "a", "pPr", algn="ctr" if center else "l")
        run = sub(para, "a", "r")
        props = sub(run, "a", "rPr", lang="en-US", sz=size, b=int(bold))
        sub(sub(props, "a", "solidFill"), "a", "srgbClr", val=color)
        sub(props, "a", "latin", typeface="Arial")
        sub(props, "a", "ea", typeface="Arial")
        sub(run, "a", "t").text = line
        sub(para, "a", "endParaRPr", lang="en-US", sz=size)


def dm_shape(parent, idx: int, name: str, x: float, y: float, width: float,
             height: float, *, geometry: str = "rect", color: str | None = None,
             text: str | None = None, bold: bool = False,
             macro: str | None = None, text_color: str = "000000"):
    attrs = {"macro": "[0]!DrawMs." + macro} if macro else {}
    shape = sub(parent, "xdr", "sp", **attrs)
    nv = sub(shape, "xdr", "nvSpPr")
    sub(nv, "xdr", "cNvPr", id=idx, name=name)
    sub(nv, "xdr", "cNvSpPr", **({"txBox": "1"} if text is not None and color is None else {}))
    props = sub(shape, "xdr", "spPr")
    transform(props, x, y, width, height)
    geom = sub(props, "a", "prstGeom", prst=geometry)
    av = sub(geom, "a", "avLst")
    if geometry == "rightArrow":
        sub(av, "a", "gd", name="adj1", fmla="val 55000")
        sub(av, "a", "gd", name="adj2", fmla="val 60000")
    if color:
        sub(sub(props, "a", "solidFill"), "a", "srgbClr", val=color)
    else:
        sub(props, "a", "noFill")
    sub(sub(props, "a", "ln"), "a", "noFill")
    if text is not None:
        text_body(shape, text, bold=bold, color=text_color)
        if color:
            shape.find("xdr:txBody/a:bodyPr", NS).set("anchor", "ctr")
    return shape


def anchor(parent, x: float, y: float, width: float, height: float):
    node = sub(parent, "xdr", "absoluteAnchor")
    sub(node, "xdr", "pos", x=emu(x), y=emu(y))
    sub(node, "xdr", "ext", cx=emu(width), cy=emu(height))
    return node


def dm_button(parent, idx: int, label: str, x: float, y: float, width: float,
              macro: str, color: str):
    node = anchor(parent, x, y, width, 25)
    dm_shape(node, idx, "DM_Button_" + macro, x, y, width, 25,
             geometry="roundRect", color=color, text=label, bold=True,
             macro=macro, text_color="FFFFFF")
    sub(node, "xdr", "clientData", fPrintsWithSheet="0")


def chart_group(x: float = 650, y: float = 50, width: float = 690):
    drawing = ET.Element(q("xdr", "wsDr"), nsmap={"xdr": NS["xdr"], "a": NS["a"]})
    dm_button(drawing, 2, "Generate Chart", 25, 93, 171, "GenerateChart", "2563EB")
    # Keep the two surviving controls at their fix3 worksheet coordinates.
    dm_button(drawing, 3, "Widen Spacing", 335, 93, 125, "WidenSpacing", "7C3AED")
    dm_button(drawing, 4, "Narrow Spacing", 472, 93, 125, "NarrowSpacing", "64748B")

    height = 126
    node = anchor(drawing, x, y, width, height)
    group = sub(node, "xdr", "grpSp")
    nv = sub(group, "xdr", "nvGrpSpPr")
    sub(nv, "xdr", "cNvPr", id=10, name="DrawMs_Chart",
        descr="drawms: six sample milestones in input row order.")
    sub(nv, "xdr", "cNvGrpSpPr")
    xf = transform(sub(group, "xdr", "grpSpPr"), x, y, width, height)
    sub(xf, "a", "chOff", x=0, y=0)
    sub(xf, "a", "chExt", cx=emu(width), cy=emu(height))
    dm_shape(group, 11, "DM_Background", 0, 0, width, height, color="FFFFFF")
    dm_shape(group, 12, "DM_Axis", 8, 56, width - 16, 12,
             geometry="rightArrow", color="8B0000")
    marker_top = 56 + 12 * (1 - 0.55) / 2 - 8
    gap = (width - 44) / len(EXAMPLES)
    for i, (label, date) in enumerate(EXAMPLES):
        center = 22 + gap * (i + 0.5)
        dm_shape(group, 13 + i * 3, f"DM_Marker_{i}", center - 5, marker_top,
                 10, 16, geometry="triangle", color="F2A000")
        dm_shape(group, 14 + i * 3, f"DM_Date_{i}", center - 26, marker_top - 15,
                 52, 20, text=date, bold=True)
        dm_shape(group, 15 + i * 3, f"DM_Label_{i}", center - (gap - 12) / 2,
                 80, gap - 12, 36, text=label, bold=True)
    sub(node, "xdr", "clientData")
    return drawing


def drawms_styles(styles):
    fonts = styles.find("s:fonts", NS)
    fills = styles.find("s:fills", NS)
    xfs = styles.find("s:cellXfs", NS)

    def font(size: int, bold: bool = False, color: str = "172033") -> int:
        idx = len(fonts)
        f = sub(fonts, "s", "font")
        if bold:
            sub(f, "s", "b")
        sub(f, "s", "sz", val=size)
        sub(f, "s", "color", rgb="FF" + color)
        sub(f, "s", "name", val="Microsoft YaHei")
        sub(f, "s", "charset", val=0)
        return idx

    def fill(color: str) -> int:
        idx = len(fills)
        pf = sub(sub(fills, "s", "fill"), "s", "patternFill", patternType="solid")
        sub(pf, "s", "fgColor", rgb="FF" + color)
        sub(pf, "s", "bgColor", indexed=64)
        return idx

    def style(font_id: int, fill_id: int = 0, num: int = 0,
              wrap: bool = False, align: str = "left") -> int:
        idx = len(xfs)
        xf = sub(xfs, "s", "xf", numFmtId=num, fontId=font_id, fillId=fill_id,
                 borderId=0, xfId=0, applyFont=1, applyFill=1,
                 applyAlignment=1, applyNumberFormat=1)
        sub(xf, "s", "alignment", horizontal=align, vertical="center",
            wrapText=int(wrap))
        return idx

    regular = font(11)
    title = font(20, True)
    muted = font(10, color="64748B")
    header = font(10, True, "FFFFFF")
    status = font(10, color="0F766E")
    blue = fill("EFF6FF")
    slate = fill("334155")
    teal = fill("F0FDFA")
    ids = {
        "title": style(title),
        "muted": style(muted, wrap=True),
        "header": style(header, slate, wrap=True, align="center"),
        # Built-in format 49 is @ (Text). It is applied to both input columns.
        "text": style(regular, blue, num=49, wrap=True),
        "status": style(status, teal, wrap=True),
    }
    for node in (fonts, fills, xfs):
        node.set("count", str(len(node)))
    fmts = styles.find("s:numFmts", NS)
    if fmts is not None:
        fmts.set("count", str(len(fmts)))
    return ids


def drawms_sheet(ids):
    ws = ET.Element(q("s", "worksheet"), nsmap={None: NS["s"], "r": NS["r"]})
    sub(sub(ws, "s", "sheetPr", codeName="Milestones"), "s", "tabColor", rgb="FFF2A000")
    sub(ws, "s", "dimension", ref="A1:C213")
    view = sub(sub(ws, "s", "sheetViews"), "s", "sheetView", showGridLines=0,
               tabSelected=1, zoomScale=80, workbookViewId=0)
    sub(view, "s", "pane", ySplit=13, topLeftCell="B14", activePane="bottomLeft",
        state="frozen")
    sub(view, "s", "selection", pane="bottomLeft", activeCell="B14", sqref="B14")
    sub(ws, "s", "sheetFormatPr", defaultRowHeight=20)
    cols = sub(ws, "s", "cols")
    for i, width in enumerate(DRAWMS_COLUMN_WIDTHS, 1):
        sub(cols, "s", "col", min=i, max=i, width=width, customWidth=1)
    data = sub(ws, "s", "sheetData")
    rows = {}

    def cell(row: int, col: str, value: str | None = None, style: str = "text"):
        if row not in rows:
            rows[row] = ET.Element(q("s", "row"), r=str(row), ht=str(32 if 14 <= row <= 19 else 20),
                                   customHeight="1")
        node = sub(rows[row], "s", "c", r=f"{col}{row}", s=ids[style])
        if value is not None:
            node.set("t", "inlineStr")
            text = sub(sub(node, "s", "is"), "s", "t")
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text.text = value

    rows[5] = ET.Element(q("s", "row"), r="5", ht="32", customHeight="1")
    rows[11] = ET.Element(q("s", "row"), r="11", ht="20", customHeight="1")
    rows[12] = ET.Element(q("s", "row"), r="12", ht="20", customHeight="1")
    cell(2, "B", "drawms", "title")
    rows[2].set("ht", "30")
    cell(3, "B", "Enter milestone name and date as text, then click Generate Chart.", "muted")
    cell(8, "B", "Rows with either value are drawn; blank rows are skipped. Maximum 60 milestones.", "muted")
    cell(10, "B", "6 sample milestones.", "status")
    rows[10].set("ht", "30")
    cell(13, "B", "Milestone name", "header")
    cell(13, "C", "Milestone date", "header")
    rows[13].set("ht", "28")
    for r in range(14, 214):
        value = EXAMPLES[r - 14] if r < 20 else None
        cell(r, "B", value[0] if value else None, "text")
        cell(r, "C", value[1] if value else None, "text")
    for r in sorted(rows):
        data.append(rows[r])
    merges = sub(ws, "s", "mergeCells", count=3)
    for r in (2, 3, 8, 10):
        sub(merges, "s", "mergeCell", ref=f"B{r}:C{r}")
    merges.set("count", str(len(merges)))
    sub(ws, "s", "pageMargins", left=.25, right=.25, top=.4, bottom=.4,
        header=.2, footer=.2)
    sub(ws, "s", "drawing", **{q("r", "id"): "rId1"})
    return ws


def remove_relationships(parts: dict[str, bytes], path: str, *, targets: set[str] = frozenset()):
    if path not in parts:
        return
    root = ET.fromstring(parts[path])
    for node in list(root):
        if node.get("Target") in targets or node.get("Type", "").rsplit("/", 1)[-1] in targets:
            root.remove(node)
    parts[path] = xml(root)


def build(source: Path, output: Path):
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}

    styles = ET.fromstring(parts["xl/styles.xml"])
    ids = drawms_styles(styles)
    parts["xl/styles.xml"] = xml(styles)

    # The template contributes the theme, styles and a clean VBA container;
    # every WBS/helper worksheet and its related drawing/control part is removed.
    legacy_parts = {
        "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml",
        "xl/calcChain.xml", "xl/sharedStrings.xml", "xl/drawings/vmlDrawing1.vml",
        "xl/ctrlProps/ctrlProp1.xml", "xl/printerSettings/printerSettings1.bin",
    }
    for name in legacy_parts:
        parts.pop(name, None)

    workbook = ET.fromstring(parts["xl/workbook.xml"])
    sheets = workbook.find("s:sheets", NS)
    for node in list(sheets):
        sheets.remove(node)
    sub(sheets, "s", "sheet", name="Milestones", sheetId=1, **{q("r", "id"): "rId1"})
    names = workbook.find("s:definedNames", NS)
    if names is not None:
        workbook.remove(names)
    for node in list(workbook):
        if ET.QName(node).localname == "AlternateContent":
            workbook.remove(node)
    book_view = workbook.find("s:bookViews/s:workbookView", NS)
    if book_view is not None:
        book_view.set("activeTab", "0")
    calc = workbook.find("s:calcPr", NS)
    if calc is not None:
        calc.set("fullCalcOnLoad", "1")
        calc.set("forceFullCalc", "1")
    parts["xl/workbook.xml"] = xml(workbook)

    workbook_rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    for node in list(workbook_rels):
        target = node.get("Target", "")
        if target in {"worksheets/sheet2.xml", "worksheets/sheet3.xml", "calcChain.xml", "sharedStrings.xml"}:
            workbook_rels.remove(node)
    parts["xl/_rels/workbook.xml.rels"] = xml(workbook_rels)

    types = ET.fromstring(parts["[Content_Types].xml"])
    for node in list(types):
        part_name = node.get("PartName", "").lstrip("/")
        if part_name in {name for name in legacy_parts} or part_name in {
            "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml",
        }:
            types.remove(node)
        if node.get("Extension") == "vml":
            types.remove(node)
        if node.get("Extension") == "bin" and node.get("ContentType", "").endswith("printerSettings"):
            types.remove(node)
    parts["[Content_Types].xml"] = xml(types)

    parts["xl/worksheets/sheet1.xml"] = xml(drawms_sheet(ids))
    parts["xl/drawings/drawing1.xml"] = xml(chart_group())
    sheet_rels = ET.Element(q("rel", "Relationships"), nsmap={None: NS["rel"]})
    sub(sheet_rels, "rel", "Relationship", Id="rId1",
        Type=NS["r"] + "/drawing", Target="../drawings/drawing1.xml")
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = xml(sheet_rels)

    # No customUI is added: native Excel Undo/Redo remains available.
    parts.pop("customUI/customUI.xml", None)
    root_rels = ET.fromstring(parts["_rels/.rels"])
    for node in list(root_rels):
        if node.get("Target", "").endswith("customUI/customUI.xml"):
            root_rels.remove(node)
    parts["_rels/.rels"] = xml(root_rels)

    app = ET.fromstring(parts["docProps/app.xml"])
    app_ns = {
        "p": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    }
    titles = app.find("p:TitlesOfParts/vt:vector", app_ns)
    if titles is not None:
        for node in list(titles):
            titles.remove(node)
        ET.SubElement(titles, "{" + app_ns["vt"] + "}lpstr").text = "Milestones"
        titles.set("size", "1")
    heading = app.find("p:HeadingPairs/vt:vector", app_ns)
    if heading is not None:
        for value in heading.findall("vt:variant/vt:i4", app_ns):
            value.text = "1"
    for node in app.iter():
        if node.text == "工作表":
            node.text = "Worksheets"
    parts["docProps/app.xml"] = xml(app)

    parts["xl/vbaProject.bin"] = rebuild_vba_modules(
        parts["xl/vbaProject.bin"],
        [("ThisWorkbook", ROOT / "src/DrawMs_ThisWorkbook.cls", True),
         ("Milestones", ROOT / "src/Milestones.cls", True),
         ("DrawMs", ROOT / "src/DrawMs.bas", False)],
    )

    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, data in sorted(parts.items()):
            archive.writestr(name, data)
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=ROOT / "WBS.xlsm")
    parser.add_argument("--output", type=Path, default=ROOT / "dist" / WORKBOOK_NAME)
    args = parser.parse_args()
    print(build(args.source, args.output))
