"""Build the English v5.1 WBS/Gantt workbook from the legacy template."""
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET

from vba_project import rebuild_vba
from upgrade_wbs import upgrade

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_NAME = "WBS-Gantt-ProgressChart-v5.1.xlsm"
NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
EXAMPLES = [("First samples\nprocessing", dt.date(2026, 8, 20)), ("Appearance / size\ninspection", dt.date(2026, 8, 22)),
            ("Probe electrical\nvalidation", dt.date(2026, 8, 25)), ("Wire bond\nvalidation", dt.date(2026, 8, 28)),
            ("Three-lot stability\nvalidation", dt.date(2026, 9, 5)), ("Production release", dt.date(2026, 9, 10))]


def q(prefix, name):
    return f"{{{NS[prefix]}}}{name}"


def sub(parent, prefix, tag, **attrs):
    return ET.SubElement(parent, q(prefix, tag), {k: str(v) for k, v in attrs.items()})


def xml(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True, standalone=True)


def emu(points):
    return str(round(points * 12700))


def transform(parent, x, y, width, height):
    xf = sub(parent, "a", "xfrm")
    sub(xf, "a", "off", x=emu(x), y=emu(y))
    sub(xf, "a", "ext", cx=emu(width), cy=emu(height))
    return xf


def text_body(parent, text, bold=False, color="000000", size=1000, center=True):
    body = sub(parent, "xdr", "txBody")
    sub(body, "a", "bodyPr", wrap="square", lIns=0, rIns=0, tIns=0, bIns=0, anchor="t")
    sub(body, "a", "lstStyle")
    for line in text.split("\n"):
        p = sub(body, "a", "p")
        sub(p, "a", "pPr", algn="ctr" if center else "l")
        run = sub(p, "a", "r")
        props = sub(run, "a", "rPr", lang="en-US", sz=size, b=int(bold))
        sub(sub(props, "a", "solidFill"), "a", "srgbClr", val=color)
        sub(props, "a", "latin", typeface="Arial")
        sub(props, "a", "ea", typeface="Arial")
        sub(run, "a", "t").text = line
        sub(p, "a", "endParaRPr", lang="en-US", sz=size)


def shape(parent, idx, name, x, y, width, height, geometry="rect", color=None,
          text=None, bold=False, macro=None, text_color="000000"):
    attrs = {"macro": "[0]!ProgressChart." + macro} if macro else {}
    sp = sub(parent, "xdr", "sp", **attrs)
    nv = sub(sp, "xdr", "nvSpPr")
    sub(nv, "xdr", "cNvPr", id=idx, name=name)
    sub(nv, "xdr", "cNvSpPr", **({"txBox": "1"} if text is not None and color is None else {}))
    props = sub(sp, "xdr", "spPr")
    transform(props, x, y, width, height)
    geom = sub(props, "a", "prstGeom", prst=geometry)
    values = sub(geom, "a", "avLst")
    if geometry == "rightArrow":
        sub(values, "a", "gd", name="adj1", fmla="val 55000")
        sub(values, "a", "gd", name="adj2", fmla="val 60000")
    if color:
        sub(sub(props, "a", "solidFill"), "a", "srgbClr", val=color)
    else:
        sub(props, "a", "noFill")
    sub(sub(props, "a", "ln"), "a", "noFill")
    if text is not None:
        text_body(sp, text, bold=bold, color=text_color)
        if color:
            body_pr = sp.find("xdr:txBody/a:bodyPr", NS)
            body_pr.set("anchor", "ctr")
    return sp


def anchor(parent, x, y, width, height):
    a = sub(parent, "xdr", "absoluteAnchor")
    sub(a, "xdr", "pos", x=emu(x), y=emu(y))
    sub(a, "xdr", "ext", cx=emu(width), cy=emu(height))
    return a


def button(parent, idx, label, x, y, width, macro, color="2563EB"):
    a = anchor(parent, x, y, width, 25)
    shape(a, idx, "PG_Button_" + macro, x, y, width, 25, geometry="roundRect",
          color=color, text=label, bold=True, macro=macro, text_color="FFFFFF")
    sub(a, "xdr", "clientData", fPrintsWithSheet="0")


def progress_drawing():
    drawing = ET.Element(q("xdr", "wsDr"), nsmap={"xdr": NS["xdr"], "a": NS["a"]})
    # Row 5 is reserved for these controls so they do not cover the C6/C7
    # position and spacing settings.  Coordinates are points.
    button(drawing, 2, "Generate / Refresh Chart", 25, 93, 171, "GenerateProgressChart")
    button(drawing, 3, "Copy Chart", 208, 93, 114, "CopyProgressChart", "0F766E")
    button(drawing, 4, "Widen Spacing", 335, 93, 125, "WidenProgressChart", "7C3AED")
    button(drawing, 5, "Narrow Spacing", 472, 93, 125, "NarrowProgressChart", "64748B")
    # L3 matches the VBA anchor and keeps the chart clear of the B:J input
    # table.  The chart can grow beyond this width when more nodes are shown.
    x, y, width, height = 921, 50, 690, 126
    a = anchor(drawing, x, y, width, height)
    group = sub(a, "xdr", "grpSp")
    nv = sub(group, "xdr", "nvGrpSpPr")
    sub(nv, "xdr", "cNvPr", id=10, name="WBS_ProgressChart", descr="Milestones sample: six nodes sorted by date with equal spacing.")
    sub(nv, "xdr", "cNvGrpSpPr")
    xf = transform(sub(group, "xdr", "grpSpPr"), x, y, width, height)
    sub(xf, "a", "chOff", x=0, y=0)
    sub(xf, "a", "chExt", cx=emu(width), cy=emu(height))
    shape(group, 11, "PG_Background", 0, 0, width, height, color="FFFFFF")
    shape(group, 12, "PG_Axis", 8, 56, width - 16, 12, geometry="rightArrow", color="8B0000")
    marker_top = 56 + 12 * (1 - .55) / 2 - 8
    gap = (width - 44) / 6
    for i, (label, date) in enumerate(EXAMPLES):
        center = 22 + gap * (i + .5)
        shape(group, 13 + i * 3, f"PG_Marker_{i}", center - 5, marker_top, 10, 16, geometry="triangle", color="F2A000")
        shape(group, 14 + i * 3, f"PG_Date_{i}", center - 26, marker_top - 15, 52, 13, text=f"{date.month}/{date.day}", bold=True)
        shape(group, 15 + i * 3, f"PG_Label_{i}", center - (gap - 12) / 2, 80, gap - 12, 36, text=label, bold=True)
    sub(a, "xdr", "clientData")
    return drawing


def add_styles(styles):
    fonts = styles.find("s:fonts", NS)
    fills = styles.find("s:fills", NS)
    xfs = styles.find("s:cellXfs", NS)

    def font(size, bold=False, color="172033"):
        idx = len(fonts)
        f = sub(fonts, "s", "font")
        if bold:
            sub(f, "s", "b")
        sub(f, "s", "sz", val=size)
        sub(f, "s", "color", rgb="FF" + color)
        sub(f, "s", "name", val="Microsoft YaHei")
        sub(f, "s", "charset", val=0)
        return idx

    def fill(color):
        idx = len(fills)
        pf = sub(sub(fills, "s", "fill"), "s", "patternFill", patternType="solid")
        sub(pf, "s", "fgColor", rgb="FF" + color)
        sub(pf, "s", "bgColor", indexed=64)
        return idx

    def style(font_id, fill_id=0, num=0, wrap=False, align="left"):
        idx = len(xfs)
        xf = sub(xfs, "s", "xf", numFmtId=num, fontId=font_id, fillId=fill_id, borderId=0, xfId=0,
                 applyFont=1, applyFill=1, applyAlignment=1, applyNumberFormat=1)
        sub(xf, "s", "alignment", horizontal=align, vertical="center", wrapText=int(wrap))
        return idx

    regular, title, light, header = font(10), font(20, True), font(10, color="64748B"), font(10, True, "FFFFFF")
    date_id = 181
    fmts = styles.find("s:numFmts", NS)
    date_id = max(date_id, max(int(x.get("numFmtId")) for x in fmts) + 1)
    sub(fmts, "s", "numFmt", numFmtId=date_id, formatCode="yyyy/m/d")
    ids = {"regular": style(regular), "title": style(title), "muted": style(light, wrap=True),
           "header": style(header, fill("334155"), wrap=True, align="center"),
           "input": style(regular, fill("EFF6FF"), wrap=True),
           "center": style(regular, fill("EFF6FF"), align="center"),
           "percent": style(regular, fill("EFF6FF"), num=9, align="center"),
           "date": style(regular, fill("EFF6FF"), num=date_id),
           "status": style(font(10, color="0F766E"), fill("F0FDFA"), wrap=True)}
    for node in [fonts, fills, xfs, fmts]:
        node.set("count", str(len(node)))
    return ids


def progress_sheet(ids):
    ws = ET.Element(q("s", "worksheet"), nsmap={None: NS["s"], "r": NS["r"]})
    sub(sub(ws, "s", "sheetPr", codeName="Sheet4"), "s", "tabColor", rgb="FFF2A000")
    sub(ws, "s", "dimension", ref="A1:T213")
    view = sub(sub(ws, "s", "sheetViews"), "s", "sheetView", showGridLines=0, tabSelected=0, zoomScale=80, workbookViewId=0)
    sub(view, "s", "pane", ySplit=13, topLeftCell="A14", activePane="bottomLeft", state="frozen")
    sub(view, "s", "selection", pane="bottomLeft", activeCell="D14", sqref="D14")
    sub(ws, "s", "sheetFormatPr", defaultRowHeight=20)
    cols = sub(ws, "s", "cols")
    for i, width in enumerate([3.5, 12, 23, 29, 17, 30, 18, 10, 12, 10, 4], 1):
        sub(cols, "s", "col", min=i, max=i, width=width, customWidth=1)
    sub(cols, "s", "col", min=12, max=30, width=12, customWidth=1)
    data = sub(ws, "s", "sheetData")
    rows = {}

    def cell(row, col, value=None, style="regular"):
        if row not in rows:
            rows[row] = ET.Element(q("s", "row"), r=str(row), ht=str(32 if 14 <= row <= 19 else 20), customHeight="1")
        node = sub(rows[row], "s", "c", r=f"{col}{row}", s=ids[style])
        if isinstance(value, dt.date):
            sub(node, "s", "v").text = str((value - dt.date(1899, 12, 30)).days)
        elif isinstance(value, (int, float)):
            sub(node, "s", "v").text = str(value)
        elif value is not None:
            node.set("t", "inlineStr")
            t = sub(sub(node, "s", "is"), "s", "t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = value

    # Keep a blank, taller row for the drawing buttons.  Without an explicit
    # row 5, Excel collapses the reserved band and the buttons overlap C6/C7.
    rows[5] = ET.Element(q("s", "row"), r="5", ht="32", customHeight="1")
    cell(2, "B", "Milestones", "title")
    rows[2].set("ht", "30")
    cell(3, "B", "Enter milestones or import WBS tasks, then Generate.", "muted")
    cell(6, "B", "Position", "muted")
    cell(6, "C", "Equal spacing", "center")
    cell(7, "B", "Spacing", "muted")
    cell(7, "C", 1, "percent")
    cell(8, "B", "Show = Yes to include. Clear samples before import.", "muted")
    cell(9, "B", "Dates: yyyy/m/d, yyyymmdd, or mmdd (year from row above).", "muted")
    cell(10, "B", "6 sample milestones.", "status")
    rows[10].set("ht", "30")
    for col, name in zip("BCDEFGHIJ", ["Show", "WBS row", "Milestone name", "Milestone date", "Notes / source", "Owner", "Progress", "Status", "Type"]):
        cell(13, col, name, "header")
    rows[13].set("ht", "28")
    for r in range(14, 214):
        value = EXAMPLES[r - 14] if r < 20 else None
        cell(r, "B", "Yes" if value else None, "center")
        cell(r, "C", None, "center")
        cell(r, "D", value[0] if value else None, "input")
        cell(r, "E", value[1] if value else None, "date")
        cell(r, "F", None, "input")
        cell(r, "G", None, "input")
        cell(r, "H", None, "center")
        cell(r, "I", None, "status")
        cell(r, "J", None, "center")
    for r in sorted(rows):
        data.append(rows[r])
    merges = sub(ws, "s", "mergeCells", count=9)
    for r in [2, 3, 4, 8, 9, 10, 11, 12]:
        sub(merges, "s", "mergeCell", ref=f"B{r}:J{r}")
    merges.set("count", str(len(merges)))
    dvs = sub(ws, "s", "dataValidations", count=3)
    dv = sub(dvs, "s", "dataValidation", type="list", errorStyle="stop", allowBlank=1,
             showErrorMessage=1, showInputMessage=1, sqref="B14:B213", errorTitle="Show option", error="Choose Yes or No.")
    sub(dv, "s", "formula1").text = '"Yes,No"'
    dv = sub(dvs, "s", "dataValidation", type="list", errorStyle="stop", allowBlank=0,
             showErrorMessage=1, showInputMessage=1, sqref="C6", errorTitle="Node position", error="Choose Equal spacing or Date proportional.")
    sub(dv, "s", "formula1").text = '"Equal spacing,Date proportional"'
    dv = sub(dvs, "s", "dataValidation", type="decimal", errorStyle="stop", allowBlank=0,
             showErrorMessage=1, showInputMessage=1, sqref="C7", errorTitle="Node spacing", error="Enter a value from 50% to 200%.")
    sub(dv, "s", "formula1").text = "0.5"
    sub(dv, "s", "formula2").text = "2"
    sub(ws, "s", "pageMargins", left=.25, right=.25, top=.4, bottom=.4, header=.2, footer=.2)
    sub(ws, "s", "drawing", **{q("r", "id"): "rId1"})
    return ws


def _remove_sheet(parts: dict[str, bytes], sheet_name: str, rel_id: str, sheet_path: str):
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    sheets = workbook.find("s:sheets", NS)
    for node in list(sheets):
        if node.get("name") == sheet_name or node.get(q("r", "id")) == rel_id:
            sheets.remove(node)
    parts["xl/workbook.xml"] = xml(workbook)
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    for node in list(rels):
        if node.get("Id") == rel_id or node.get("Target") == sheet_path:
            rels.remove(node)
    parts["xl/_rels/workbook.xml.rels"] = xml(rels)
    types = ET.fromstring(parts["[Content_Types].xml"])
    for node in list(types):
        if node.get("PartName") == "/xl/" + sheet_path:
            types.remove(node)
    parts["[Content_Types].xml"] = xml(types)
    parts.pop("xl/" + sheet_path, None)


def build(source: Path, output: Path):
    with ZipFile(source) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts["xl/vbaProject.bin"] = rebuild_vba(
        parts["xl/vbaProject.bin"], ROOT / "src/ProgressChart.bas",
        {"Model_1": ROOT / "src/Model_1.bas", "Sheet1": ROOT / "src/Sheet1.cls"},
    )
    styles = ET.fromstring(parts["xl/styles.xml"])
    ids = add_styles(styles)
    parts["xl/styles.xml"] = xml(styles)
    # Remove the legacy helper sheets before adding the v5 Milestones sheet.
    _remove_sheet(parts, "Data", "rId2", "worksheets/sheet2.xml")
    _remove_sheet(parts, "Resource", "rId3", "worksheets/sheet3.xml")
    parts.pop("xl/calcChain.xml", None)
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    for node in list(rels):
        if node.get("Target") == "calcChain.xml": rels.remove(node)
    parts["xl/_rels/workbook.xml.rels"] = xml(rels)
    types = ET.fromstring(parts["[Content_Types].xml"])
    for node in list(types):
        if node.get("PartName") == "/xl/calcChain.xml": types.remove(node)
    parts["[Content_Types].xml"] = xml(types)
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    sheets = workbook.find("s:sheets", NS)
    sub(sheets, "s", "sheet", name="Milestones", sheetId=4, **{q("r", "id"): "rId9"})
    workbook.find("s:bookViews/s:workbookView", NS).set("activeTab", "0")
    names = workbook.find("s:definedNames", NS)
    if names is not None: workbook.remove(names)
    for node in list(workbook):
        if ET.QName(node).localname == "AlternateContent": workbook.remove(node)
    calc = workbook.find("s:calcPr", NS)
    calc.set("fullCalcOnLoad", "1"); calc.set("forceFullCalc", "1")
    parts["xl/workbook.xml"] = xml(workbook)
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    sub(rels, "rel", "Relationship", Id="rId9", Type=NS["r"] + "/worksheet", Target="worksheets/sheet4.xml")
    parts["xl/_rels/workbook.xml.rels"] = xml(rels)
    types = ET.fromstring(parts["[Content_Types].xml"])
    for name, content_type in [("/xl/worksheets/sheet4.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
                               ("/xl/drawings/drawing2.xml", "application/vnd.openxmlformats-officedocument.drawing+xml")]:
        sub(types, "ct", "Override", PartName=name, ContentType=content_type)
    parts["[Content_Types].xml"] = xml(types)
    wbs = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    wbs.find("s:sheetViews/s:sheetView", NS).set("tabSelected", "0")
    # Persist this setting so typing a task without display spaces works on
    # the first open, including before macros are enabled and via Excel COM.
    # Keep the list and its dropdown; only disable the blocking error alert.
    predecessor = wbs.find("s:dataValidations/s:dataValidation[@sqref='O6:O64']", NS)
    if predecessor is None or predecessor.get("type") != "list":
        raise ValueError("The WBS template must contain the O6:O64 predecessor list validation")
    predecessor.set("showErrorMessage", "0")
    validations = wbs.find("s:dataValidations", NS)
    level_dv = sub(validations, "s", "dataValidation", type="list", allowBlank="0", showErrorMessage="1", sqref="D2", errorTitle="Task level", error="Choose a level from 1 to 5.")
    sub(level_dv, "s", "formula1").text = '"1,2,3,4,5"'
    validations.set("count", str(len(validations)))
    # Translate all visible shared-string cells on WBS. The source template is
    # intentionally kept as the data fixture; only the generated v5 copy is
    # translated so the original workbook remains available for comparison.
    strings = ET.fromstring(parts["xl/sharedStrings.xml"])
    translations = {
        "天数": "Days", "工数：H": "Hours", "任务": "Task", "资源": "Resource", "投入百分比": "Allocation %", "成果物单位": "Deliverable unit",
        "成果物量": "Deliverable qty", "完成成果物量": "Completed qty", "作业进度": "Progress",
        "前置任务": "Predecessor", "计划开始时间": "Planned start", "计划结束时间": "Planned end",
        "实际开始时间": "Actual start", "实际结束时间": "Actual end", "基线开始时间": "Baseline start",
        "基线结束时间": "Baseline end", "任务类型": "Task type", "级别": "Level", "完成": "Complete",
        "未完成": "Not started", "空白": "Blank", "进行中": "In progress", "星期六": "Sat", "星期日": "Sun",
        "开发平台概述": "Platform overview", "新增功能": "New features", "前世今生": "Background",
        "特点": "Characteristics", "直观方便的功能区": "Intuitive interface", "开发应用程序有啥好处": "Benefits of application development",
        "应用程序结构": "Application structure", "程序有哪几部分组成": "Program components", "面向对象编程是什么": "Object-oriented programming",
        "应用程序开发流程": "Application development process", "开发前准备": "Preparation", "开发过程": "Development",
        "测试": "Testing", "发布给最终用户": "Release to users", "使用宏": "Using macros", "宏简介": "Macro overview",
        "宏是什么": "What is a macro", "宏有哪些特点": "Macro characteristics", "有哪些方法创建宏": "Ways to create macros",
        "将你的宏记录下来": "Record your macro", "页数": "Pages", "行数": "Rows", "项目数": "Items",
        "任务类型": "Task type", "资源": "Resource", "节假日": "Holiday", "周末上班": "Weekend work",
    }
    # Translate task strings while preserving numeric prefixes and indentation.
    for si in strings.findall("s:si", NS):
        text_nodes = si.findall(".//s:t", NS)
        if not text_nodes: continue
        value = "".join((node.text or "") for node in text_nodes)
        leading = value[:len(value) - len(value.lstrip(" "))]
        body = value[len(leading):]
        prefix = ""
        if " " in body and body.split(" ", 1)[0].replace(".", "").isdigit():
            prefix, body = body.split(" ", 1)
        for cn, en in sorted(translations.items(), key=lambda item: len(item[0]), reverse=True): body = body.replace(cn, en)
        new_value = leading + ((prefix + " ") if prefix else "") + body
        if new_value != value:
            text_nodes[0].text = new_value
            for node in text_nodes[1:]: node.text = ""
    parts["xl/sharedStrings.xml"] = xml(strings)
    upgrade(wbs, styles, strings)
    parts["xl/styles.xml"] = xml(styles)
    parts["xl/worksheets/sheet1.xml"] = xml(wbs)
    drawing = ET.Element(q("xdr", "wsDr"), nsmap={"xdr": NS["xdr"], "a": NS["a"]})
    button(drawing, 2000, "Update Gantt", 25, 55, 108, "UpdateGantt")
    button(drawing, 2001, "Open Milestones", 141, 55, 124, "OpenProgressChart")
    button(drawing, 2002, "Import Selected Tasks", 273, 55, 155, "ImportSelectedMilestones", "0F766E")
    button(drawing, 2003, "Refresh All", 436, 55, 94, "RefreshAll", "7C3AED")
    button(drawing, 2004, "Add Task", 538, 55, 85, "AddTask", "2563EB")
    button(drawing, 2005, "Add Subtask", 631, 55, 104, "AddSubtask", "0F766E")
    button(drawing, 2006, "Promote (-1)", 743, 55, 103, "PromoteTasks", "475569")
    button(drawing, 2007, "Demote (+1)", 854, 55, 103, "DemoteTasks", "475569")
    parts["xl/drawings/drawing1.xml"] = xml(drawing)
    # Replace the old form control with ordinary DrawingML macro buttons.
    obsolete = {"xl/drawings/vmlDrawing1.vml", "xl/ctrlProps/ctrlProp1.xml"}
    for name in obsolete: parts.pop(name, None)
    sheet_rels = ET.fromstring(parts["xl/worksheets/_rels/sheet1.xml.rels"])
    for node in list(sheet_rels):
        if node.get("Type").rsplit("/", 1)[-1] in {"vmlDrawing", "ctrlProp"}: sheet_rels.remove(node)
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = xml(sheet_rels)
    types = ET.fromstring(parts["[Content_Types].xml"])
    for node in list(types):
        if node.get("PartName", "").lstrip("/") in obsolete: types.remove(node)
    parts["[Content_Types].xml"] = xml(types)
    parts["xl/drawings/drawing2.xml"] = xml(progress_drawing())
    parts["xl/worksheets/sheet4.xml"] = xml(progress_sheet(ids))
    sheet_rels = ET.Element(q("rel", "Relationships"), nsmap={None: NS["rel"]})
    sub(sheet_rels, "rel", "Relationship", Id="rId1", Type=NS["r"] + "/drawing", Target="../drawings/drawing2.xml")
    parts["xl/worksheets/_rels/sheet4.xml.rels"] = xml(sheet_rels)
    # Update document metadata for the two-sheet release.
    app = ET.fromstring(parts["docProps/app.xml"])
    app_ns = {"p": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
              "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"}
    titles = app.find("p:TitlesOfParts/vt:vector", app_ns)
    if titles is not None:
        for node in list(titles): titles.remove(node)
        for name in ["WBS", "Milestones"]:
            ET.SubElement(titles, "{" + app_ns["vt"] + "}lpstr").text = name
        titles.set("size", "2")
    heading = app.find("p:HeadingPairs/vt:vector", app_ns)
    if heading is not None:
        for value in heading.findall("vt:variant/vt:i4", app_ns):
            value.text = "2"
    for node in app.iter():
        if node.text == "工作表": node.text = "Worksheets"
    parts["docProps/app.xml"] = xml(app)
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, "w", ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=ROOT / "WBS.xlsm")
    parser.add_argument("--output", type=Path, default=ROOT / "dist" / WORKBOOK_NAME)
    args = parser.parse_args()
    print(build(args.source, args.output))
