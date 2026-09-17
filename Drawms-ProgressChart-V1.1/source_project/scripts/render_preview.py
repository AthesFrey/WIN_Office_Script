"""Render the actual saved DrawingML group to a PNG for reviewing the layout.

This is a limited renderer for the shapes used here, not an Excel screenshot.
"""
from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET
from PIL import Image, ImageDraw, ImageFont

from build_workbook import ROOT, NS, WORKBOOK_NAME


def render(workbook: Path, font_path: Path, output: Path):
    with ZipFile(workbook) as z:
        drawing_name = "xl/drawings/drawing2.xml" if "xl/drawings/drawing2.xml" in z.namelist() else "xl/drawings/drawing1.xml"
        drawing = ET.fromstring(z.read(drawing_name))
    group = drawing.find("xdr:absoluteAnchor/xdr:grpSp", NS)
    extent = group.find("xdr:grpSpPr/a:xfrm/a:chExt", NS)
    scale = 2

    def points(value):
        return int(value) / 12700 * scale

    im = Image.new("RGB", (round(points(extent.get("cx"))), round(points(extent.get("cy")))), "white")
    draw = ImageDraw.Draw(im)
    for sp in group.findall("xdr:sp", NS):
        props = sp.find("xdr:spPr", NS)
        offset, size = props.find("a:xfrm/a:off", NS), props.find("a:xfrm/a:ext", NS)
        x, y = points(offset.get("x")), points(offset.get("y"))
        w, h = points(size.get("cx")), points(size.get("cy"))
        geometry = props.find("a:prstGeom", NS).get("prst")
        color = props.find("a:solidFill/a:srgbClr", NS)
        if color is not None:
            color = "#" + color.get("val")
            if geometry == "rect":
                draw.rectangle((x, y, x + w, y + h), fill=color)
            elif geometry == "triangle":
                draw.polygon([(x + w / 2, y), (x + w, y + h), (x, y + h)], fill=color)
            elif geometry == "rightArrow":
                head = h * .6
                inset = h * (1 - .55) / 2
                draw.polygon([(x, y + inset), (x + w - head, y + inset), (x + w - head, y),
                              (x + w, y + h / 2), (x + w - head, y + h),
                              (x + w - head, y + h - inset), (x, y + h - inset)], fill=color)
        for i, p in enumerate(sp.findall("xdr:txBody/a:p", NS)):
            value = "".join(p.itertext())
            rpr = p.find("a:r/a:rPr", NS)
            font_size = int(rpr.get("sz")) / 100 * scale
            font = ImageFont.truetype(str(font_path), round(font_size))
            bold = rpr.get("b") == "1"
            draw.text((x + w / 2, y + i * 15 * scale), value, font=font, fill="black", anchor="mt",
                      stroke_width=0 if not bold else 0.3)
    im.save(output)


def render_toolbar(workbook: Path, font_path: Path, output: Path):
    """Review top-level shape transforms, independently of the anchor markers."""
    with ZipFile(workbook) as z:
        drawing_name = "xl/drawings/drawing2.xml" if "xl/drawings/drawing2.xml" in z.namelist() else "xl/drawings/drawing1.xml"
        sheet_name = "xl/worksheets/sheet4.xml" if "xl/worksheets/sheet4.xml" in z.namelist() else "xl/worksheets/sheet1.xml"
        drawing = ET.fromstring(z.read(drawing_name))
        sheet = ET.fromstring(z.read(sheet_name))
    standalone = sheet_name.endswith("sheet1.xml") and sheet.find("s:sheetPr", NS).get("codeName") == "Milestones"
    scale = 2
    im = Image.new("RGB", (800 * scale, 310 * scale), "white")
    draw = ImageDraw.Draw(im)
    font = ImageFont.truetype(str(font_path), 10 * scale)
    title = ImageFont.truetype(str(font_path), 20 * scale)
    heights = {int(row.get("r")): float(row.get("ht", "20"))
               for row in sheet.findall("s:sheetData/s:row", NS)}
    row_top = lambda row: sum(heights.get(r, 20) for r in range(1, row))
    rows_to_draw = (2, 3, 8, 10, 13) if standalone else (2, 3, 6, 7, 8, 9, 10, 13)
    for row in rows_to_draw:
        top = row_top(row) * scale
        row_height = heights.get(row, 20)
        fill = "#334155" if row == 13 else "#F0FDFA" if row == 10 else "#F8FAFC"
        draw.rectangle((25 * scale, top, 775 * scale, top + row_height * scale), fill=fill)
        cells = sheet.findall(f's:sheetData/s:row[@r="{row}"]/s:c', NS)
        labels = ["".join(cell.findall("s:is/s:t", NS)[0].itertext())
                  for cell in cells if cell.find("s:is/s:t", NS) is not None]
        text = " | ".join(labels)
        if row == 7 and not standalone:
            text += " | 100%"
        draw.text((27 * scale, top), text, font=title if row == 2 else font,
                  fill="white" if row == 13 else "#334155")
    for sp in drawing.findall("xdr:absoluteAnchor/xdr:sp", NS):
        props = sp.find("xdr:spPr", NS)
        off, ext = props.find("a:xfrm/a:off", NS), props.find("a:xfrm/a:ext", NS)
        x, y = [int(off.get(k)) / 12700 * scale for k in ("x", "y")]
        w, h = [int(ext.get(k)) / 12700 * scale for k in ("cx", "cy")]
        color = "#" + props.find("a:solidFill/a:srgbClr", NS).get("val")
        draw.rounded_rectangle((x, y, x + w, y + h), radius=4 * scale, fill=color)
        label = " ".join(sp.findall("xdr:txBody/a:p/a:r/a:t", NS)[0].itertext())
        draw.text((x + w / 2, y + h / 2), label, font=font, fill="white", anchor="mm")
    im.save(output)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--font", type=Path, required=True)
    parser.add_argument("--workbook", type=Path, default=ROOT / "dist" / WORKBOOK_NAME)
    parser.add_argument("--toolbar", action="store_true", help="Render saved button transforms and reserved rows")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.output is None:
        args.output = ROOT / "dist" / ("milestones_toolbar_preview.png" if args.toolbar else "progress_chart_preview.png")
    renderer = render_toolbar if args.toolbar else render
    renderer(args.workbook, args.font, args.output)
    print(args.output)
