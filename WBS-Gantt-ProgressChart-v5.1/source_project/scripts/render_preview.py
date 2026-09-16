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
        drawing = ET.fromstring(z.read("xl/drawings/drawing2.xml"))
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


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--font", type=Path, required=True)
    parser.add_argument("--workbook", type=Path, default=ROOT / "dist" / WORKBOOK_NAME)
    parser.add_argument("--output", type=Path, default=ROOT / "dist/progress_chart_preview.png")
    args = parser.parse_args()
    render(args.workbook, args.font, args.output)
    print(args.output)
