"""Validate added worksheet/drawing against the ECMA-376 schemas.

Pass a local checkout of https://github.com/t-yuki/ooxml-xsd. The in-memory
resolver combines imports for libxml2, which loads only one per namespace.
"""
import argparse
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET

XS = "http://www.w3.org/2001/XMLSchema"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def check(workbook: Path, schema_dir: Path):
    combined = ET.Element(f"{{{XS}}}schema", nsmap={"xsd": XS}, targetNamespace=A, elementFormDefault="qualified")
    for path in sorted(schema_dir.glob("dml-*.xsd")):
        if path.name == "dml-all.xsd":
            continue
        root = ET.parse(str(path)).getroot()
        if root.get("targetNamespace") == A:
            ET.SubElement(combined, f"{{{XS}}}include", schemaLocation=path.name)

    class Resolver(ET.Resolver):
        def resolve(self, url, public_id, context):
            if url.endswith("combined-dml.xsd"):
                return self.resolve_string(ET.tostring(combined), context, base_url=str(schema_dir / "combined-dml.xsd"))

    parser = ET.XMLParser()
    parser.resolvers.add(Resolver())
    drawing = ET.parse(str(schema_dir / "dml-spreadsheetDrawing.xsd"), parser)
    for node in list(drawing.getroot()):
        if node.tag == f"{{{XS}}}import" and node.get("namespace") == A:
            drawing.getroot().remove(node)
    drawing.getroot().insert(0, ET.Element(f"{{{XS}}}import", namespace=A, schemaLocation="combined-dml.xsd"))
    drawing_schema = ET.XMLSchema(drawing)
    sheet_schema = ET.XMLSchema(ET.parse(str(schema_dir / "sml-sheet.xsd")))
    with ZipFile(workbook) as z:
        for name in ["xl/drawings/drawing1.xml", "xl/drawings/drawing2.xml"]:
            drawing_schema.assertValid(ET.fromstring(z.read(name)))
        for name in ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet4.xml"]:
            sheet = ET.fromstring(z.read(name))
            # The 2006 schema omits later Office extension attributes.
            for node in sheet.iter():
                for attr in list(node.attrib):
                    if attr.startswith("{") and not attr.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"):
                        del node.attrib[attr]
            sheet_schema.assertValid(sheet)
    print("Both worksheet and DrawingML schema validations passed")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--schema-dir", type=Path, required=True)
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    check(args.workbook, args.schema_dir)
