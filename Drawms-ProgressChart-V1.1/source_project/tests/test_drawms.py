"""Static regression tests for the standalone drawms workbook."""
from __future__ import annotations

import io
import re
from pathlib import Path
from zipfile import ZipFile

import olefile
import pytest
from lxml import etree as ET
from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "scripts"))
from build_drawms import EXAMPLES, NS, WORKBOOK_NAME, build


@pytest.fixture(scope="module")
def artifact(tmp_path_factory):
    return build(ROOT / "WBS.xlsm", tmp_path_factory.mktemp("drawms") / WORKBOOK_NAME)


@pytest.fixture(scope="module")
def parts(artifact):
    with ZipFile(artifact) as archive:
        assert archive.testzip() is None
        return {name: archive.read(name) for name in archive.namelist()}


def text_of(node) -> str:
    return "".join(node.itertext())


def test_single_sheet_and_no_legacy_parts(parts):
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    sheets = workbook.findall("s:sheets/s:sheet", NS)
    assert len(sheets) == 1
    assert sheets[0].get("name") == "Milestones"
    assert set(parts).isdisjoint({
        "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml",
        "xl/drawings/drawing2.xml", "xl/sharedStrings.xml",
        "xl/calcChain.xml", "customUI/customUI.xml",
        "xl/drawings/vmlDrawing1.vml", "xl/ctrlProps/ctrlProp1.xml",
    })
    project = olefile.OleFileIO(io.BytesIO(parts["xl/vbaProject.bin"]))
    metadata = project.openstream("PROJECT").read().decode("cp936")
    for stale in ("Sheet1", "Sheet4", "Model_1", "ProgressChart", "Undo", "Redo", "WBS"):
        assert stale not in metadata


def test_input_cells_are_literal_text_without_constraints(parts):
    sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    assert sheet.find("s:dataValidations", NS) is None
    assert not sheet.findall(".//s:f", NS)
    styles = ET.fromstring(parts["xl/styles.xml"])
    xfs = styles.findall("s:cellXfs/s:xf", NS)
    for cell in sheet.findall("s:sheetData/s:row/s:c", NS):
        ref = cell.get("r")
        if ref and re.fullmatch(r"[BC](?:1[4-9]|[2-9][0-9]|1[0-9]{2}|20[0-9]|21[0-3])", ref):
            if cell.find("s:is", NS) is None:
                # Blank prepared input cells have no value node but retain
                # the Text-format style below.
                assert cell.find("s:v", NS) is None
                assert int(xfs[int(cell.get("s"))].get("numFmtId")) == 49
                continue
            assert cell.get("t") == "inlineStr"
            assert cell.find("s:f", NS) is None
            assert cell.find("s:v", NS) is None
            assert int(xfs[int(cell.get("s"))].get("numFmtId")) == 49
    rows = {int(row.get("r")): row for row in sheet.findall("s:sheetData/s:row", NS)}
    for offset, (name, date) in enumerate(EXAMPLES, 14):
        cells = {cell.get("r"): text_of(cell.find("s:is", NS)) for cell in rows[offset].findall("s:c", NS)}
        assert cells[f"B{offset}"] == name
        assert cells[f"C{offset}"] == date


def test_three_buttons_and_fix3_style_chart(parts):
    drawing = ET.fromstring(parts["xl/drawings/drawing1.xml"])
    anchors = drawing.findall("xdr:absoluteAnchor", NS)
    assert len(anchors) == 4
    buttons = drawing.findall("xdr:absoluteAnchor/xdr:sp", NS)
    assert len(buttons) == 3
    macros = {shape.get("macro") for shape in buttons}
    assert macros == {"[0]!DrawMs.GenerateChart", "[0]!DrawMs.WidenSpacing", "[0]!DrawMs.NarrowSpacing"}
    chart = drawing.find("xdr:absoluteAnchor/xdr:grpSp", NS)
    assert chart is not None
    assert chart.find("xdr:nvGrpSpPr/xdr:cNvPr", NS).get("name") == "DrawMs_Chart"
    children = chart.findall("xdr:sp", NS)
    assert len(children) == 20
    text_values = ["\n".join(text_of(para) for para in shape.find("xdr:txBody", NS).findall("a:p", NS))
                   for shape in children if shape.find("xdr:txBody", NS) is not None]
    assert len(text_values) == 12
    assert [date for _, date in EXAMPLES] == [value for value in text_values if value.startswith("2026/")]
    assert "First samples\nprocessing" in text_values
    # Controls remain fixed absolute worksheet coordinates and retain 25 pt height.
    expected = {"DM_Button_GenerateChart": (25, 93, 171),
                "DM_Button_WidenSpacing": (335, 93, 125),
                "DM_Button_NarrowSpacing": (472, 93, 125)}
    for shape in buttons:
        name = shape.find("xdr:nvSpPr/xdr:cNvPr", NS).get("name")
        pos = shape.getparent().find("xdr:pos", NS)
        ext = shape.getparent().find("xdr:ext", NS)
        values = tuple(round(int(pos.get(key)) / 12700) for key in ("x", "y"))
        width = round(int(ext.get("cx")) / 12700)
        assert values + (width,) == expected[name]
        assert round(int(ext.get("cy")) / 12700) == 25


def test_embedded_vba_is_exactly_three_modules(artifact, parts):
    parser = VBA_Parser(str(artifact))
    try:
        sources = {name: code.replace("\r\n", "\n") for _, _, name, code in parser.extract_macros()}
    finally:
        parser.close()
    assert set(sources) == {"ThisWorkbook.cls", "Milestones.cls", "DrawMs.bas"}
    expected = {
        "ThisWorkbook.cls": ROOT / "src/DrawMs_ThisWorkbook.cls",
        "Milestones.cls": ROOT / "src/Milestones.cls",
        "DrawMs.bas": ROOT / "src/DrawMs.bas",
    }
    for name, path in expected.items():
        assert sources[name] == path.read_text().replace("\r\n", "\n")
    code = sources["DrawMs.bas"]
    assert set(re.findall(r"Public Sub (\w+)\(", code)) == {"GenerateChart", "WidenSpacing", "NarrowSpacing"}
    for stale in ("Undo", "Redo", "WBS", "ProgressChart", "Copy", "Import", "DeleteLevel"):
        assert stale not in code
    ole = olefile.OleFileIO(io.BytesIO(parts["xl/vbaProject.bin"]), raise_defects=olefile.DEFECT_INCORRECT)
    assert not ole.parsing_issues
