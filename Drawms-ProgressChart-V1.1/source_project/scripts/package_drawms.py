"""Package the standalone drawms workbook and reproducible source files."""
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_NAME = "drawms.xlsm"


def package(output: Path):
    files = {
        WORKBOOK_NAME: ROOT / "dist" / WORKBOOK_NAME,
        "drawms_toolbar_preview.png": ROOT / "dist" / "drawms_toolbar_preview.png",
        "drawms_chart_preview.png": ROOT / "dist" / "drawms_chart_preview.png",
        "README_drawms.md": ROOT / "README_drawms.md",
        "USAGE_drawms.md": ROOT / "USAGE_drawms.md",
        "CHECK_RESULTS_drawms.md": ROOT / "CHECK_RESULTS_drawms.md",
    }
    for name in ["pyproject.toml", "uv.lock"]:
        files["source_project/" + name] = ROOT / name
    files["source_project/WBS.xlsm"] = ROOT / "WBS.xlsm"
    for name in ["build_drawms.py", "package_drawms.py", "build_workbook.py", "upgrade_wbs.py",
                 "check_schemas.py", "vba_project.py", "render_preview.py"]:
        files["source_project/scripts/" + name] = ROOT / "scripts" / name
    for name in ["DrawMs.bas", "Milestones.cls", "DrawMs_ThisWorkbook.cls"]:
        files["source_project/src/" + name] = ROOT / "src" / name
    files["source_project/tests/test_drawms.py"] = ROOT / "tests" / "test_drawms.py"
    files["source_project/tests/excel_drawms.ps1"] = ROOT / "tests" / "excel_drawms.ps1"
    assert all(path.is_file() for path in files.values()), [str(path) for path in files.values() if not path.is_file()]
    assert all(name.isascii() for name in files)
    output.parent.mkdir(parents=True, exist_ok=True)
    checksums = []
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, path in sorted(files.items()):
            data = path.read_bytes()
            archive.writestr(name, data)
            checksums.append(hashlib.sha256(data).hexdigest() + "  " + name)
        archive.writestr("SHA256SUMS.txt", "\n".join(checksums) + "\n")
    with ZipFile(output) as archive:
        assert archive.testzip() is None
        assert archive.read(WORKBOOK_NAME) == (ROOT / "dist" / WORKBOOK_NAME).read_bytes()
        for line in archive.read("SHA256SUMS.txt").decode().splitlines():
            digest, name = line.split("  ", 1)
            assert hashlib.sha256(archive.read(name)).hexdigest() == digest, name
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT.parent / "drawms.zip")
    print(package(parser.parse_args().output))
