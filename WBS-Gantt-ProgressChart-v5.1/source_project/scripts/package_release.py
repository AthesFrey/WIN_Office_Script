"""Package the v5.1 workbook and reproducible source project with English paths."""
import argparse
import hashlib
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from build_workbook import ROOT, WORKBOOK_NAME


def package(output: Path):
    files = {WORKBOOK_NAME: ROOT / 'dist' / WORKBOOK_NAME,
             'progress_chart_preview.png': ROOT / 'dist/progress_chart_preview.png'}
    for name in ['README.md', 'USAGE.md', 'CHECK_RESULTS.md']:
        files[name] = ROOT / name
    for name, source in [('USAGE_zh-CN.md', '使用说明.md'), ('CHECK_RESULTS_zh-CN.md', '检查结果.md')]:
        path = ROOT / source if (ROOT / source).is_file() else ROOT / name
        files[name] = path
        files['source_project/' + name] = path
    for name in ['WBS.xlsm', 'pyproject.toml', 'uv.lock', 'README.md', 'USAGE.md', 'CHECK_RESULTS.md']:
        files['source_project/' + name] = ROOT / name
    for folder, extensions in [('src', {'.bas', '.cls'}), ('scripts', {'.py'}), ('tests', {'.py', '.ps1'})]:
        for path in sorted((ROOT / folder).iterdir()):
            if path.is_file() and path.suffix in extensions:
                files['source_project/' + path.relative_to(ROOT).as_posix()] = path
    assert all(name.isascii() for name in files)
    output.parent.mkdir(parents=True, exist_ok=True)
    checksums = []
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for name, path in sorted(files.items()):
            data = path.read_bytes()
            archive.writestr(name, data)
            checksums.append(hashlib.sha256(data).hexdigest() + '  ' + name)
        archive.writestr('SHA256SUMS.txt', '\n'.join(checksums) + '\n')
    with ZipFile(output) as archive:
        assert archive.testzip() is None
        assert archive.read(WORKBOOK_NAME) == (ROOT / 'dist' / WORKBOOK_NAME).read_bytes()
        for line in archive.read('SHA256SUMS.txt').decode().splitlines():
            digest, name = line.split('  ', 1)
            assert hashlib.sha256(archive.read(name)).hexdigest() == digest, name
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, default=ROOT.parent / 'WBS-Gantt-ProgressChart-v5.1.zip')
    print(package(parser.parse_args().output))
