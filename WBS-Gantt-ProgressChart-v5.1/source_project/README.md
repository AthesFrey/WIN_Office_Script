# WBS Gantt Progress Chart v5.1

Open `WBS-Gantt-ProgressChart-v5.1.xlsm` in desktop Microsoft Excel with macros enabled. It retains v5's five task levels, calendar-day scheduling, parent summaries, daily Gantt timeline and linked milestone charts.

v5.1 reduces the WBS sample from 59 to 14 tasks, adds batch Level selection and Promote / Demote buttons, and moves routine date/progress calculations into memory. Milestone dates sit closer to their markers; marker centers align with the visible top edge of the red bar. Worksheet instructions are shorter.

See `USAGE.md` for operation and `CHECK_RESULTS.md` for actual verification scope. Chinese copies are provided as `USAGE_zh-CN.md` and `CHECK_RESULTS_zh-CN.md` in the delivery ZIP. No desktop Excel or timing benchmark was run in the Linux build environment.

## Source project

The ZIP includes the workbook, preview, documentation and `source_project/`. The source folder contains VBA, build scripts, tests, the dependency lockfile and the unchanged legacy `WBS.xlsm` build template. Use the v5.1 workbook at the ZIP root for planning.

```sh
uv sync --locked
uv run python scripts/build_workbook.py
uv run pytest -q
```

Output: `dist/WBS-Gantt-ProgressChart-v5.1.xlsm`. Python 3.11 or later is required. Only the project version changed in `pyproject.toml` and `uv.lock`; package dependencies are unchanged.

Optional checks use external ANTLR VBA 7.1 parser files and ECMA-376 schemas:

```sh
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/Model_1.bas
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/ProgressChart.bas
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/Sheet1.cls
uv run python scripts/check_schemas.py --schema-dir /path/to/ooxml-xsd dist/WBS-Gantt-ProgressChart-v5.1.xlsm
uv run python scripts/render_preview.py --font /path/to/font.otf
uv run python scripts/package_release.py
```

The preview renders saved DrawingML; it is not an Excel screenshot. The package contains SHA-256 checksums for every payload file.
