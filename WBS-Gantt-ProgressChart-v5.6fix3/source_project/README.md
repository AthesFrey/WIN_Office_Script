# WBS Gantt Progress Chart v5.6fix3

Open `WBS-Gantt-ProgressChart-v5.6fix3.xlsm` in Windows desktop Microsoft Excel 2021/2024 with macros enabled. It retains v5's five task levels, calendar-day scheduling, parent summaries, daily Gantt timeline and linked milestone charts.

5.6fix3 makes all 11 Milestones buttons use the same fixed worksheet positioning as Sheet1 (WBS): Excel's ‘Don't move or size with cells’. Resizing rows/columns or moving cells no longer relies on a later macro to restore button positions. The initial layout retains the 3-point top inset, 25-point height, horizontal gaps and all button bindings. The obsolete cell-anchor conversion and runtime repositioning routine have been removed. See `FIX3_NOTES.md`.

v5.6fix3 retains seven buttons above the Milestones table: Undo / Redo and Delete Level 1–5. Each delete clears only entries linked to WBS tasks of exactly that level; other levels, manual entries and the WBS source remain. Row positions, formatting and validation are preserved. Click Generate after deletion or history replay to update the drawing.

WBS and Milestones now keep independent three-operation histories. Milestones records data edits, pastes, clearing, batch import, level deletion and position/spacing settings. Each batch counts once; pure chart redraws do not count. Date/format normalization during generation counts as one data operation. Ctrl+Z / Ctrl+Y follow the active sheet. Failed replay retains its entry for retry, and stale snapshots cannot overwrite detected later changes. Excel's native Undo, Redo and Repeat remain disabled for this workbook.

The delivered data, styles, WBS controls, scheduling and hierarchy are preserved from 5.6fix2. Prior fixes for error 28, failed-refresh rollback, date validation and import validation remain. Milestone generation retains its cleanup-and-retry behavior. Refresh All clears both histories; history is not saved across close. Invalid source-row values are skipped during import duplicate checks and level deletion.

See `USAGE.md` for operation and `CHECK_RESULTS.md` for actual verification scope. Chinese copies are provided as `USAGE_zh-CN.md` and `CHECK_RESULTS_zh-CN.md` in the delivery ZIP. No desktop Excel or timing benchmark was run in the Linux build environment.

## Source project

The ZIP includes the workbook, chart and toolbar previews, repair notes, documentation and `source_project/`. The source folder contains VBA, build scripts, tests, the dependency lockfile and the unchanged legacy `WBS.xlsm` build template. Use the v5.6fix3 workbook at the ZIP root for planning.

```sh
uv sync --locked
uv run python scripts/build_workbook.py
uv run pytest -q
```

Output: `dist/WBS-Gantt-ProgressChart-v5.6fix3.xlsm`. Python 3.11 or later is required. Package dependencies are unchanged.

Optional checks use external ANTLR VBA 7.1 parser files and ECMA-376 schemas:

```sh
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/Model_1.bas
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/ProgressChart.bas
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/Sheet1.cls
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/Sheet4.cls
uv run python scripts/check_vba.py --parser-dir /path/to/vba-parser src/ThisWorkbook.cls
uv run python scripts/check_schemas.py --schema-dir /path/to/ooxml-xsd dist/WBS-Gantt-ProgressChart-v5.6fix3.xlsm
uv run python scripts/render_preview.py --font /path/to/font.otf
uv run python scripts/render_preview.py --toolbar --font /path/to/font.otf
uv run python scripts/package_release.py
```

The previews render saved DrawingML; they are not Excel screenshots. The package contains SHA-256 checksums for every payload file.
