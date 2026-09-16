# v5.1 Verification Results

Verified in the Linux build environment on 2026-09-16. Desktop Microsoft Excel is unavailable here; VBA execution and performance timings are not claimed.

## Completed

- `uv sync --locked`: passed. Only the project version changes from 5.0.0 to 5.1.0 in `pyproject.toml` and `uv.lock`; dependencies and their locked versions are unchanged.
- `uv run python scripts/build_workbook.py`: generated `dist/WBS-Gantt-ProgressChart-v5.1.xlsm`.
- `uv run pytest -q --disable-warnings`: **18 passed**. The 19 suppressed warnings are existing oletools/pyparsing deprecations. Checks cover workbook relationships, styles, embedded VBA, public button bindings, schedule examples and rollups, ordered cells, merge ranges, DrawingML geometry, and VBA/compound-file round trips.
- All three VBA sources pass ANTLR VBA 7.1 syntax parsing with **zero syntax errors**. Source lines stay within VBA's physical-line limit. This is syntax validation, not Excel compilation or COM type checking.
- Both worksheets and both drawings pass ECMA-376 XML schema checks. The workbook and VBA container open with openpyxl and olefile/oletools. Embedded modified modules match source; stale compiled caches are removed.
- The WBS example has **14 rows instead of 59**. Every retained B:Y cell value matches v5. Dates, predecessors, leaf-weighted progress and the calendar-day timeline remain consistent; the shorter sample spans 2018/8/1–2018/9/3. The legacy build template is byte-for-byte unchanged.
- The Level validation still offers 1–5. Two new buttons bind to `PromoteTasks` and `DemoteTasks`.
- Saved marker centers align with the **visible top of the red arrow shaft**. The nearest date box ends 2 points above the triangle tip. The runtime drawing uses the same dimensions and explicit arrow adjustments. The regenerated PNG was visually inspected; it renders saved DrawingML and is not an Excel screenshot.
- Milestones instructions were reduced from 648 to 186 characters; repeated sample-note text was removed.
- Removed unused `code.vbs`, the old recursive scheduling/name-scan code, the single-row level setter and unused VBA state. The new dependency queue avoids recursive calls and detects incomplete/circular graphs before writing schedule values. `Refresh All` now stops when Gantt scheduling fails.
- The ZIP includes the workbook, preview, English/Chinese documentation and source project. Packaging checks ZIP integrity, the workbook's identity, and every payload SHA-256 checksum.

## Windows checks prepared, not executed

`tests/excel_smoke.ps1` retains scheduling, insertion, compact-date, weekend, Gantt cleanup, linked-import and crowded-chart scenarios and adds:

- Continuous/Ctrl/whole-row task selection; moving from task selection into D2; Promote/Demote; mixed levels and overlapping selected parent/child subtrees.
- Preserved predecessor formulas after level renumbering; invalid first-row promotion and skipped levels; circular predecessors without partial schedule writes.
- Date edits that preserve row formatting, predecessor validation and manual calculation mode.
- A 3,000-task predecessor chain with an actual Excel timing output, without an assumed speed threshold.
- Marker alignment and the reduced nearest-date gap after generation and spacing changes.

The script opens a read-only workbook, edits in memory and closes without saving. No Excel macro-security settings are changed. Interactive Excel checks still include invalid dates/level boundaries, shape-button clicks with real Ctrl selections, no repair/compile prompts, copying to PowerPoint, and saving/reopening.

File/schema/syntax results cannot establish desktop behavior or a measured speedup. The performance changes remove identified redundant work; the remaining schedule scan and necessary writes still grow with task/dependency count.
