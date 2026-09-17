# v5.6fix3 Verification Results

Checked on 2026-09-17 in Linux, targeting Windows desktop Microsoft Excel 2021/2024. No desktop Excel is available here. File checks and parser results do not establish that Excel compilation or interaction has passed.

## Executed

- `uv sync --locked` passed. Project version is `5.6.3`; all dependency entries are unchanged from fix2, as is the legacy `WBS.xlsm` build template.
- `uv run python scripts/build_workbook.py` produced `dist/WBS-Gantt-ProgressChart-v5.6fix3.xlsm`.
- `uv run pytest -q`: **24 passed**, with 19 existing oletools/pyparsing deprecation warnings. Existing schedule, progress, fonts, shape bindings, package relationships, VBA storage and compression checks remain.
- Updated toolbar checks verify that every button on both sheets uses an absolute worksheet anchor with no cell markers, and matching shape coordinates/extents. All 11 Milestones buttons retain their exact initial positions, 3-point top inset, 25-point height, gaps, reserved-row clearance and macro bindings.
- The same fixed-toolbar check rejects the actual fix2 workbook and passes the actual fix3 workbook. This detects the saved placement defect without relying on an event or button macro to repair it.
- Changed VBA modules `Sheet4.cls` and `ProgressChart.bas` passed ANTLR VBA 7.1 parsing with zero syntax errors. The other three modules are byte-identical to fix2. All five embedded modules in the delivered workbook match current source after CRLF normalization; physical lines are below 1,023 bytes. ANTLR does not compile VBA or type-check Excel COM.
- Both worksheets, both drawings and styles passed ECMA-376 schema checks. Later Office extension attributes are excluded as documented by the checker.
- `excel_smoke.ps1`, `excel_milestones.ps1` and `excel_regressions.ps1` passed PowerShell 7.4.6 syntax parsing. Their Excel scenarios were **not executed**.
- Comparing the generated XLSM with fix2 changes exactly three parts: VBA, Milestones drawing and the WBS worksheet version title. All other parts are byte-identical; the WBS worksheet also matches after replacing only its version label. Saved data, styles, validation, WBS toolbar and chart geometry are preserved.
- Source review confirms removal of the cell-anchor converter, `move_row` option, unused imports, layout guard, duplicate runtime coordinates, `ResetToolbarLayout` and both callers. No obsolete runtime layout references remain. The replacement tests check fixed placement directly. All public button bindings are retained.
- The toolbar preview renders actual saved shape transforms and was visually checked for button overlap and settings/status/header clearance. It is not an Excel screenshot; cell text is simplified. The sample chart preview also renders DrawingML.
- Release ZIP CRC and SHA-256 checks passed for all 34 payload files. The bundled source project was extracted to a temporary directory and rebuilt; every internal XLSM part matches the delivered workbook. The standalone XLSM matches the copy in the ZIP. The package includes current repair notes, both previews and reproducible source.


## Prepared but not executed in desktop Excel

Run `tests/excel_smoke.ps1` from a trusted location; it loads the two supplementary scripts. It opens the supplied workbook read-only, closes without overwriting it, and uses a disposable copy for save/reopen layout checks. It does not change macro security.

Updated layout checks assert Excel `Placement = 3` (Don't move or size with cells) and exact Left/Top/Width/Height. They cover first open, 80%/100%/125% zoom, increasing/decreasing column width and row height, and moving a sample range down/right and back. Each move is checked immediately before activation or macro invocation, with events both disabled and enabled. They also cover sheet activation, the Generate button, Refresh All, and saving/reopening after row/column changes.

Previous input-error scenarios, WBS scheduling, mixed-format paste, both independent three-step histories, shortcuts, five-level deletion, protected replay, source structural changes and 3,000-task scenarios remain prepared. Physical mouse dragging, actual Excel compilation/repair prompts, Save As, clipboard export and other interactive checks still require desktop Excel. No runtime or performance result is claimed for this environment.
