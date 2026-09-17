# WBS Gantt Progress Chart v5.6fix3

Open `WBS-Gantt-ProgressChart-v5.6fix3.xlsm` in desktop Microsoft Excel and enable macros. The sheets remain `WBS` and `Milestones`.

Milestones buttons use the same fixed placement as Sheet1 (WBS): ‘Don't move or size with cells’. Their initial positions are within row 5 for Generate / Copy / Widen / Narrow, and rows 11–12 for Undo / Redo and Delete Level 1–5, with a 3-point top inset and 25-point height. Resizing rows/columns or moving cells leaves the buttons at their worksheet coordinates. Opening, activating or refreshing the sheet no longer repositions buttons or adjusts toolbar row heights.

## Tasks and batch levels

Task rows start at row 6. The 14 example rows demonstrate parents, three levels, predecessors and progress; up to five levels remain available. WBS B6:BG19, including blank cells, hidden fields and Gantt cells, is initialized to Microsoft YaHei 11 pt before any refresh.

- `Add Task`: insert a peer after the selected task's subtree.
- `Add Subtask`: insert a child, up to level 5.
- `Level` (D2): select one or more task cells or entire rows, then click D2 and choose a target level. Ctrl-click supports non-contiguous selections. The remembered selection survives the click into D2. Mixed levels show a blank D2 until a value is chosen.
- `Promote (-1)`: move the selected tasks one level toward level 1.
- `Demote (+1)`: move the selected tasks one level deeper.

Each selected parent moves with its descendants, preserving their relative levels. When a parent and its descendant are both selected, the subtree moves once. D2 applies its target level to each selected subtree root. The batch is checked before writing: levels must remain 1–5, the first task must remain level 1, and no task may skip a parent level. Invalid batches show a message without partially changing levels. If the move creates a circular scheduling dependency, the captured body, including numbering, formatting, schedule and predecessor references, is restored.

Numbers and indentation update automatically. Four leading spaces per level still work when editing column C. As in v5, indentation determines which parent owns the following tasks.

For leaf tasks, enter start (G), days (F) and progress (N). WBS progress, Gantt progress bars and Milestones progress display two percentage decimals, such as `33.33%`. Stored values and duration-weighted rollups retain full precision; only the display is rounded. Resource (D), allocation (E) and predecessor (O) remain available. Hidden legacy actual/baseline/deliverable columns are retained.

Parents use the earliest descendant leaf start, latest end, inclusive date span and progress weighted by leaf duration. Empty leaf progress counts as zero. A predecessor starts the task the next calendar day after that predecessor ends and overrides a manual start; clear O to schedule independently. Circular links produce an error.

## Dates and updates

| Input | Result |
| --- | --- |
| `2018/8/30` or `20180830` | `2018/8/30` |
| `0830` or `830`, with a valid 2018 date immediately above in the same column | `2018/8/30` |

Short dates use only the immediately preceding row's year. With no valid date above, enter a full date. Invalid compact dates stay visible as text with a message.

Changing start or days recalculates end. Changing end recalculates days. When pasting start/end/days together, end determines the inclusive duration. Days must be positive whole numbers. Weekends count.

Date, duration and progress changes immediately update dependent tasks and parents. The scheduler reads data into arrays, looks up predecessors by name, processes dependencies without recursion and writes changed results. Resource/allocation edits do not trigger scheduling. Calculation and screen-update settings are restored afterward, including on errors. The custom WBS Undo / Redo buttons also restore linked schedule results. After Gantt refresh, a Days mismatch is marked in F with a light red fill. For adjacent direct siblings under the same parent, a start date that is not the previous end date plus one day is marked in G with a light yellow fill. Actual speed depends on Excel and the plan; no measured speedup is claimed.

## Three-step Undo / Redo

Use `Undo (n)` and `Redo (n)` on either sheet, or Ctrl+Z / Ctrl+Y for the active sheet. WBS keeps its buttons at the top left; Milestones has its buttons immediately above the table. Blue means available; grey means the action does nothing. Each sheet independently retains its latest three edits: after three edits, Undo three times and Redo three times; the fourth edit drops the oldest entry. An edit after Undo discards that sheet's previous redo branch.

Supported operations are WBS body edits in A6:Y, a single multi-cell paste, clearing contents, the D2 level selector and Promote / Demote. Each entry includes both the original and resulting raw values, formulas, literal text, number formats, common cell styles, borders, row heights, predecessor validation and linked schedule results. Dates, leading zeros, text beginning with `=` and full numeric precision are preserved. Replay restores the captured state and recalculates formulas, including linked Milestones cells, without running the scheduler again.

Selecting another cell, changing the displayed D2 selection level, switching sheets or switching workbooks preserves valid history. Invalid or unchanged level operations do not consume a step. Ctrl+Z / Ctrl+Y follow the active WBS or Milestones sheet and release when leaving this workbook or closing it. Excel's native Undo, Redo and Repeat buttons remain disabled for this workbook; native controls are available again in other workbooks.

`Update Gantt`, task insertion and edits outside the supported WBS body retain their existing WBS history boundaries. `Refresh All` clears both histories; structural row/column changes invalidate the affected history. History lives only in memory and is not saved. Pure formatting commands are not recorded as separate operations. Unobserved changes to captured cells invalidate old history so it cannot overwrite later changes. Formatting applied to the same selected cells immediately before a content edit can be included in that edit because Excel has no before-formatting event. Merge/unmerge, conditional-format editing, comments, drawings and page layout are outside this history.

Before a large WBS paste into unused rows, select its full destination to capture the old cells and formats. If a paste extends beyond the captured area, the tool keeps the new data, clears history and explains why. WBS capture includes at least 1,000 rows and the predecessor-validation extension below the task list.

If Undo or Redo fails, the error is shown, Excel's event/calculation/screen settings are restored and the history entry remains available for retry in the same direction. Resolve the reported cause (for example, unprotect WBS) before retrying. A new edit after a failure starts a fresh history.

Warning fills use tool-owned conditional-format rules, so removing an obsolete warning reveals the user's underlying fill. Validation runs after scheduling and calculation: ordinary numeric Days/start edits continue to calculate end dates automatically. Top-level tasks share one sibling group; a parent's descendants between siblings make them nonadjacent and therefore exempt from the yellow warning.

`Update Gantt` repaints the daily timeline from the earliest start through the latest end, clears stale bars when the plan shrinks, and checks schedule warning fills. `Refresh All` calls the same Gantt refresh, then updates the milestone chart and sets body text on both sheets to Microsoft YaHei 11 pt. WBS B6:Y through the last task and Gantt task rows are formatted inside Update Gantt; Refresh All additionally formats Milestones B3:J10 and B14:J213. Titles, headers, buttons and chart-shape text keep their existing fonts. If milestone generation fails, its old generated drawings are deleted and generation is retried once. If the retry also fails, Milestones B10 reports the reason, the chart is skipped and the remaining Refresh All formatting/history reset still runs. The standalone chart and spacing buttons use the same retry behavior. Scheduling failures still stop the refresh.

## Milestones

Enter names and dates with `Show = Yes`, then click `Generate / Refresh Chart`. To import: clear the samples in B:J, select WBS tasks (Ctrl-click supported), then click `Import Selected Tasks`. Imported names, end dates, owners, progress and status remain linked, including after source row insertion. Regenerate after changes.

Choose `Equal spacing` or `Date proportional` in C6. C7 and the spacing buttons adjust width from 50% to 200%. Up to 60 input milestones and 120 characters per name are supported. Milestones sharing a date use one marker and one text box, with names separated by `; `. Alt+Enter adds a line break. Crowded labels/dates use separate lanes and leader lines.

`Delete Level 1` through `Delete Level 5` clear B:J for entries linked to a WBS task of exactly that current level. The buttons scan all 200 input rows, including hidden rows and `Show = No`. They preserve other levels, descendants of other levels, manual entries without a valid source, and WBS source tasks. Invalid or dangling source rows are skipped. Rows do not shift; formats and dropdowns remain. A deletion is one undoable operation, and a deletion with no matches does not consume a step or discard Redo.

Milestones history covers B14:J213 and C6:C7: cell edits, one multi-cell paste, clearing, a whole import batch, level deletion, and position/spacing changes. Raw values, formulas, literal text, number formats, common styles, borders, row heights and validation are captured. Each import or button operation counts once. Generate counts only when it normalizes dates or input formatting; a pure redraw preserves history. Failed import, deletion and generation attempts roll back their input changes and report any rollback failure. If Milestones generation fails during Refresh All, its partial date and number-format normalization is rolled back before the remaining refresh formatting proceeds. Spacing buttons retain their existing automatic chart generation, while Undo / Redo only restore the setting.

After data changes, deletion, Undo or Redo, click `Generate / Refresh Chart` to update the drawing. The old drawing remains until generation; drawings are not part of history. `Refresh All` clears Milestones history. An edit spanning both supported and unsupported cells clears its history; standalone changes outside the captured areas are not recorded. WBS source recalculation keeps linked formulas live without consuming a Milestones step; structural source changes that alter link formulas invalidate old snapshots.

Dates sit just above the triangles. Triangle centers align with the visible top edge of the dark-red arrow shaft. `Copy Chart` copies the grouped drawing for Word or PowerPoint.

## Windows verification

From `source_project`, run in a trusted location with macros allowed:

```powershell
pwsh -File tests/excel_smoke.ps1 -WorkbookPath ../WBS-Gantt-ProgressChart-v5.6fix3.xlsm
```

It opens Excel visibly, modifies an in-memory read-only workbook and closes without saving over the supplied file. The layout test saves and reopens a disposable copy, then deletes that copy. It does not change macro security. Keep Excel in the foreground during the keyboard scenarios. The script checks delivered fonts before refresh, date and percentage precision, mixed-format/text/formula round trips, three-step capacity and both directions, selection and sheet changes, new-edit branching, unchanged and invalid hierarchy operations, protected Undo/Redo retry, stale-history rejection, actual Ctrl+Z / Ctrl+Y, other-workbook native controls, whole-row history boundaries, tail-row validation and close/reopen. Scheduling, milestone grouping, warning colors, font scopes, chart refresh and the local 3,000-task timing scenario remain included.

The runner also loads `tests/excel_milestones.ps1` for all five deletion levels, manual/invalid sources, deletion of every entry, import/paste round trips, independent histories, shortcut dispatch, protected retry, stale snapshots, source row insertion, date normalization, settings and structural boundaries. The runner additionally loads `tests/excel_regressions.ps1` for fixed button coordinates at 80%/100%/125%, cell moves and row/column resizing with events both disabled and enabled (checked before clicking a macro), Refresh All, saving/reopening a temporary copy, failed-refresh rollback, malformed dates, invalid sources, empty imports and import duplicate detection in manual calculation mode. All three scripts passed PowerShell syntax checks on Linux; none of the Excel COM/runtime scenarios was executed here. Windows Office 2021/2024 still needs runtime verification, including actual mouse clicks, Save As, canceled close, partial replay failure, compile/repair prompts and clipboard export to PowerPoint. See the check-results document for the exact executed scope.
