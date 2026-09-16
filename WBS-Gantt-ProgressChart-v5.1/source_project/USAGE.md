# WBS Gantt Progress Chart v5.1

Open `WBS-Gantt-ProgressChart-v5.1.xlsm` in desktop Microsoft Excel and enable macros. The sheets remain `WBS` and `Milestones`.

## Tasks and batch levels

Task rows start at row 6. The 14 example rows demonstrate parents, three levels, predecessors and progress; up to five levels remain available.

- `Add Task`: insert a peer after the selected task's subtree.
- `Add Subtask`: insert a child, up to level 5.
- `Level` (D2): select one or more task cells or entire rows, then click D2 and choose a target level. Ctrl-click supports non-contiguous selections. The remembered selection survives the click into D2. Mixed levels show a blank D2 until a value is chosen.
- `Promote (-1)`: move the selected tasks one level toward level 1.
- `Demote (+1)`: move the selected tasks one level deeper.

Each selected parent moves with its descendants, preserving their relative levels. When a parent and its descendant are both selected, the subtree moves once. D2 applies its target level to each selected subtree root. The batch is checked before writing: levels must remain 1–5, the first task must remain level 1, and no task may skip a parent level. Invalid batches show a message without partially changing levels. If the move creates a circular scheduling dependency, its level and predecessor changes are restored.

Numbers and indentation update automatically. Four leading spaces per level still work when editing column C. As in v5, indentation determines which parent owns the following tasks.

For leaf tasks, enter start (G), days (F) and progress (N). Resource (D), allocation (E) and predecessor (O) remain available. Hidden legacy actual/baseline/deliverable columns are retained.

Parents use the earliest descendant leaf start, latest end, inclusive date span and progress weighted by leaf duration. Empty leaf progress counts as zero. A predecessor starts the task the next calendar day after that predecessor ends and overrides a manual start; clear O to schedule independently. Circular links produce an error.

## Dates and updates

| Input | Result |
| --- | --- |
| `2018/8/30` or `20180830` | `2018/8/30` |
| `0830` or `830`, with a valid 2018 date immediately above in the same column | `2018/8/30` |

Short dates use only the immediately preceding row's year. With no valid date above, enter a full date. Invalid compact dates stay visible as text with a message.

Changing start or days recalculates end. Changing end recalculates days. When pasting start/end/days together, end determines the inclusive duration. Days must be positive whole numbers. Weekends count.

Date, duration and progress changes immediately update dependent tasks and parents. v5.1 skips numbering, row formatting and dropdown rebuilding for these edits; it reads data into arrays, looks up predecessors by name, processes dependencies without recursion and writes changed results. Resource/allocation edits do not trigger scheduling. Calculation and screen-update settings are restored afterward, including on errors. Actual speed depends on Excel and the plan; no measured speedup is claimed.

`Update Gantt` repaints the daily timeline from the earliest start through the latest end and clears stale bars when the plan shrinks. `Refresh All` updates both charts and stops if scheduling fails.

## Milestones

Enter names and dates with `Show = Yes`, then click `Generate / Refresh Chart`. To import: clear the samples in B:J, select WBS tasks (Ctrl-click supported), then click `Import Selected Tasks`. Imported names, end dates, owners, progress and status remain linked, including after source row insertion. Regenerate after changes.

Choose `Equal spacing` or `Date proportional` in C6. C7 and the spacing buttons adjust width from 50% to 200%. Up to 60 milestones and 120 characters per name are supported. Alt+Enter adds a line break. Crowded labels/dates use separate lanes and leader lines.

Dates sit just above the triangles. Triangle centers align with the visible top edge of the dark-red arrow shaft. `Copy Chart` copies the grouped drawing for Word or PowerPoint.

## Windows verification

From `source_project`, run in a trusted location with macros allowed:

```powershell
pwsh -File tests/excel_smoke.ps1 -WorkbookPath ../WBS-Gantt-ProgressChart-v5.1.xlsm
```

It opens Excel visibly, modifies an in-memory read-only workbook and closes without saving. It does not change macro security. It also prints a local 3,000-task date-edit timing. This script was not run in the Linux build environment.

Interactive checks still needed: invalid dates, invalid hierarchy batches, first-row promotion, fifth-level demotion, Ctrl-selection via actual button clicks, Excel repair/compile messages, and copying into PowerPoint. On invalid level changes, confirm that all selected tasks keep their original levels and calculation/events still work.
