# Loaded by excel_smoke.ps1; uses its visible, read-only Excel workbook and helpers.
function Test-Milestones {
    $prefix = "'$($book.Name.Replace("'", "''"))'!"
    $excel.EnableEvents = $false
    $wbs.Activate()
    $wbs.Range('B6:Y3010').ClearContents()
    for ($group = 0; $group -lt 2; $group++) {
        for ($level = 1; $level -le 5; $level++) {
            $row = 5 + 5 * $group + $level
            $wbs.Cells.Item($row, 3).Value2 = (' ' * (4 * ($level - 1))) + "Group $group level $level"
        }
    }
    foreach ($row in @(10, 15)) {
        $wbs.Cells.Item($row, 7).Value2 = ([datetime]'2026-09-01').ToOADate()
        $wbs.Cells.Item($row, 6).Value2 = 2.0
        $wbs.Cells.Item($row, 14).Value2 = 1.0 / 3.0
    }
    Run-Macro 'Model_1.RefreshParentTasks'
    $pg.Range('B14:J213').ClearContents()
    $pg.Range('C6').Value2 = 'Equal spacing'
    $pg.Range('C7').Value2 = 1.0
    Fill-Node 30 '    2.1 Manual entry, keep at every level' ([datetime]'2026-09-05')
    $pg.Range('F30').NumberFormat = '@'
    $pg.Range('F30').Value2 = '00123'
    $pg.Range('F31').NumberFormat = '@'
    $pg.Range('F31').Value2 = '=literal text'
    $pg.Range('J30').NumberFormat = 'General'
    $pg.Range('J30').Formula = '=1/3'
    $pg.Range('J30').NumberFormat = '@'
    $pg.Range('H30').Value2 = 1.0 / 3.0
    $pg.Range('H30').NumberFormat = '0.00%'
    # Invalid, fractional, overflowing and dangling sources must never select a task.
    $invalid = @('=#REF!', 3.0, 6.5, 2e12, 'not a row', $true, 900001.0)
    for ($i = 0; $i -lt $invalid.Count; $i++) {
        $row = 31 + $i
        $pg.Cells.Item($row, 2).Value2 = 'No'
        $pg.Cells.Item($row, 4).Value2 = "Keep invalid source $i"
        if ($i -eq 0) { $pg.Cells.Item($row, 3).Formula = $invalid[$i] }
        else { $pg.Cells.Item($row, 3).Value2 = $invalid[$i] }
    }
    $excel.EnableEvents = $true
    Run-Macro 'Model_1.ResetUndoTracking'
    Run-Macro 'Sheet4.ResetUndoTracking'
    $beforeImport = Get-RangeState $pg.Range('B14:J37')
    $wbs.Range('C6:C15').Select()
    Run-Macro 'ProgressChart.ImportSelectedMilestones'
    Assert-MilestoneHistory 1 0
    $imported = Get-RangeState $pg.Range('B14:J37')
    Undo-Milestones
    Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $beforeImport) 'Import Undo did not remove the entire batch'
    Redo-Milestones
    Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $imported) 'Import Redo lost links or formats'
    Assert-History 0 0
    $pg.Range('B15').Value2 = 'No'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    $chartId = $pg.Shapes.Item('WBS_ProgressChart').ID
    Run-Macro 'Sheet4.ResetUndoTracking'

    # Each real button deletes exactly its level, including Show=No, in one step.
    $baseline = Get-RangeState $pg.Range('B14:J37')
    $sourceState = Get-RangeState $wbs.Range('B6:Y15')
    for ($level = 1; $level -le 5; $level++) {
        $rowA = 13 + $level
        $rowB = 18 + $level
        $excel.Run($pg.Shapes.Item("PG_Button_DeleteMilestoneLevel$level").OnAction) | Out-Null
        Assert-MilestoneHistory 1 0
        for ($row = 14; $row -le 23; $row++) {
            if ($row -eq $rowA -or $row -eq $rowB) {
                Assert-True ($excel.WorksheetFunction.CountA($pg.Range("B${row}:J${row}")) -eq 0) 'Matching level was not cleared'
                Assert-True ($pg.Cells.Item($row, 2).Validation.Formula1 -eq 'Yes,No') 'Delete removed Show validation'
            } else {
                Assert-True ($pg.Cells.Item($row, 4).HasFormula) 'Delete removed another level or a descendant'
            }
        }
        Assert-True ($pg.Range('D30').Value2 -like '*Manual entry*') 'Delete removed a manual entry'
        Assert-True ($excel.WorksheetFunction.CountA($pg.Range('D31:D37')) -eq 7) 'Delete removed an invalid-source entry'
        Assert-True ((Get-RangeState $wbs.Range('B6:Y15')) -eq $sourceState) 'Delete changed WBS source tasks'
        Assert-True ($pg.Shapes.Item('WBS_ProgressChart').ID -eq $chartId) 'Delete redrew the chart without Generate'
        $deleted = Get-RangeState $pg.Range('B14:J37')
        # Repeating a deletion with no matches must not consume a step.
        Run-Macro "ProgressChart.DeleteMilestoneLevel$level"
        Assert-MilestoneHistory 1 0
        Undo-Milestones
        Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $baseline) 'Delete Undo lost row positions, formulas or precision'
        Redo-Milestones
        Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $deleted) 'Delete Redo did not repeat the same batch'
        Undo-Milestones
        Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $baseline) 'Repeated replay did not restore the original data'
        Assert-True ($pg.Shapes.Item('WBS_ProgressChart').ID -eq $chartId) 'History replay redrew the chart'
    }
    Assert-True (-not $pg.Range('F31').HasFormula -and $pg.Range('F31').Value2 -ceq '=literal text') 'Replay parsed literal equals text'
    Assert-True ($pg.Range('F30').Value2 -is [string] -and $pg.Range('F30').Value2 -ceq '00123') 'Replay lost leading zeros'
    Assert-True ($pg.Range('J30').HasFormula -and $pg.Range('J30').NumberFormat -eq '@') 'Replay lost a Text-formatted formula'
    Assert-Progress $pg.Range('H30') (1.0 / 3.0) 'Milestones replay precision'

    # Multi-cell paste captures foreign formats and validation as a single action.
    $pg.Range('L230:N231').NumberFormat = 'General'
    $pg.Range('L230:N231').Value2 = 7.0
    $pg.Range('L230:N231').Font.Italic = $true
    $pg.Range('L230:N231').Font.Color = 255
    $pg.Range('L230:N231').Interior.Color = 65535
    $pg.Range('L230:N231').Borders.LineStyle = -4118
    $pg.Range('G31').Validation.Delete()
    $pg.Range('G31').Validation.Add(3, 1, 1, 'alpha,beta')
    $pg.Range('E30:G31').Select()
    Run-Macro 'Sheet4.ResetUndoTracking'
    $before = Get-RangeState $pg.Range('B14:J37')
    $pg.Range('L230:N231').Copy()
    $pg.Range('E30:G31').PasteSpecial(-4104) | Out-Null
    $excel.CutCopyMode = $false
    Assert-MilestoneHistory 1 0
    $after = Get-RangeState $pg.Range('B14:J37')
    Undo-Milestones
    Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $before) 'Milestones paste Undo lost mixed formats'
    Assert-True ($pg.Range('G31').Validation.Formula1 -eq 'alpha,beta') 'Milestones paste Undo lost validation'
    Redo-Milestones
    Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $after) 'Milestones paste Redo lost pasted values or styles'
    Undo-Milestones

    # Four consecutive edits keep only the latest three; selection preserves them.
    Run-Macro 'Sheet4.ResetUndoTracking'
    $states = @((Get-RangeState $pg.Range('B14:J37')))
    for ($i = 1; $i -le 4; $i++) {
        $pg.Range('G30').Value2 = "Milestone owner $i"
        $states += Get-RangeState $pg.Range('B14:J37')
        Assert-MilestoneHistory ([math]::Min($i, 3)) 0
    }
    for ($i = 3; $i -ge 1; $i--) {
        $pg.Range('D30').Select()
        Undo-Milestones
        Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $states[$i]) 'Milestones three-step Undo used the wrong state'
    }
    Assert-MilestoneHistory 0 3
    Run-Macro 'ProgressChart.UndoMilestones'
    for ($i = 2; $i -le 4; $i++) {
        Redo-Milestones
        Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $states[$i]) 'Milestones three-step Redo used the wrong state'
    }
    Undo-Milestones
    Assert-MilestoneHistory 2 1
    $pg.Range('G30').Value2 = $pg.Range('G30').Value2
    Assert-MilestoneHistory 2 1
    $pg.Range('G30').Value2 = 'New branch'
    Assert-MilestoneHistory 3 0
    Run-Macro 'ProgressChart.RedoMilestones'
    Assert-True ($pg.Range('G30').Value2 -eq 'New branch') 'Milestones retained an obsolete Redo branch'

    # Keyboard dispatch and independent histories on both pages.
    $wbs.Activate()
    Run-Macro 'Model_1.ResetUndoTracking'
    $wbs.Range('D20').Select()
    $wbs.Range('D20').Value2 = 'Independent WBS edit'
    Assert-History 1 0
    $pg.Activate()
    Assert-MilestoneHistory 3 0
    $pg.Range('G30').Select()
    Send-ExcelKeys '^z'
    Assert-MilestoneHistory 2 1
    Assert-History 1 0
    Send-ExcelKeys '^y'
    Assert-MilestoneHistory 3 0
    $wbs.Activate()
    Send-ExcelKeys '^z'
    Assert-History 0 1
    Assert-MilestoneHistory 3 0
    $pg.Activate()
    $other = $excel.Workbooks.Add()
    try {
        $other.Worksheets.Item(1).Range('A1').Select()
        Send-ExcelKeys '789{ENTER}'
        Assert-True $excel.CommandBars.GetEnabledMso('Undo') 'Leaving Milestones did not restore native Undo'
        Send-ExcelKeys '^z'
        Assert-True ($null -eq $other.Worksheets.Item(1).Range('A1').Value2) 'Milestones shortcut leaked to another workbook'
    } finally { $other.Close($false) }
    $book.Activate()
    $pg.Activate()
    Assert-MilestoneHistory 3 0

    # Protected replay retains the failed direction for retry and changes no cells.
    Undo-Milestones
    $protected = Get-RangeState $pg.Range('B14:J37')
    $calculation = $excel.Calculation
    $pg.Protect()
    $failed = $false
    try { $excel.Run($prefix + 'Sheet4.RedoLastEdit', $true) | Out-Null } catch { $failed = $true }
    $pg.Unprotect()
    Assert-True $failed 'Protected Milestones Redo did not report a failure'
    Assert-MilestoneHistory 2 1
    Assert-True ((Get-RangeState $pg.Range('B14:J37')) -eq $protected) 'Protected Redo corrupted cells'
    Assert-True ($excel.EnableEvents -and $excel.ScreenUpdating -and $excel.Calculation -eq $calculation) 'Failed replay changed application settings'
    Redo-Milestones
    $pg.Protect()
    $failed = $false
    try { $excel.Run($prefix + 'Sheet4.UndoLastEdit', $true) | Out-Null } catch { $failed = $true }
    $pg.Unprotect()
    Assert-True $failed 'Protected Milestones Undo did not report a failure'
    Assert-MilestoneHistory 3 0
    Undo-Milestones
    Redo-Milestones
    $pg.Protect()
    $failed = $false
    try { $excel.Run($prefix + 'ProgressChart.DeleteMilestonesByLevel', 1, $true) | Out-Null } catch { $failed = $true }
    $pg.Unprotect()
    Assert-True $failed 'Protected level deletion did not report a failure'
    Assert-MilestoneHistory 3 0

    # External changes invalidate snapshots; new input cannot absorb outside edits.
    Undo-Milestones
    $excel.EnableEvents = $false
    $pg.Range('G31').Value2 = 'External edit'
    $excel.EnableEvents = $true
    Run-Macro 'ProgressChart.RedoMilestones'
    Assert-MilestoneHistory 0 0
    Assert-True ($pg.Range('G31').Value2 -eq 'External edit') 'Stale history overwrote an external edit'
    $pg.Range('G30').Select()
    $pg.Range('G30').Value2 = 'Recorded edit'
    $pg.Range('G31').Font.Italic = -not $pg.Range('G31').Font.Italic
    Run-Macro 'ProgressChart.UndoMilestones'
    Assert-MilestoneHistory 0 0
    Run-Macro 'Sheet4.ResetUndoTracking'
    $pg.Range('G30').Value2 = 'Before outside edit'
    $excel.EnableEvents = $false
    $pg.Range('G31').Value2 = 'Keep outside edit'
    $excel.EnableEvents = $true
    $pg.Range('G30').Value2 = 'After outside edit'
    Assert-MilestoneHistory 0 0
    Assert-True ($pg.Range('G31').Value2 -eq 'Keep outside edit') 'An edit absorbed unrelated external changes'

    # Recalculation changes formula results without invalidating formula history.
    $pg.Range('G30').Select()
    $pg.Range('G30').Value2 = 'Before source change'
    $excel.Calculation = -4135
    $wbs.Range('N10').Value2 = 0.75
    Undo-Milestones
    Assert-Progress $pg.Range('H18') 0.75 'Milestones Undo restored stale cached WBS progress'
    Redo-Milestones
    Assert-Progress $pg.Range('H18') 0.75 'Milestones Redo restored stale cached WBS progress'
    $excel.Calculation = -4105

    # Level resolution is live, and WBS insertion must not replay obsolete links.
    $wbs.Range('C10').Value2 = '            Leaf moved to level four'
    Run-Macro 'ProgressChart.DeleteMilestoneLevel5'
    Assert-True ($pg.Range('D18').HasFormula -and $null -eq $pg.Range('D23').Value2) 'Deletion used an imported/stale level'
    Undo-Milestones
    $wbs.Rows.Item(8).Insert() | Out-Null
    $pg.Calculate()
    Run-Macro 'ProgressChart.RedoMilestones'
    Assert-MilestoneHistory 0 0
    Assert-True ($pg.Range('C18').Value2 -eq 11) 'WBS insertion broke source-row tracking'
    $linkedBefore = $pg.Range('C18').Formula
    Run-Macro 'ProgressChart.DeleteMilestoneLevel4'
    Assert-True ($null -eq $pg.Range('D18').Value2) 'Deletion did not use the adjusted source row'
    Undo-Milestones
    Assert-True ($pg.Range('C18').Formula -eq $linkedBefore) 'Undo restored an obsolete source-row formula'

    # Settings and normalizing many dates each consume one step; pure redraw none.
    $excel.EnableEvents = $false
    $pg.Range('B14:J213').ClearContents()
    Fill-Node 14 'First normalization' ([datetime]'2026-09-01')
    Fill-Node 15 'Second normalization' ([datetime]'2026-09-02')
    $pg.Range('E14:E15').NumberFormat = '@'
    $pg.Range('E14').Value2 = '20260901'
    $pg.Range('E15').Value2 = '0902'
    $excel.EnableEvents = $true
    Run-Macro 'Sheet4.ResetUndoTracking'
    $before = Get-RangeState $pg.Range('B14:J15')
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-MilestoneHistory 1 0
    Assert-Date $pg.Range('E15') ([datetime]'2026-09-02') 'Date normalization failed'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-MilestoneHistory 1 0
    $chartId = $pg.Shapes.Item('WBS_ProgressChart').ID
    Undo-Milestones
    Assert-True ((Get-RangeState $pg.Range('B14:J15')) -eq $before) 'Normalization Undo did not restore input text'
    Assert-True ($pg.Shapes.Item('WBS_ProgressChart').ID -eq $chartId) 'Normalization Undo unexpectedly redrew the chart'
    Redo-Milestones
    $pg.Range('C6').Value2 = 'Date proportional'
    Assert-MilestoneHistory 2 0
    Run-Macro 'ProgressChart.WidenProgressChart'
    Assert-MilestoneHistory 3 0
    Undo-Milestones
    Assert-True ([math]::Abs($pg.Range('C7').Value2 - 1) -lt 1e-12) 'Spacing Undo did not restore the setting'
    Undo-Milestones
    Assert-True ($pg.Range('C6').Value2 -eq 'Equal spacing') 'Position Undo did not restore the setting'
    # Failed generation must roll back any earlier row normalization as one transaction.
    $pg.Range('E14').Value2 = '20260903'
    $pg.Range('E15').Value2 = 'invalid'
    $before = Get-RangeState $pg.Range('B14:J15')
    $failed = $false
    try { $excel.Run($prefix + 'ProgressChart.GenerateProgressChart', $true) | Out-Null } catch { $failed = $true }
    Assert-True $failed 'Invalid milestone date was accepted'
    Assert-True ((Get-RangeState $pg.Range('B14:J15')) -eq $before) 'Failed chart generation left partially normalized dates'
    Assert-NoMilestoneChart

    # Delete every populated row, then no-op with Redo available and an empty sheet.
    $pg.Range('B14:J213').ClearContents()
    $wbs.Activate()
    $wbs.Range('C16').Select() # The second group's level-five leaf after insertion.
    Run-Macro 'ProgressChart.ImportSelectedMilestones'
    Run-Macro 'Sheet4.ResetUndoTracking'
    Run-Macro 'ProgressChart.DeleteMilestoneLevel5'
    Assert-True ($excel.WorksheetFunction.CountA($pg.Range('B14:J213')) -eq 0) 'Deleting all matches did not empty the input'
    Assert-MilestoneHistory 1 0
    Undo-Milestones
    Assert-MilestoneHistory 0 1
    Run-Macro 'ProgressChart.DeleteMilestoneLevel1'
    Assert-MilestoneHistory 0 1
    Redo-Milestones
    Run-Macro 'ProgressChart.DeleteMilestoneLevel5'
    Assert-MilestoneHistory 1 0
    # Even with no remaining live links, source insertion invalidates deleted links.
    $wbs.Rows.Item(8).Insert() | Out-Null
    Assert-MilestoneHistory 0 0
    Run-Macro 'ProgressChart.UndoMilestones'
    Assert-True ($excel.WorksheetFunction.CountA($pg.Range('B14:J213')) -eq 0) 'Undo resurrected links to obsolete source rows'
    $pg.Range('F14').Value2 = 'Before Add Subtask'
    Assert-MilestoneHistory 1 0
    $wbs.Activate()
    $wbs.Range('C6').Select()
    Run-Macro 'ProgressChart.AddSubtask'
    Assert-MilestoneHistory 0 0
    $pg.Activate()
    # Row insertion/deletion and out-of-range pastes are history boundaries.
    $pg.Rows.Item(20).Insert() | Out-Null
    Assert-MilestoneHistory 0 0
    $pg.Rows.Item(20).Delete() | Out-Null
    $pg.Range('F14').Select()
    $pg.Range('F14').Value2 = 'Boundary edit'
    Assert-MilestoneHistory 1 0
    $pg.Range('J213:K214').Value2 = 'Outside snapshot'
    Assert-MilestoneHistory 0 0
    Write-Output 'Milestones deletion, history and cross-sheet scenarios passed in desktop Excel.'
}
