# Windows + Microsoft Excel integration checks. Run in an Excel trusted location:
# pwsh -File tests/excel_smoke.ps1 -WorkbookPath ./dist/WBS-Gantt-ProgressChart-v5.1.xlsm
# Opens read-only and closes without saving. Does not change macro security.
param([Parameter(Mandatory=$true)][string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Run-Macro([string]$Name) {
    $excel.Run("'$($book.Name.Replace("'", "''"))'!$Name") | Out-Null
}
function Assert-Date([object]$Cell, [datetime]$Date, [string]$Message) {
    Assert-True ([math]::Abs([double]$Cell.Value2 - $Date.ToOADate()) -lt 0.001) $Message
    Assert-True ($Cell.NumberFormat -eq 'yyyy/m/d') "Incorrect date format: $Message"
}
function Fill-Node([int]$Row, [string]$Label, [datetime]$Date) {
    $pg.Cells.Item($Row, 2).Value2 = 'Yes'
    $pg.Cells.Item($Row, 4).Value2 = $Label
    $pg.Cells.Item($Row, 5).Value2 = $Date.ToOADate()
}
function Assert-Chart([int]$NodeCount) {
    $chart = $pg.Shapes.Item('WBS_ProgressChart')
    Assert-True ($chart.Type -eq 6) 'Chart is not a shape group'
    Assert-True ($pg.Shapes.Count -eq 5) 'Duplicate chart or temporary shapes remain'
    Assert-True ($pg.Range('B10').Value2.StartsWith("Generated $NodeCount milestones;")) 'Generation did not finish'
    $markers = @()
    $textboxes = @()
    $axis = $null
    for ($i = 1; $i -le $chart.GroupItems.Count; $i++) {
        $shape = $chart.GroupItems.Item($i)
        if ($shape.Type -eq 17) { $textboxes += $shape }
        if ($shape.Type -eq 1 -and $shape.AutoShapeType -eq 7) { $markers += $shape }
        if ($shape.Type -eq 1 -and $shape.AutoShapeType -eq 33) { $axis = $shape }
    }
    Assert-True ($markers.Count -eq $NodeCount) 'Incorrect milestone marker count'
    Assert-True ($textboxes.Count -eq 2 * $NodeCount) 'Missing milestone labels or dates'
    Assert-True ($null -ne $axis -and $axis.Fill.ForeColor.RGB -eq 139) 'Progress bar is not dark red'
    foreach ($marker in $markers) {
        Assert-True ([math]::Abs(($marker.Top + $marker.Height / 2) - ($axis.Top + $axis.Height * (1 - 0.55) / 2)) -lt 0.1) 'Marker center is off the visible top edge of the bar'
    }
    $dates = @($textboxes | Where-Object { $_.TextFrame2.TextRange.Text -match '^\d{1,2}/\d{1,2}$' })
    Assert-True ($dates.Count -eq $NodeCount) 'Missing date boxes'
    $nearestDateBottom = ($dates | ForEach-Object { $_.Top + $_.Height } | Measure-Object -Maximum).Maximum
    Assert-True ([math]::Abs($markers[0].Top - $nearestDateBottom - 2) -lt 0.1) 'Dates are too far from the markers'
    for ($i = 0; $i -lt $textboxes.Count; $i++) {
        $a = $textboxes[$i]
        Assert-True ($a.TextFrame2.TextRange.BoundHeight -le $a.Height + 1) 'Text is clipped vertically'
        Assert-True ($a.TextFrame2.TextRange.BoundWidth -le $a.Width + 1) 'Text is clipped horizontally'
        for ($j = $i + 1; $j -lt $textboxes.Count; $j++) {
            $b = $textboxes[$j]
            $safe = ($a.Left + $a.Width + 8 -le $b.Left + 0.1) -or
                    ($b.Left + $b.Width + 8 -le $a.Left + 0.1) -or
                    ($a.Top + $a.Height + 8 -le $b.Top + 0.1) -or
                    ($b.Top + $b.Height + 8 -le $a.Top + 0.1)
            Assert-True $safe 'Milestone textboxes overlap or lack the safety gap'
        }
    }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $book = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path, 0, $true)
    Assert-True ($book.Worksheets.Count -eq 2) 'Workbook should contain only two sheets'
    $wbs = $book.Worksheets.Item('WBS')
    $pg = $book.Worksheets.Item('Milestones')
    $wbs.Activate()
    $excel.EnableEvents = $false
    $wbs.Range('B6:Y100').ClearContents()
    $names = @('Project', '    Phase A', '        Leaf A', '        Leaf B', '    Phase B', '        Leaf C')
    for ($i = 0; $i -lt $names.Count; $i++) { $wbs.Cells.Item(6 + $i, 3).Value2 = $names[$i] }
    foreach ($r in @(8, 9, 11)) { $wbs.Cells.Item($r, 14).Value2 = 0.0 }
    $wbs.Range('G8').Value2 = ([datetime]'2018-08-31').ToOADate()
    $wbs.Range('F8').Value2 = 2.0
    $wbs.Range('N8').Value2 = 1.0
    $wbs.Range('G9').Value2 = ([datetime]'2018-09-10').ToOADate()
    $wbs.Range('F9').Value2 = 1.0
    $wbs.Range('G11').Value2 = ([datetime]'2018-09-04').ToOADate()
    $wbs.Range('F11').Value2 = 3.0
    $wbs.Range('N11').Value2 = 0.5
    $excel.EnableEvents = $true
    Run-Macro 'Model_1.RefreshParentTasks'
    Assert-Date $wbs.Range('H8') ([datetime]'2018-09-01') 'Friday plus two calendar days should end Saturday'
    Assert-Date $wbs.Range('G6') ([datetime]'2018-08-31') 'Parent start does not follow earliest leaf'
    Assert-Date $wbs.Range('H6') ([datetime]'2018-09-10') 'Parent end does not follow latest leaf'
    Assert-True ($wbs.Range('F6').Value2 -eq 11) 'Parent duration is not its inclusive span'
    Assert-True ([math]::Abs($wbs.Range('N6').Value2 - 3.5/6) -lt 0.00001) 'Parent progress must weight descendant leaves, not parent spans'

    $wbs.Range('G8').Value2 = 20180830.0
    Assert-Date $wbs.Range('G8') ([datetime]'2018-08-30') 'Numeric yyyymmdd did not normalize'
    Assert-Date $wbs.Range('H8') ([datetime]'2018-08-31') 'Start date change did not update end'
    $wbs.Range('H8').Value2 = 20180902.0
    Assert-True ($wbs.Range('F8').Value2 -eq 4) 'End date change was overwritten instead of recalculating days'
    $wbs.Range('G9').Value2 = '0830'
    Assert-Date $wbs.Range('G9') ([datetime]'2018-08-30') 'Four-digit date failed to use previous row year'
    $wbs.Range('H9').Value2 = 831.0
    Assert-Date $wbs.Range('H9') ([datetime]'2018-08-31') 'Three-digit end date failed'
    Assert-True ($wbs.Range('F9').Value2 -eq 2) 'Short end date did not recalculate days'
    $wbs.Range('N9').Value2 = 0.25
    Assert-True ([math]::Abs($wbs.Range('N6').Value2 - 6/9) -lt 0.00001) 'Progress edits do not roll up immediately'
    $wbs.Range('O9').Value2 = $wbs.Range('C8').Value2.Replace(' ', '')
    Assert-Date $wbs.Range('G9') ([datetime]'2018-09-03') 'Predecessor should schedule next calendar day'
    $wbs.Range('F8').Value2 = 5.0
    Assert-Date $wbs.Range('G9') ([datetime]'2018-09-04') 'Predecessor change did not propagate'
    $wbs.Range('O9').ClearContents()

    # Fast date edits must not rewrite task formatting or the predecessor dropdown.
    $wbs.Rows.Item(8).RowHeight = 57
    $wbs.Range('C8').WrapText = $false
    $validationFormula = $wbs.Range('O8').Validation.Formula1
    $excel.Calculation = -4135
    $wbs.Range('F8').Value2 = 6.0
    Assert-True ($wbs.Rows.Item(8).RowHeight -eq 57 -and -not $wbs.Range('C8').WrapText) 'Date edit rebuilt hierarchy formatting'
    Assert-True ($wbs.Range('O8').Validation.Formula1 -eq $validationFormula) 'Date edit replaced predecessor validation'
    Assert-True ($excel.Calculation -eq -4135 -and $excel.EnableEvents) 'Fast edit did not restore application state'
    $wbs.Range('F8').Value2 = 5.0
    $excel.Calculation = -4105

    # Add a peer after the entire project subtree, then add two nested tasks.
    $wbs.Range('C6').Select()
    Run-Macro 'Model_1.AddTask'
    $newRow = $excel.Selection.Row
    Assert-True ($newRow -eq 12 -and $wbs.Cells.Item($newRow, 25).Value2 -eq 1) 'Add Task split the parent subtree'
    Run-Macro 'Model_1.AddSubtask'
    Assert-True ($wbs.Cells.Item(13,25).Value2 -eq 2) 'Level 2 task not created'
    Run-Macro 'Model_1.AddSubtask'
    Assert-True ($wbs.Cells.Item(14,25).Value2 -eq 3) 'Level 3 task not created'
    $wbs.Range('G14').Value2 = 20180831.0
    $wbs.Range('F14').Value2 = 2.0
    Assert-Date $wbs.Range('H12') ([datetime]'2018-09-01') 'New three-level parent did not roll up'
    $wbs.Range('C14').Select()
    $wbs.Range('D2').Value2 = 2.0
    Assert-True ($wbs.Range('Y14').Value2 -eq 2) 'Level selector does not update the selected task'
    Run-Macro 'Model_1.AddSubtask'
    Run-Macro 'Model_1.AddSubtask'
    Run-Macro 'Model_1.AddSubtask'
    Assert-True ($wbs.Cells.Item($excel.Selection.Row, 25).Value2 -eq 5) 'Five-level hierarchy is unavailable'

    Run-Macro 'Model_1.UpdateGanttChart'
    $first = $excel.WorksheetFunction.Min($wbs.Range('G6:G30'))
    $last = $excel.WorksheetFunction.Max($wbs.Range('H6:H30'))
    $lastCol = 26 + [int]($last - $first)
    for ($c = 26; $c -le $lastCol; $c++) {
        Assert-True ($wbs.Cells.Item(4,$c).Value2 -eq $first + $c - 26) 'Daily timeline has missing or stale dates'
    }
    $satCol = 26 + [int](([datetime]'2018-09-01').ToOADate() - $first)
    Assert-True ($wbs.Cells.Item(5,$satCol).Value2 -eq 'Sat') 'Weekdays must use English labels'
    Assert-True ($wbs.Cells.Item(5,$satCol).Interior.Color -ne $wbs.Cells.Item(5,$satCol-1).Interior.Color) 'Weekend shading is missing'
    Assert-True ($null -eq $wbs.Range('Z3').Value2 -and $null -eq $wbs.Range('AD1').Value2) 'Stray top dates remain'

    # Import remains linked even in manual calculation mode and after row inserts.
    $pg.Range('B14:J213').ClearContents()
    $wbs.Range('C8:C9').Select()
    Run-Macro 'ProgressChart.ImportSelectedMilestones'
    Assert-True ($pg.Range('C14').Formula -eq '=ROW(WBS!C8)') 'WBS import row link missing'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 2
    $excel.Calculation = -4135
    $wbs.Range('H8').Value2 = 20180905.0
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Date $pg.Range('E14') ([datetime]'2018-09-05') 'Linked date is stale in manual calculation'
    $wbs.Rows.Item(8).Insert() | Out-Null
    $pg.Calculate()
    Assert-True ($pg.Range('C14').Value2 -eq 9) 'Row insertion broke linked imports'
    $excel.Calculation = -4105

    $pg.Range('B14:J213').ClearContents()
    Fill-Node 14 'Later milestone' ([datetime]'2027-01-02')
    Fill-Node 15 'Earlier milestone' ([datetime]'2026-12-30')
    Fill-Node 16 'Same-day milestone with an unusually long descriptive English label' ([datetime]'2026-12-30')
    $pg.Range('C6').Value2 = 'Date proportional'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 3
    $oldId = $pg.Shapes.Item('WBS_ProgressChart').ID
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 3
    Assert-True ($pg.Shapes.Item('WBS_ProgressChart').ID -ne $oldId) 'Refresh retained old chart'
    Run-Macro 'ProgressChart.NarrowProgressChart'
    Assert-Chart 3
    Run-Macro 'ProgressChart.WidenProgressChart'
    Assert-Chart 3

    $pg.Range('B14:J213').ClearContents()
    for ($i=0; $i -lt 60; $i++) { Fill-Node (14+$i) "Dense same-day milestone $i" ([datetime]'2026-08-20') }
    $pg.Range('C7').Value2 = 0.5
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 60
    $pg.Range('C6').Value2 = 'Equal spacing'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 60
    $pg.Range('B14:J213').ClearContents()
    Fill-Node 14 'Single node' ([datetime]'2026-08-20')
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 1

    # Shrinking a plan must clear previously drawn dates, bars and merges.
    $wbs.Activate()
    $excel.EnableEvents = $false
    $wbs.Range('B6:Y100').ClearContents()
    $wbs.Range('C6').Value2 = 'Short plan'
    $wbs.Range('G6').Value2 = ([datetime]'2018-08-31').ToOADate()
    $wbs.Range('F6').Value2 = 2.0
    $excel.EnableEvents = $true
    Run-Macro 'Model_1.UpdateGanttChart'
    Assert-True ($excel.WorksheetFunction.CountA($wbs.Range('AB4:HD100')) -eq 0) 'Old timeline or bars remain after shrinking'
    Assert-True $excel.EnableEvents 'Worksheet events were left disabled'
    # Adjacent and non-contiguous selection, entire rows, and overlapping subtrees.
    $excel.EnableEvents = $false
    $wbs.Range('B6:Y100').ClearContents()
    $batchNames = @('Project', '    Phase A', '        Leaf A', '        Leaf B', '    Phase B', '        Leaf C', '        Leaf D', '    Closing')
    for ($i = 0; $i -lt $batchNames.Count; $i++) { $wbs.Cells.Item(6+$i,3).Value2 = $batchNames[$i] }
    foreach ($r in @(8,9,11,12,13)) {
        $wbs.Cells.Item($r,7).Value2 = ([datetime]'2026-09-01').ToOADate()
        $wbs.Cells.Item($r,6).Value2 = 2.0
        $wbs.Cells.Item($r,14).Value2 = 0.25
    }
    $excel.EnableEvents = $true
    Run-Macro 'Model_1.RefreshParentTasks'
    $wbs.Range('C8:C9').Select()
    $wbs.Range('D2').Select()
    $wbs.Range('D2').Value2 = 2.0
    Assert-True ($wbs.Range('Y8').Value2 -eq 2 -and $wbs.Range('Y9').Value2 -eq 2) 'Dropdown changed only the first selected task'
    Run-Macro 'ProgressChart.DemoteTasks'
    Assert-True ($wbs.Range('Y8').Value2 -eq 3 -and $wbs.Range('Y9').Value2 -eq 3) 'Demote after dropdown lost selected tasks'
    $wbs.Range('C8,C11').Select()
    Run-Macro 'ProgressChart.PromoteTasks'
    Assert-True ($wbs.Range('Y8').Value2 -eq 2 -and $wbs.Range('Y9').Value2 -eq 3 -and $wbs.Range('Y11').Value2 -eq 2) 'Ctrl-selection promotion failed'
    $wbs.Range('C8:C9,C11').Select()
    Run-Macro 'ProgressChart.DemoteTasks'
    Assert-True ($wbs.Range('Y8').Value2 -eq 3 -and $wbs.Range('Y9').Value2 -eq 4 -and $wbs.Range('Y11').Value2 -eq 3) 'Overlapping parent/child selection moved a child twice'
    $wbs.Rows.Item(9).Select()
    Run-Macro 'ProgressChart.PromoteTasks'
    Assert-True ($wbs.Range('Y9').Value2 -eq 3) 'Whole-row promotion failed'
    $wbs.Range('C11:C12').Select()
    $wbs.Range('D2').Select()
    $wbs.Range('D2').Value2 = 2.0
    Assert-True ($wbs.Range('Y11').Value2 -eq 2 -and $wbs.Range('Y12').Value2 -eq 3) 'Dropdown did not preserve the selected parent subtree'

    # Raw predecessor input becomes a formula and survives renumbering.
    $wbs.Range('O13').Value2 = $wbs.Range('C9').Value2.Replace(' ', '')
    Assert-True ($wbs.Range('O13').Formula -eq '=C9') 'Predecessor text was not linked'
    $wbs.Range('C7:C9').Select()
    Run-Macro 'ProgressChart.PromoteTasks'
    Assert-True ($wbs.Range('O13').Formula -eq '=C9') 'Level change broke predecessor formula'
    $wbs.Range('F9').Value2 = 4.0
    Assert-Date $wbs.Range('G13') ([datetime]'2026-09-05') 'Predecessor stopped updating after renumbering'

    # Invalid batches and scheduling cycles raise errors without partial updates.
    $excel.EnableEvents = $false
    $savedLevels = ($wbs.Range('Y6:Y13').Value2 | Out-String)
    $rejected = $false
    try { $excel.Run("'$($book.Name)'!Model_1.SetTaskLevels", $wbs.Range('C6'), 0, -1) | Out-Null } catch { $rejected = $true }
    Assert-True $rejected 'Promoting the first task should fail'
    Assert-True (($wbs.Range('Y6:Y13').Value2 | Out-String) -eq $savedLevels) 'Rejected batch partially changed levels'
    $rejected = $false
    try { $excel.Run("'$($book.Name)'!Model_1.SetTaskLevels", $wbs.Range('C8:C9'), 5, 0) | Out-Null } catch { $rejected = $true }
    Assert-True $rejected 'Skipping a parent level should fail'
    Assert-True (($wbs.Range('Y6:Y13').Value2 | Out-String) -eq $savedLevels) 'Skipped-level batch partially changed tasks'
    $oldEnd = $wbs.Range('H9').Value2
    $wbs.Range('O9').Formula = '=C13'
    $rejected = $false
    try { Run-Macro 'Model_1.RefreshParentTasks' } catch { $rejected = $true }
    Assert-True $rejected 'Circular predecessors should fail'
    Assert-True ($wbs.Range('H9').Value2 -eq $oldEnd -and -not $excel.EnableEvents) 'Cycle wrote a partial schedule or changed caller event state'
    $wbs.Range('O9').ClearContents()
    Run-Macro 'Model_1.RefreshParentTasks'
    $excel.EnableEvents = $true

    # Long predecessor chain: no VBA recursion and one data edit propagates to the end.
    $excel.EnableEvents = $false
    $wbs.Range('B6:Y3010').ClearContents()
    $data = New-Object 'object[,]' 3000,13
    for ($i = 0; $i -lt 3000; $i++) {
        $data[$i,0] = "Task $i"
        $data[$i,3] = 1.0
        $data[$i,4] = ([datetime]'2026-01-01').ToOADate()
        $data[$i,11] = 0.5
        if ($i -gt 0) { $data[$i,12] = "Task $($i-1)" }
    }
    $wbs.Range('C6:O3005').Value2 = $data
    Run-Macro 'Model_1.RefreshParentTasks'
    $excel.EnableEvents = $true
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $wbs.Range('G6').Value2 = 20260102.0
    $timer.Stop()
    Assert-Date $wbs.Range('H3005') (([datetime]'2026-01-02').AddDays(2999)) 'Long predecessor chain failed'
    Assert-True $excel.EnableEvents 'Long-chain edit left events disabled'
    Write-Output "3000-task date edit: $($timer.ElapsedMilliseconds) ms (local Excel measurement)."
    Write-Output 'Excel smoke checks passed. See USAGE.md for manual invalid-input and clipboard checks.'
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) {
        $excel.EnableEvents = $true
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
