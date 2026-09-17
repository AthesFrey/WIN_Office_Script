# Windows + Microsoft Excel integration checks. Run in an Excel trusted location:
# pwsh -File tests/excel_smoke.ps1 -WorkbookPath ./dist/WBS-Gantt-ProgressChart-v5.6fix3.xlsm
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
function Send-ExcelKeys([string]$Keys) {
    # SendKeys needs foreground focus; wait for Excel to consume the queued input.
    [Microsoft.VisualBasic.Interaction]::AppActivate([string]$excel.Caption)
    $excel.SendKeys($Keys, $true)
    Start-Sleep -Milliseconds 300
}
function Assert-Date([object]$Cell, [datetime]$Date, [string]$Message) {
    Assert-True ([math]::Abs([double]$Cell.Value2 - $Date.ToOADate()) -lt 0.001) $Message
    Assert-True ($Cell.NumberFormat -eq 'yyyy/m/d') "Incorrect date format: $Message"
    Assert-True ($Cell.Text -eq $excel.WorksheetFunction.Text($Date.ToOADate(), 'yyyy/m/d')) "Date displays a serial or is clipped: $Message"
}
function Assert-Progress([object]$Cell, [double]$Value, [string]$Message) {
    Assert-True ([math]::Abs([double]$Cell.Value2 - $Value) -lt 1e-12) "Progress lost precision: $Message"
    Assert-True ($Cell.NumberFormat -eq '0.00%') "Incorrect progress format: $Message"
    Assert-True ($Cell.Text -eq $excel.WorksheetFunction.Text($Value, '0.00%')) "Progress does not display two percentage decimals: $Message"
}
function Fill-Node([int]$Row, [string]$Label, [datetime]$Date) {
    $pg.Cells.Item($Row, 2).Value2 = 'Yes'
    $pg.Cells.Item($Row, 4).Value2 = $Label
    $pg.Cells.Item($Row, 5).Value2 = $Date.ToOADate()
}
function Assert-Chart([int]$NodeCount, [int]$InputCount = $NodeCount) {
    $chart = $pg.Shapes.Item('WBS_ProgressChart')
    Assert-True ($chart.Type -eq 6) 'Chart is not a shape group'
    Assert-True ($pg.Shapes.Count -eq 12) 'Duplicate chart or temporary shapes remain'
    Assert-True ($pg.Range('B10').Value2.StartsWith("Generated $InputCount milestones;")) 'Generation did not finish'
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
function Assert-NoMilestoneChart {
    for ($i = 1; $i -le $pg.Shapes.Count; $i++) {
        $name = $pg.Shapes.Item($i).Name
        Assert-True ($name -notmatch '^(WBS_ProgressChart|WBS_PG_TMP_|PG_Background|PG_Axis|PG_Marker_|PG_Date_|PG_Label_)') "Old or partial milestone drawing remains: $name"
    }
    foreach ($name in @('GenerateProgressChart', 'CopyProgressChart', 'WidenProgressChart', 'NarrowProgressChart',
                         'UndoMilestones', 'RedoMilestones', 'DeleteMilestoneLevel1', 'DeleteMilestoneLevel2',
                         'DeleteMilestoneLevel3', 'DeleteMilestoneLevel4', 'DeleteMilestoneLevel5')) {
        Assert-True ($pg.Shapes.Item("PG_Button_$name").OnAction.Length -gt 0) "Cleanup removed the $name button"
    }
}
function Assert-Font([object]$Range, [string]$Message) {
    Assert-True ($Range.Font.Name -eq 'Microsoft YaHei' -and [math]::Abs([double]$Range.Font.Size - 11) -lt 0.01) $Message
}
function Get-ChartLabels {
    $group = $pg.Shapes.Item('WBS_ProgressChart')
    for ($i = 1; $i -le $group.GroupItems.Count; $i++) {
        $shape = $group.GroupItems.Item($i)
        if ($shape.Type -eq 17) {
            $text = $shape.TextFrame2.TextRange.Text -replace '[\r\n]', ''
            if ($text -notmatch '^\d{1,2}/\d{1,2}$') { $text }
        }
    }
}
function Get-RangeState([object]$Range) {
    $rows = foreach ($cell in $Range.Cells) {
        @($cell.Address(), $cell.Value2, $cell.Formula, $cell.HasFormula, $cell.NumberFormat, $cell.Font.Name,
          $cell.Font.Size, $cell.Font.Bold, $cell.Font.Italic, $cell.Font.Color,
          $cell.Font.Underline, $cell.Font.Strikethrough, $cell.WrapText, $cell.RowHeight,
          $cell.HorizontalAlignment, $cell.VerticalAlignment, $cell.Orientation, $cell.ShrinkToFit, $cell.IndentLevel,
          $cell.Borders.Item(7).LineStyle, $cell.Borders.Item(8).LineStyle,
          $cell.Borders.Item(9).LineStyle, $cell.Borders.Item(10).LineStyle,
          $cell.Interior.Pattern, $cell.Interior.Color, $cell.DisplayFormat.Interior.Color) | ConvertTo-Json -Compress
    }
    $rows -join "`n"
}
function Invoke-History([bool]$Redo) {
    $oldCalculation = $excel.Calculation
    $oldScreen = $excel.ScreenUpdating
    $name = if ($Redo) { 'RedoWbs' } else { 'UndoWbs' }
    $countName = if ($Redo) { 'RedoCount' } else { 'UndoCount' }
    $count = [int]$excel.Run("'$($book.Name.Replace("'", "''"))'!Model_1.$countName")
    Assert-True ($count -gt 0) "$name has no available history"
    $button = $wbs.Shapes.Item("PG_Button_$name")
    Assert-True ($button.Fill.ForeColor.RGB -eq 15426341) "$name unexpectedly grey"
    # Invoke the exact macro assigned to the worksheet button.
    $excel.Run($button.OnAction) | Out-Null
    Assert-True $excel.EnableEvents "$name left worksheet events disabled"
    Assert-True ($excel.Calculation -eq $oldCalculation -and $excel.ScreenUpdating -eq $oldScreen) "$name changed application settings"
    Assert-True (-not $excel.CommandBars.GetEnabledMso('Undo')) 'Native Undo must remain disabled'
    Assert-True (-not $excel.CommandBars.GetEnabledMso('Redo')) 'Native Redo must remain disabled'
}
function Undo-Wbs { Invoke-History $false }
function Redo-Wbs { Invoke-History $true }
function Assert-History([int]$Undo, [int]$Redo) {
    $prefix = "'$($book.Name.Replace("'", "''"))'!Model_1."
    Assert-True ([int]$excel.Run($prefix + 'UndoCount') -eq $Undo) "Expected $Undo undo steps"
    Assert-True ([int]$excel.Run($prefix + 'RedoCount') -eq $Redo) "Expected $Redo redo steps"
    Assert-True ($wbs.Shapes.Item('PG_Button_UndoWbs').TextFrame2.TextRange.Text -eq "Undo ($Undo)") 'Undo count label is stale'
    Assert-True ($wbs.Shapes.Item('PG_Button_RedoWbs').TextFrame2.TextRange.Text -eq "Redo ($Redo)") 'Redo count label is stale'
}
function Assert-RoundTrip([object]$Range, [string]$Before) {
    $after = Get-RangeState $Range
    Undo-Wbs
    Assert-True ((Get-RangeState $Range) -eq $Before) 'Undo did not restore the full before state'
    $wbs.Range('D6').Select()
    Redo-Wbs
    Assert-True ((Get-RangeState $Range) -eq $after) 'Redo did not restore the full after state'
    Undo-Wbs
    Assert-True ((Get-RangeState $Range) -eq $Before) 'Undo after Redo did not round-trip'
}
function Assert-MilestoneHistory([int]$Undo, [int]$Redo) {
    $prefix = "'$($book.Name.Replace("'", "''"))'!Sheet4."
    Assert-True ([int]$excel.Run($prefix + 'UndoCount') -eq $Undo) "Expected $Undo Milestones undo steps"
    Assert-True ([int]$excel.Run($prefix + 'RedoCount') -eq $Redo) "Expected $Redo Milestones redo steps"
    Assert-True ($pg.Shapes.Item('PG_Button_UndoMilestones').TextFrame2.TextRange.Text -eq "Undo ($Undo)") 'Milestones Undo label is stale'
    Assert-True ($pg.Shapes.Item('PG_Button_RedoMilestones').TextFrame2.TextRange.Text -eq "Redo ($Redo)") 'Milestones Redo label is stale'
}
function Invoke-MilestoneHistory([bool]$Redo) {
    $oldCalculation = $excel.Calculation
    $oldScreen = $excel.ScreenUpdating
    $name = if ($Redo) { 'RedoMilestones' } else { 'UndoMilestones' }
    $button = $pg.Shapes.Item("PG_Button_$name")
    Assert-True ($button.Fill.ForeColor.RGB -eq 15426341) 'Available Milestones history button is not blue'
    $binding = [string]$button.OnAction
    Assert-True ($binding -match "ProgressChart\.$name$") 'Milestones history binding is missing'
    $excel.Run($binding) | Out-Null
    Assert-True $excel.EnableEvents 'Milestones history left events disabled'
    Assert-True ($excel.Calculation -eq $oldCalculation -and $excel.ScreenUpdating -eq $oldScreen) 'Milestones history changed application settings'
}
function Undo-Milestones { Invoke-MilestoneHistory $false }
function Redo-Milestones { Invoke-MilestoneHistory $true }
. (Join-Path $PSScriptRoot 'excel_milestones.ps1')
. (Join-Path $PSScriptRoot 'excel_regressions.ps1')
try {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $book = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path, 0, $true)
    Assert-True ($book.Worksheets.Count -eq 2) 'Workbook should contain only two sheets'
    $wbs = $book.Worksheets.Item('WBS')
    $pg = $book.Worksheets.Item('Milestones')
    Assert-MilestonesToolbar -CheckReservedRows
    $wbs.Activate()
    Assert-Font $wbs.Range('B6:BG19') 'Delivered body must be YaHei 11 before refresh'
    Assert-History 0 0
    Assert-MilestoneHistory 0 0
    Assert-True (-not $excel.CommandBars.GetEnabledMso('Undo')) 'Native Undo not disabled on open'
    Assert-True (-not $excel.CommandBars.GetEnabledMso('Redo')) 'Native Redo not disabled on open'
    # v5.5: reproduce a body edit on the delivered sample before any refresh.
    Assert-Date $wbs.Range('G6') ([datetime]'2018-08-01') 'Original sample start'
    Assert-Date $wbs.Range('G15') ([datetime]'2018-08-22') 'Original sample later start'
    Assert-Progress $wbs.Range('N6') 0.6441176470588236 'Saved parent progress'
    Assert-Progress $wbs.Range('AB9') 1.0 'Saved one-day Gantt progress'
    Assert-True $wbs.Range('AB9').ShrinkToFit 'One-day progress should fit its bar'
    $wbs.Range('C8').Select()
    $sampleBefore = Get-RangeState $wbs.Range('B6:Y19')
    $wbs.Range('C8').Value2 = '        Edited task text'
    $wbs.Range('D6').Select()
    Assert-RoundTrip $wbs.Range('B6:Y19') $sampleBefore
    Assert-True ((Get-RangeState $wbs.Range('B6:Y19')) -eq $sampleBefore) 'Sample text-edit Undo changed body values or formats'
    Assert-Date $wbs.Range('G6') ([datetime]'2018-08-01') 'Undo exposed serial 43313'
    Assert-Date $wbs.Range('G15') ([datetime]'2018-08-22') 'Undo exposed serial 43334'
    Assert-Progress $wbs.Range('N6') 0.6441176470588236 'Undo exposed long parent decimals'
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
    Assert-Progress $wbs.Cells.Item($newRow,14) 0.0 'New task progress'
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
    $barCol = 26 + [int]($wbs.Range('G8').Value2 - $first)
    Assert-Progress $wbs.Cells.Item(8,$barCol) $wbs.Range('N8').Value2 'Refreshed Gantt progress'
    Assert-True $wbs.Cells.Item(8,$barCol).ShrinkToFit 'Refreshed progress should fit its bar'

    # Import remains linked even in manual calculation mode and after row inserts.
    $pg.Range('B14:J213').ClearContents()
    $wbs.Range('C8:C9').Select()
    Run-Macro 'ProgressChart.ImportSelectedMilestones'
    Assert-True ($pg.Range('C14').Formula -eq '=ROW(WBS!C8)') 'WBS import row link missing'
    $pg.Calculate()
    Assert-Progress $pg.Range('H14') $wbs.Range('N8').Value2 'Imported progress'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 2
    $excel.Calculation = -4135
    $wbs.Range('H8').Value2 = 20180905.0
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Date $pg.Range('E14') ([datetime]'2018-09-05') 'Linked date is stale in manual calculation'
    Assert-Progress $pg.Range('H14') $wbs.Range('N8').Value2 'Refreshed milestone progress'
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
    Assert-Chart 2 3
    $mergedLabels = @(Get-ChartLabels | Where-Object { $_ -like '*Earlier milestone*' })
    Assert-True ($mergedLabels.Count -eq 1 -and $mergedLabels[0] -eq 'Earlier milestone; Same-day milestone with an unusually long descriptive English label') 'Same-day merge lost names, separator or input order'
    $oldId = $pg.Shapes.Item('WBS_ProgressChart').ID
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 2 3
    Assert-True ($pg.Shapes.Item('WBS_ProgressChart').ID -ne $oldId) 'Refresh retained old chart'
    Run-Macro 'ProgressChart.NarrowProgressChart'
    Assert-Chart 2 3
    Run-Macro 'ProgressChart.WidenProgressChart'
    Assert-Chart 2 3

    $pg.Range('B14:J213').ClearContents()
    for ($i=0; $i -lt 60; $i++) { Fill-Node (14+$i) ("Dense same-day milestone $i " + ('x' * 100)).Substring(0, 120) ([datetime]'2026-08-20') }
    $pg.Range('C7').Value2 = 0.5
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 1 60
    $pg.Range('C6').Value2 = 'Equal spacing'
    Run-Macro 'ProgressChart.GenerateProgressChart'
    Assert-Chart 1 60
    Assert-True ((@(Get-ChartLabels)[0] -split '; ').Count -eq 60) 'The 60 input names were truncated'
    $oldId = $pg.Shapes.Item('WBS_ProgressChart').ID
    Fill-Node 74 'Rejected 61st same-day input' ([datetime]'2026-08-20')
    $rejected = $false
    try { $excel.Run("'$($book.Name)'!ProgressChart.GenerateProgressChart", $true) | Out-Null } catch { $rejected = $true }
    Assert-True $rejected 'The limit was incorrectly applied to unique dates instead of input tasks'
    Assert-NoMilestoneChart
    Assert-True ($pg.Range('B10').Value2.StartsWith('Milestones skipped after retry:')) 'Failed chart refresh did not report the skipped result'
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

    # v5.5: warnings inspect the final, calculated schedule through both buttons.
    $wbs.Activate()
    $excel.EnableEvents = $false
    $wbs.Range('B6:Y100').ClearContents()
    $warnNames = @('Project', '    A', '    B', '    Parent', '        Grandchild', '    C', '', '    D', 'Other project', '    E', '    F', 'Top A', 'Top B')
    for ($i = 0; $i -lt $warnNames.Count; $i++) { $wbs.Cells.Item(6+$i,3).Value2 = $warnNames[$i] }
    $warnRows = @(7,8,10,11,13,15,16,17,18)
    $warnOffsets = @(0,4,9,14,19,0,4,0,4)
    for ($i = 0; $i -lt $warnRows.Count; $i++) {
        $wbs.Cells.Item($warnRows[$i],7).Value2 = ([datetime]'2026-09-01').AddDays($warnOffsets[$i]).ToOADate()
        $wbs.Cells.Item($warnRows[$i],6).Value2 = 2.0
    }
    # A calculated duration changes after H is written; verify the final mismatch is caught.
    # Ordinary fixed numeric Days retain v5.1's automatic end-date calculation.
    $wbs.Range('F7').Formula = '=IF(H7=DATE(2026,9,2),3,2)'
    $wbs.Range('F7').Interior.Color = 5287936
    $wbs.Range('G8').Interior.Color = 12611584
    $excel.EnableEvents = $true
    Run-Macro 'ProgressChart.UpdateGantt'
    $red = 13551615  # RGB(255,199,206)
    $yellow = 10284031  # RGB(255,235,156)
    Assert-True ($wbs.Range('F7').DisplayFormat.Interior.Color -eq $red) 'Final Days mismatch is not red'
    Assert-True ($wbs.Range('F7').Interior.Color -eq 5287936) 'Warning overwrote the custom underlying fill'
    foreach ($r in @(8,16,18)) { Assert-True ($wbs.Cells.Item($r,7).DisplayFormat.Interior.Color -eq $yellow) "Adjacent sibling gap is not yellow: $r" }
    foreach ($r in @(6,7,10,11,13,14,15,17)) { Assert-True ($wbs.Cells.Item($r,7).DisplayFormat.Interior.Color -ne $yellow) "First/nonadjacent sibling incorrectly yellow: $r" }
    $wbs.Range('F7').Value2 = 2.0
    $wbs.Range('G8').Value2 = ([datetime]'2026-09-03').ToOADate()
    $pg.Range('B14:J213').ClearContents()
    Fill-Node 14 'Formatting check' ([datetime]'2026-09-20')
    Run-Macro 'ProgressChart.RefreshAll'
    Assert-True ($wbs.Range('F7').DisplayFormat.Interior.Color -eq 5287936) 'Clearing red warning lost original fill'
    Assert-True ($wbs.Range('G8').DisplayFormat.Interior.Color -eq 12611584) 'Clearing yellow warning lost original fill'
    $ganttState = Get-RangeState $wbs.Range('Z4:AS18')
    Run-Macro 'ProgressChart.UpdateGantt'
    Assert-True ((Get-RangeState $wbs.Range('Z4:AS18')) -eq $ganttState) 'Refresh All and Update Gantt produce different Gantt output'
    $titleFont = $wbs.Range('B1').Font.Size
    $headerFont = $pg.Range('B13').Font.Size
    $wbs.Range('B6:AS18').Font.Name = 'Arial'
    $wbs.Range('B6:AS18').Font.Size = 9
    $pg.Range('B3:J10,B14:J213').Font.Name = 'Arial'
    $pg.Range('B3:J10,B14:J213').Font.Size = 9
    Run-Macro 'ProgressChart.RefreshAll'
    Assert-Font $wbs.Range('B6:AS18') 'WBS/Gantt body font is not YaHei 11'
    Assert-Font $pg.Range('B3:J10') 'Milestones settings font is not YaHei 11'
    Assert-Font $pg.Range('B14:J213') 'Milestones input font is not YaHei 11'
    Assert-True ($wbs.Range('B1').Font.Size -eq $titleFont -and $pg.Range('B13').Font.Size -eq $headerFont) 'Body formatting changed a title/header'
    $chart = $pg.Shapes.Item('WBS_ProgressChart')
    for ($i=1; $i -le $chart.GroupItems.Count; $i++) {
        $shape = $chart.GroupItems.Item($i)
        if ($shape.Type -eq 17) { Assert-True ($shape.TextFrame2.TextRange.Font.Size -eq 10) 'Body formatting changed chart text' }
    }
    $excel.EnableEvents = $false
    $wbs.Range('G8').Value2 = 'unreadable'
    Run-Macro 'Model_1.CheckScheduleWarnings'
    Assert-True ($wbs.Range('G8').DisplayFormat.Interior.Color -ne $yellow) 'Unreadable start retained a warning'
    $wbs.Range('B6:Y100').ClearContents()
    Run-Macro 'Model_1.CheckScheduleWarnings'
    Assert-True ($wbs.Range('G16').DisplayFormat.Interior.Color -ne $yellow) 'Empty plan retained stale warning rules'

    # Undo restores edits AND all derived data, without moving worksheet rows.
    $undoNames = @('Project', '    First', '    Second', 'Last task')
    for ($i=0; $i -lt $undoNames.Count; $i++) { $wbs.Cells.Item(6+$i,3).Value2 = $undoNames[$i] }
    foreach ($r in @(7,8,9)) {
        $wbs.Cells.Item($r,6).Value2 = 2.0
        $wbs.Cells.Item($r,7).Value2 = ([datetime]'2026-09-01').ToOADate()
        $wbs.Cells.Item($r,14).Value2 = 0.5
    }
    $wbs.Range('O8').Formula = '=C7'
    $wbs.Range('D7').NumberFormat = '@'
    $wbs.Range('D7').Value2 = '=literal text'
    $wbs.Range('D8').NumberFormat = '@'
    $wbs.Range('D8').Value2 = '00123'
    $wbs.Range('D8').NumberFormat = 'General'
    $wbs.Range('D9').NumberFormat = '@'
    $wbs.Range('D9').Value2 = '2018/8/22'
    $wbs.Range('D10').NumberFormat = '@'
    $wbs.Range('D10').Value2 = '=literal in General'
    $wbs.Range('D10').NumberFormat = 'General'
    $wbs.Range('L7').NumberFormat = 'General'
    $wbs.Range('L7').Formula = '=1/3'
    $wbs.Range('L7').NumberFormat = '@'
    $wbs.Range('I7').NumberFormat = 'yyyy-mm-dd'
    $wbs.Range('I7').Value2 = ([datetime]'2018-08-01').ToOADate()
    $wbs.Range('P8').NumberFormat = 'm/d/yyyy'
    $wbs.Range('P8').Value2 = ([datetime]'2018-08-22').ToOADate()
    $excel.EnableEvents = $true
    Run-Macro 'Model_1.RefreshParentTasks'
    $wbs.Activate()
    $wbs.Range('F7').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $wbs.Range('F7').Value2 = 4.0
    Assert-True ($wbs.Range('F6').Value2 -eq 6) 'Undo fixture did not propagate to parent and predecessor'
    $wbs.Range('C8').Select()  # Selection/D2 updates must not clear history.
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Undo did not restore linked values/formats'
    Assert-True (-not $wbs.Range('D7').HasFormula -and $wbs.Range('D7').Value2 -eq '=literal text') 'Undo converted literal text to a formula'
    Assert-True ($wbs.Range('D8').Value2 -is [string] -and $wbs.Range('D8').Value2 -ceq '00123') 'Undo stripped leading zeros'
    Assert-True ($wbs.Range('D9').Value2 -is [string] -and $wbs.Range('D9').Value2 -ceq '2018/8/22') 'Undo parsed literal date text'
    Assert-True (-not $wbs.Range('D10').HasFormula -and $wbs.Range('D10').Value2 -eq '=literal in General') 'Undo parsed a General-format literal'
    Assert-True ($wbs.Range('L7').HasFormula -and $wbs.Range('L7').NumberFormat -eq '@') 'Undo lost a formula with a Text display format'
    Assert-True ([math]::Abs($wbs.Range('L7').Value2 - 1.0/3.0) -lt 1e-12) 'Undo rounded a formula result'
    Assert-True ($excel.WorksheetFunction.CountA($wbs.Range('C10:C1000')) -eq 0) 'Undo filled originally empty cells with text'
    $wbs.Range('F7').Select()
    $wbs.Range('F7').Value2 = 3.0
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $wbs.Range('F7').Value2 = 5.0  # Same selected cell, no SelectionChange between edits.
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Consecutive edits used a stale snapshot'

    $wbs.Range('N7').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $wbs.Range('N7').Value2 = 1.0/3.0
    Assert-Progress $wbs.Range('N7') (1.0/3.0) 'Edited fractional progress'
    $parentProgress = ($wbs.Range('F7').Value2 / 3.0 + $wbs.Range('F8').Value2 * $wbs.Range('N8').Value2) / ($wbs.Range('F7').Value2 + $wbs.Range('F8').Value2)
    Assert-Progress $wbs.Range('N6') $parentProgress 'Weighted progress after edit'
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Progress Undo lost precision or formatting'

    $paste = New-Object 'object[,]' 2,3
    for ($i=0; $i -lt 2; $i++) {
        $paste[$i,0] = 4.0
        $paste[$i,1] = ([datetime]'2026-09-10').ToOADate()
        $paste[$i,2] = ([datetime]'2026-09-13').ToOADate()
    }
    $pg.Range('L220:N221').Value2 = $paste
    $wbs.Range('F7:H8').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $pg.Range('L220:N221').Copy()
    $wbs.Range('F7:H8').PasteSpecial(-4163) | Out-Null  # Actual multi-cell values paste.
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Multi-cell paste Undo missed linked changes'
    $wbs.Range('F7:H8').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $pg.Range('L220:N221').NumberFormat = 'General'
    $pg.Range('L220:N221').Font.Italic = $true
    $pg.Range('L220:N221').Font.Color = 255
    $pg.Range('L220:N221').VerticalAlignment = -4160
    $pg.Range('L220:N221').Interior.Color = 65535
    $pg.Range('L220:N221').Borders.LineStyle = -4118
    $pg.Range('L220:N221').Copy()
    $wbs.Range('F7:H8').PasteSpecial(-4104) | Out-Null  # Paste values and foreign formats.
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Paste-all Undo did not restore mixed number formats'
    $wbs.Range('C8').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    Run-Macro 'ProgressChart.PromoteTasks'
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Level-button Undo missed numbering, parent formatting or predecessor changes'
    $wbs.Range('C8').Select()
    $wbs.Range('D2').Select()
    $wbs.Range('D2').Value2 = 1.0
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Level-dropdown Undo missed linked changes'
    $pg.Range('C14').Formula = '=ROW(WBS!C9)'
    $wbs.Range('C9').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $validationBefore = $wbs.Range('O7').Validation.Formula1
    $wbs.Range('C9').ClearContents()
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Clearing final task was not undone'
    Assert-True ($pg.Range('C14').Formula -eq '=ROW(WBS!C9)') 'Content Undo moved linked source rows'
    Assert-True ($wbs.Range('O7').Validation.Formula1 -eq $validationBefore) 'Undo did not restore predecessor validation'
    $wbs.Range('D7').Select()
    $before = Get-RangeState $wbs.Range('B6:Y10')
    $wbs.Range('D7').Value2 = 'New owner'
    $wbs.Range('C6').Select()
    Assert-RoundTrip $wbs.Range('B6:Y10') $before
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $before) 'Body edits outside scheduling columns lost Undo'

    # Linked milestone cells must recalculate on replay even with manual calculation.
    $pg.Range('H14').Formula = '=WBS!N7'
    $wbs.Range('N7').Select()
    $linkedProgress = [double]$wbs.Range('N7').Value2
    $excel.Calculation = -4135
    $wbs.Range('N7').Value2 = 0.75
    Undo-Wbs
    Assert-True ([math]::Abs([double]$pg.Range('H14').Value2 - $linkedProgress) -lt 1e-12) 'Undo left a linked milestone stale'
    Redo-Wbs
    Assert-True ([math]::Abs([double]$pg.Range('H14').Value2 - 0.75) -lt 1e-12) 'Redo left a linked milestone stale'
    Undo-Wbs
    $excel.Calculation = -4105

    # Three-step capacity, both directions, overflow and branching.
    Run-Macro 'Model_1.ResetUndoTracking'
    $states = @((Get-RangeState $wbs.Range('B6:Y10')))
    for ($i=1; $i -le 4; $i++) {
        $wbs.Range('D7').Value2 = "History owner $i"
        $states += Get-RangeState $wbs.Range('B6:Y10')
        Assert-History ([math]::Min($i, 3)) 0
    }
    for ($i=3; $i -ge 1; $i--) {
        Undo-Wbs
        Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $states[$i]) 'Three-step Undo used the wrong state'
    }
    Assert-History 0 3
    Run-Macro 'ProgressChart.UndoWbs'  # Disabled action must be a no-op.
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $states[1]) 'History exceeded its three-step limit'
    for ($i=2; $i -le 4; $i++) {
        Redo-Wbs
        Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $states[$i]) 'Three-step Redo used the wrong state'
    }
    Assert-History 3 0
    Undo-Wbs
    Undo-Wbs
    $pg.Activate()
    $wbs.Activate()
    Assert-History 1 2
    $wbs.Range('D7').Value2 = 'New history branch'
    Assert-History 2 0
    Run-Macro 'ProgressChart.RedoWbs'
    Assert-True ($wbs.Range('D7').Value2 -eq 'New history branch') 'Old redo branch survived a new edit'

    # A no-op level command must not consume history or erase Redo.
    Undo-Wbs
    Assert-History 1 1
    $wbs.Range('C8').Select()
    $wbs.Range('D2').Select()
    $sameLevel = $wbs.Range('D2').Value2
    $wbs.Range('D2').Value2 = $sameLevel
    Assert-History 1 1

    # Invalid hierarchy requests also keep a valid history entry intact.
    $invalidBefore = Get-RangeState $wbs.Range('B6:Y10')
    $rejected = $false
    try { $excel.Run("'$($book.Name.Replace("'", "''"))'!Model_1.SetTaskLevels", $wbs.Range('C6'), 0, -1) | Out-Null } catch { $rejected = $true }
    Assert-True $rejected 'Invalid first-task promotion was accepted'
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $invalidBefore) 'Invalid level command changed the body'
    Assert-History 1 1

    # Failure before restore preserves history, settings and the retry direction.
    $protectedState = Get-RangeState $wbs.Range('B6:Y10')
    $wbs.Protect()
    $failed = $false
    try { $excel.Run("'$($book.Name.Replace("'", "''"))'!Model_1.RedoLastEdit", $true) | Out-Null } catch { $failed = $true }
    Assert-True $failed 'Protected restore should report a failure'
    Assert-True $excel.EnableEvents 'Failed restore left events disabled'
    $wbs.Unprotect()
    Assert-History 1 1
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $protectedState) 'Failed restore corrupted cells'
    Redo-Wbs
    Assert-True ($wbs.Range('D7').Value2 -eq 'New history branch') 'Redo retry did not use the retained snapshot'
    $protectedState = Get-RangeState $wbs.Range('B6:Y10')
    $wbs.Protect()
    $failed = $false
    try { $excel.Run("'$($book.Name.Replace("'", "''"))'!Model_1.UndoLastEdit", $true) | Out-Null } catch { $failed = $true }
    $wbs.Unprotect()
    Assert-True $failed 'Protected Undo should report a failure'
    Assert-True ((Get-RangeState $wbs.Range('B6:Y10')) -eq $protectedState) 'Failed Undo corrupted cells'
    Assert-History 2 0
    Undo-Wbs
    Redo-Wbs

    # Unobserved edits and standalone formatting invalidate stale history.
    Undo-Wbs
    $excel.EnableEvents = $false
    $wbs.Range('D8').Value2 = 'External edit'
    $excel.EnableEvents = $true
    Run-Macro 'ProgressChart.RedoWbs'
    Assert-History 0 0
    Assert-True ($wbs.Range('D8').Value2 -eq 'External edit') 'Stale Redo overwrote an external edit'
    $wbs.Range('D8').Select()
    $wbs.Range('D8').Value2 = 'Recorded edit'
    $wbs.Range('D8').Font.Italic = $true
    Run-Macro 'ProgressChart.UndoWbs'
    Assert-History 0 0
    Assert-True $wbs.Range('D8').Font.Italic 'Stale Undo overwrote external formatting'

    # The next edit must not absorb changes made outside its target without events.
    Run-Macro 'Model_1.ResetUndoTracking'
    $wbs.Range('D7').Value2 = 'Before outside change'
    $excel.EnableEvents = $false
    $wbs.Range('D8').Value2 = 'Keep outside change'
    $excel.EnableEvents = $true
    $wbs.Range('D7').Value2 = 'After outside change'
    Assert-History 0 0
    Run-Macro 'ProgressChart.UndoWbs'
    Assert-True ($wbs.Range('D8').Value2 -eq 'Keep outside change') 'Next edit absorbed an outside change'

    # Ribbon restrictions belong to this document, and switching back keeps history.
    $wbs.Range('D7').Select()
    $wbs.Range('D7').Value2 = 'Before workbook switch'
    Assert-History 1 0
    $other = $excel.Workbooks.Add()
    try {
        $excel.Visible = $true
        $other.Worksheets.Item(1).Range('A1').Select()
        # UI input, unlike COM Value2 assignment, creates a native Undo record.
        Send-ExcelKeys '123{ENTER}'
        Assert-True ($excel.CommandBars.GetEnabledMso('Undo')) 'Native Undo stayed disabled in another workbook'
        $excel.CommandBars.ExecuteMso('Undo')
        Assert-True ($excel.CommandBars.GetEnabledMso('Redo')) 'Native Redo stayed disabled in another workbook'
    } finally { $other.Close($false) }
    $book.Activate()
    $wbs.Activate()
    Assert-History 1 0
    Undo-Wbs
    Redo-Wbs

    # Actual shortcut dispatch uses the same history and releases on leaving WBS.
    $wbs.Range('D7').Select()
    Send-ExcelKeys '^z'
    Assert-History 0 1
    Send-ExcelKeys '^y'
    Assert-History 1 0

    # Whole-row insertion and deletion clear history without stale row replay.
    $wbs.Rows.Item(20).Insert() | Out-Null
    Assert-History 0 0
    $wbs.Range('D7').Select()
    $wbs.Range('D7').Value2 = 'Before row deletion'
    Assert-History 1 0
    $wbs.Rows.Item(20).Delete() | Out-Null
    Assert-History 0 0

    # Macro refresh is a documented history boundary.
    Run-Macro 'ProgressChart.UpdateGantt'
    Assert-History 0 0
    $pg.Range('B14:J213').ClearContents()
    Fill-Node 14 'Invalid refresh' ([datetime]'2026-09-20')
    $pg.Range('E14').Value2 = 'invalid'
    $pg.Range('B3').Font.Name = 'Arial'
    $staleCopy = $pg.Shapes.Item('WBS_ProgressChart').Duplicate()
    $staleCopy.Item(1).Name = 'WBS_ProgressChart stale copy'
    $stalePart = $pg.Shapes.AddShape(1, 900, 10, 10, 10)
    $stalePart.Name = 'WBS_PG_TMP_SHAPE_TEST'
    $note = $pg.Shapes.AddTextbox(1, 900, 30, 100, 20)
    $note.Name = 'User note'
    $note.TextFrame2.TextRange.Text = 'Keep this note'
    $refreshFailed = $false
    try { $excel.Run("'$($book.Name.Replace("'", "''"))'!ProgressChart.RefreshAll", $true) | Out-Null } catch { $refreshFailed = $true }
    Assert-True (-not $refreshFailed) 'Milestone failure interrupted the remaining refresh'
    Assert-NoMilestoneChart
    Assert-True ($pg.Range('B10').Value2.StartsWith('Milestones skipped after retry:')) 'Second failure did not record the reason'
    Assert-Font $pg.Range('B3:J10') 'Refresh All skipped settings formatting after milestone failure'
    Assert-Font $pg.Range('B14:J213') 'Refresh All skipped body formatting after milestone failure'
    Assert-History 0 0
    Assert-MilestoneHistory 0 0
    Assert-True ($pg.Shapes.Item('User note').TextFrame2.TextRange.Text -eq 'Keep this note') 'Chart cleanup deleted unrelated content'
    Assert-True $excel.ScreenUpdating 'Refresh failure left screen updating off'
    $pg.Shapes.Item('User note').Delete()
    $pg.Range('E14').Value2 = ([datetime]'2026-09-20').ToOADate()
    $excel.Run("'$($book.Name.Replace("'", "''"))'!ProgressChart.RefreshAll", $true) | Out-Null
    Assert-Chart 1

    # A tail edit captures the predecessor-validation extension as well.
    $wbs.Activate()
    $wbs.Range('C1001').Select()
    $tailBefore = Get-RangeState $wbs.Range('B999:Y1002')
    $wbs.Range('C1001').Value2 = 'New tail task'
    Assert-RoundTrip $wbs.Range('B999:Y1002') $tailBefore
    Redo-Wbs
    Assert-True ($wbs.Range('O1201').Validation.Formula1 -eq '=$C$6:$C$1001') 'Redo missed the extended predecessor dropdown'
    $wbs.Range('C1001').ClearContents()

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
    Test-Milestones
    Test-MilestoneInputFailures
    # Closing discards in-memory history; reopening the delivered workbook starts clean.
    $book.Close($false)
    $book = $null
    $book = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path, 0, $true)
    $wbs = $book.Worksheets.Item('WBS')
    $pg = $book.Worksheets.Item('Milestones')
    $wbs.Activate()
    Assert-History 0 0
    Assert-Font $wbs.Range('B6:BG19') 'Reopened initial body font changed'
    Assert-MilestoneHistory 0 0
    Test-FixedToolbar
    Write-Output 'Excel smoke checks passed. See USAGE.md for remaining interactive checks.'
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
