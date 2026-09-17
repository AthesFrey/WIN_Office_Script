# Loaded by excel_smoke.ps1. Requires Windows desktop Excel; syntax-only on Linux.
function Assert-MilestonesToolbar([switch]$CheckReservedRows) {
    $names = @('GenerateProgressChart', 'CopyProgressChart', 'WidenProgressChart', 'NarrowProgressChart',
        'UndoMilestones', 'RedoMilestones', 'DeleteMilestoneLevel1', 'DeleteMilestoneLevel2',
        'DeleteMilestoneLevel3', 'DeleteMilestoneLevel4', 'DeleteMilestoneLevel5')
    $lefts = @(25, 208, 335, 472, 25, 121, 217, 325, 433, 541, 649)
    $widths = @(171, 114, 125, 125, 88, 88, 100, 100, 100, 100, 100)
    $rectangles = @()
    for ($i = 0; $i -lt $names.Count; $i++) {
        $shape = $pg.Shapes.Item('PG_Button_' + $names[$i])
        $top = if ($i -lt 4) { 93 } else { 235 }
        Assert-True ([math]::Abs($shape.Left - $lefts[$i]) -lt 0.2) "Wrong Left for $($names[$i])"
        Assert-True ([math]::Abs($shape.Top - $top) -lt 0.2) "Wrong Top for $($names[$i])"
        Assert-True ([math]::Abs($shape.Width - $widths[$i]) -lt 0.2) "Wrong Width for $($names[$i])"
        Assert-True ([math]::Abs($shape.Height - 25) -lt 0.2) "Wrong Height for $($names[$i])"
        Assert-True ($shape.Placement -eq 3) "Button must not move or size with cells: $($names[$i])"
        if ($CheckReservedRows) {
            $row = if ($i -lt 4) { 5 } else { 11 }
            $limit = if ($i -lt 4) { 6 } else { 13 }
            Assert-True ([math]::Abs($shape.Top - $pg.Rows.Item($row).Top - 3) -lt 0.2) 'Initial top inset changed'
            Assert-True ($shape.Top + $shape.Height -le $pg.Rows.Item($limit).Top) 'Button overlaps settings/header'
        }
        Assert-True ([string]$shape.OnAction -match ('ProgressChart\.' + $names[$i] + '$')) 'Button binding changed'
        foreach ($prior in $rectangles) {
            $safe = ($shape.Left + $shape.Width -le $prior.Left) -or
                ($prior.Left + $prior.Width -le $shape.Left) -or
                ($shape.Top + $shape.Height -le $prior.Top) -or
                ($prior.Top + $prior.Height -le $shape.Top)
            Assert-True $safe 'Milestones buttons overlap in Excel'
        }
        $rectangles += $shape
    }
}

function Test-FixedToolbar {
    Assert-MilestonesToolbar -CheckReservedRows
    $pg.Activate()
    $oldZoom = $excel.ActiveWindow.Zoom
    $oldHeight = $pg.Rows.Item(2).RowHeight
    $oldWidth = $pg.Columns.Item(2).ColumnWidth
    $oldEvents = $excel.EnableEvents
    $copyBook = $null
    $copyPath = Join-Path ([IO.Path]::GetTempPath()) ('WBS-fixed-toolbar-' + [guid]::NewGuid().ToString('N') + '.xlsm')
    $before = Get-RangeState $pg.Range('B14:J19')
    try {
        foreach ($shape in $wbs.Shapes) {
            if ($shape.Name -like 'PG_Button_*') {
                Assert-True ($shape.Placement -eq 3) 'WBS toolbar must use the same fixed placement'
            }
        }
        foreach ($zoom in @(80, 100, 125)) {
            $excel.ActiveWindow.Zoom = $zoom
            Assert-MilestonesToolbar
        }
        # Check immediately, before selection, activation or any button macro.
        # With events off, VBA cannot mask a cell-dependent placement defect.
        foreach ($events in @($false, $true)) {
            $excel.EnableEvents = $events
            $pg.Columns.Item(2).ColumnWidth = $oldWidth + 5
            Assert-MilestonesToolbar
            $pg.Columns.Item(2).ColumnWidth = $oldWidth - 5
            Assert-MilestonesToolbar
            $pg.Rows.Item(2).RowHeight = $oldHeight + 17
            Assert-MilestonesToolbar
            $pg.Rows.Item(2).RowHeight = $oldHeight - 10
            Assert-MilestonesToolbar
            # Move a sample range down/right and back, as with a cell drag.
            $pg.Range('D14:E14').Cut($pg.Range('G20:H20'))
            Assert-MilestonesToolbar
            $pg.Range('G20:H20').Cut($pg.Range('D14:E14'))
            Assert-MilestonesToolbar
            $pg.Rows.Item(2).RowHeight = $oldHeight
            $pg.Columns.Item(2).ColumnWidth = $oldWidth
            Assert-MilestonesToolbar -CheckReservedRows
            Assert-True ($excel.EnableEvents -eq $events) 'Cell operations changed the event setting'
        }
        Assert-True ((Get-RangeState $pg.Range('B14:J19')) -eq $before) 'Cell move round-trip changed milestone data'
        $wbs.Activate()
        $pg.Activate()
        Assert-MilestonesToolbar -CheckReservedRows
        # Both a chart button and Refresh All must leave the fixed toolbar alone.
        $prefix = "'$($book.Name.Replace("'", "''"))'!"
        $excel.Run($pg.Shapes.Item('PG_Button_GenerateProgressChart').OnAction, $true) | Out-Null
        Assert-MilestonesToolbar
        $excel.Run($prefix + 'ProgressChart.RefreshAll', $true) | Out-Null
        Assert-MilestonesToolbar
        Assert-True $excel.EnableEvents 'Refresh All left events disabled'
        Assert-True (-not [bool]$excel.Run($prefix + 'Model_1.HistoryMaintenanceActive')) 'History guard was not released'
        Assert-True ([string]$excel.Run($prefix + 'Model_1.LastHistoryError') -eq '') 'History initialization recorded a failure'
        # Persist changed cells; reopening must keep absolute placement too.
        $pg.Rows.Item(2).RowHeight = $oldHeight + 17
        $pg.Columns.Item(2).ColumnWidth = $oldWidth + 5
        Assert-MilestonesToolbar
        $book.SaveCopyAs($copyPath)
        $copyBook = $excel.Workbooks.Open($copyPath, 0, $true)
        $originalPg = $pg
        try {
            $pg = $copyBook.Worksheets.Item('Milestones')
            Assert-MilestonesToolbar
            $pg.Activate()
            Assert-MilestonesToolbar
            $pg.Columns.Item(2).ColumnWidth = $oldWidth - 5
            Assert-MilestonesToolbar
        } finally {
            $pg = $originalPg
            $copyBook.Close($false)
            $copyBook = $null
        }
    } finally {
        if ($null -ne $copyBook) { $copyBook.Close($false) }
        if (Test-Path $copyPath) { Remove-Item -LiteralPath $copyPath }
        $excel.EnableEvents = $false
        $pg.Rows.Item(2).RowHeight = $oldHeight
        $pg.Columns.Item(2).ColumnWidth = $oldWidth
        $pg.Activate()
        $excel.ActiveWindow.Zoom = $oldZoom
        $excel.EnableEvents = $oldEvents
        $wbs.Activate()
    }
}

function Test-MilestoneInputFailures {
    $prefix = "'$($book.Name.Replace("'", "''"))'!"
    $pg.Activate()
    $excel.EnableEvents = $false
    $pg.Range('B14:J213').ClearContents()
    $pg.Range('C6').Value2 = 'Equal spacing'
    $pg.Range('C7').Value2 = 1.0
    Fill-Node 14 'Normalizes before failing' ([datetime]'2026-09-01')
    Fill-Node 15 'Invalid input follows' ([datetime]'2026-09-02')
    $pg.Range('E14:E15').NumberFormat = '@'
    $pg.Range('E14:E15').Font.Name = 'Microsoft YaHei'
    $pg.Range('E14:E15').Font.Size = 11
    $pg.Range('E14').Value2 = '20260901'
    $pg.Range('E15').Value2 = 'invalid'
    Run-Macro 'Sheet4.ResetUndoTracking'
    $excel.EnableEvents = $true
    $beforeDates = Get-RangeState $pg.Range('E14:E15')
    $oldCalculation = $excel.Calculation
    $oldScreen = $excel.ScreenUpdating
    $excel.Run($prefix + 'ProgressChart.RefreshAll', $true) | Out-Null
    Assert-True ((Get-RangeState $pg.Range('E14:E15')) -eq $beforeDates) 'Refresh All left partial date normalization'
    Assert-NoMilestoneChart
    Assert-MilestoneHistory 0 0
    Assert-True ($excel.EnableEvents -and $excel.Calculation -eq $oldCalculation -and $excel.ScreenUpdating -eq $oldScreen) 'Failed chart refresh changed Excel settings'
    $pg.Activate()
    # CLng previously accepted signed and scientific notation components.
    foreach ($invalidDate in @('2E03/1/2', '+999/1/2', '2026/+1/2', '2026/1/+2', '2026/1/2e0', '2026/2/30')) {
        $pg.Range('E15').Value2 = $invalidDate
        $beforeDates = Get-RangeState $pg.Range('E14:E15')
        $failed = $false
        try { $excel.Run($prefix + 'ProgressChart.GenerateProgressChart', $true) | Out-Null } catch { $failed = $true }
        Assert-True $failed "Accepted invalid date: $invalidDate"
        Assert-True ((Get-RangeState $pg.Range('E14:E15')) -eq $beforeDates) 'Rejected date changed input cells'
    }
    # Nonblank invalid sources must not silently become manual milestones.
    $pg.Range('E15').Value2 = '20260902'
    foreach ($source in @(6.5, 2e12, $true, 'invalid', 900001.0, '=#REF!')) {
        if ($source -is [string] -and $source.StartsWith('=')) { $pg.Range('C15').Formula = $source }
        else { $pg.Range('C15').Value2 = $source }
        $failed = $false
        try { $excel.Run($prefix + 'ProgressChart.GenerateProgressChart', $true) | Out-Null } catch { $failed = $true }
        Assert-True $failed "Accepted invalid WBS source: $source"
    }
    $pg.Range('C15').ClearContents()
    $excel.Run($prefix + 'ProgressChart.GenerateProgressChart', $true) | Out-Null
    Assert-Chart 2
    # Empty WBS must not turn rows 1:6 / its headers into import candidates.
    $excel.EnableEvents = $false
    $wbs.Range('B6:Y3010').ClearContents()
    $excel.EnableEvents = $true
    $wbs.Activate()
    $wbs.Range('C5').Select()
    $before = Get-RangeState $pg.Range('B14:J19')
    $failed = $false
    try { $excel.Run($prefix + 'ProgressChart.ImportSelectedMilestones', $true) | Out-Null } catch { $failed = $true }
    Assert-True $failed 'Empty WBS import did not report an error'
    Assert-True ((Get-RangeState $pg.Range('B14:J19')) -eq $before) 'Empty WBS import changed milestones'
    # With manual calculation, cached source formulas must be recalculated
    # before duplicate detection or the same task can be imported twice.
    $excel.EnableEvents = $false
    $wbs.Range('C6').Value2 = 'Linked task'
    $wbs.Range('G6').Value2 = ([datetime]'2026-09-01').ToOADate()
    $wbs.Range('F6').Value2 = 2.0
    $excel.Calculation = -4135
    $pg.Range('B14:J213').ClearContents()
    $pg.Range('F14').Value2 = 7.0
    $pg.Range('C14').Formula = '=F14'
    $pg.Range('D14').Value2 = 'Existing link'
    $pg.Calculate()
    $pg.Range('F14').Value2 = 6.0 # Cached C14 still says 7 until Calculate.
    $wbs.Range('C7').Value2 = 'Another task'
    $wbs.Range('G7').Value2 = ([datetime]'2026-09-03').ToOADate()
    $wbs.Range('F7').Value2 = 1.0
    $excel.EnableEvents = $true
    try {
        $wbs.Range('C6:C7').Select()
        $excel.Run($prefix + 'ProgressChart.ImportSelectedMilestones', $true) | Out-Null
        Assert-True ($pg.Range('C14').Value2 -eq 6 -and $pg.Range('C15').Value2 -eq 7) 'Manual calculation import duplicated the wrong source'
        Assert-True ($excel.WorksheetFunction.CountA($pg.Range('D14:D213')) -eq 2) 'Duplicate task imported in manual calculation mode'
        Assert-True ($excel.Calculation -eq -4135) 'Import changed manual calculation mode'
    } finally { $excel.Calculation = $oldCalculation }
    Write-Output 'Milestone input failure scenarios passed in desktop Excel.'
}
