# Windows + Microsoft Excel integration checks for drawms.
# Run in an Excel trusted location:
#   pwsh -File tests/excel_drawms.ps1 -WorkbookPath ./dist/drawms.xlsm
param([Parameter(Mandatory = $true)][string]$WorkbookPath)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Run-Macro([string]$Name) {
    $prefix = "'$($book.Name.Replace("'", "''"))'!"
    $excel.Run($prefix + $Name) | Out-Null
}

function Group-Text([object]$Group) {
    $values = @()
    for ($i = 1; $i -le $Group.GroupItems.Count; $i++) {
        $shape = $Group.GroupItems.Item($i)
        if ($shape.Type -eq 17) {
            $values += ($shape.TextFrame2.TextRange.Text -replace '[\r\n]', '')
        }
    }
    return $values
}

try {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $book = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path, 0, $true)
    Assert-True ($book.Worksheets.Count -eq 1) 'drawms must contain one worksheet'
    $ws = $book.Worksheets.Item('Milestones')
    Assert-True ($null -ne $ws) 'Milestones worksheet is missing'
    Assert-True ($ws.Shapes.Count -eq 4) 'Expected three buttons and one chart'
    foreach ($name in @('GenerateChart', 'WidenSpacing', 'NarrowSpacing')) {
        $button = $ws.Shapes.Item("DM_Button_$name")
        Assert-True ($button.OnAction -match "DrawMs\.$name$") "Missing $name binding"
    }
    foreach ($cell in @($ws.Range('B14'), $ws.Range('C14'))) {
        Assert-True ($cell.NumberFormat -eq '@') "Input cell $($cell.Address()) is not Text formatted"
        Assert-True (-not $cell.HasFormula) "Input cell $($cell.Address()) unexpectedly has a formula"
    }

    $excel.EnableEvents = $false
    $ws.Range('B14:C213').ClearContents()
    $ws.Range('B14').Value2 = 'Second row entered first'
    $ws.Range('C14').Value2 = 'later'
    $ws.Range('B15').Value2 = 'Same date'
    $ws.Range('C15').Value2 = 'later'
    $ws.Range('B16').Value2 = 'Name only'
    $ws.Range('B17').Value2 = 'Earlier text'
    $ws.Range('C17').Value2 = 'before'
    $excel.EnableEvents = $true
    Run-Macro 'DrawMs.GenerateChart'

    $chart = $ws.Shapes.Item('DrawMs_Chart')
    Assert-True ($chart.Type -eq 6) 'drawms chart is not a grouped shape'
    Assert-True ($ws.Shapes.Count -eq 4) 'Generation left stale top-level drawings'
    $texts = @(Group-Text $chart)
    Assert-True ($texts -contains 'Second row entered first') 'Name text was changed'
    Assert-True ($texts -contains 'later') 'Date text was changed'
    Assert-True ($texts -contains 'before') 'Out-of-order date text was changed'
    Assert-True ($ws.Range('C14').Value2 -eq 'later' -and -not $ws.Range('C14').HasFormula) 'Date input was normalized'
    Assert-True ($chart.AlternativeText -match 'scale=1\.0') 'Initial spacing metadata is wrong'

    Run-Macro 'DrawMs.NarrowSpacing'
    $chart = $ws.Shapes.Item('DrawMs_Chart')
    Assert-True ($chart.AlternativeText -match 'scale=0\.9') 'Narrow spacing did not reduce scale'
    Run-Macro 'DrawMs.WidenSpacing'
    $chart = $ws.Shapes.Item('DrawMs_Chart')
    Assert-True ($chart.AlternativeText -match 'scale=1\.0') 'Widen spacing did not restore scale'

    # There is no document-level Ribbon override or shortcut replacement in drawms.
    Assert-True ([bool]$excel.CommandBars.GetEnabledMso('Undo')) 'Excel native Undo is unexpectedly disabled'
    Assert-True ([bool]$excel.CommandBars.GetEnabledMso('Redo')) 'Excel native Redo is unexpectedly disabled'
    Write-Output 'drawms desktop Excel scenarios passed.'
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $book) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
