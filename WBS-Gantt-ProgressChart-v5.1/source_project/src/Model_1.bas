Attribute VB_Name = "Model_1"
Option Explicit

Public Const g_max_level As Long = 5
Private g_lastRow As Long
Private g_firstDate As Date
Private g_lastDate As Date
Private g_lastDateCol As Long
Private g_numbers(1 To g_max_level) As Long
Private g_values As Variant
Private g_nameRows As Object
Private g_aliases As Object
Private g_levels() As Long
Private g_parents() As Long
Private g_firstChild() As Long
Private g_nextSibling() As Long
Private g_subtreeEnd() As Long
Private g_predecessors() As Long
Private g_leafDays() As Double
Private g_weightedProgress() As Double

Private Function Wbs() As Worksheet
    Set Wbs = ThisWorkbook.Worksheets("WBS")
End Function

Public Function CalcWorkDate(ByVal startDate As Date, ByVal days As Long) As Date
    CalcWorkDate = DateAdd("d", Application.Max(1, days) - 1, startDate)
End Function

Public Sub UpdateGanttChart(Optional ByVal propagateErrors As Boolean = False)
    Dim oldEvents As Boolean, oldScreen As Boolean, r As Long, errText As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    ThisWorkbook.Activate: Wbs.Select Replace:=True
    RefreshParentTasks
    FindTimelineRange
    ClearGanttChart
    BuildTimeline
    For r = 6 To g_lastRow: DrawGanttChart r: Next r
    DrawCurrentDayLine
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If Len(errText) > 0 Then
        If propagateErrors Then
            On Error GoTo 0
            Err.Raise vbObjectError + 506, "Update Gantt", errText
        End If
        MsgBox errText, vbExclamation, "Update Gantt"
    End If
    Exit Sub
Failed:
    errText = Err.Description
    Resume Done
End Sub

Public Sub RefreshParentTasks(Optional ByVal rebuild As Boolean = True, Optional ByVal linkPredecessors As Boolean = True)
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation
    Dim errorNumber As Long, errorText As String, before As Variant, r As Long, c As Variant
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    g_lastRow = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
    If g_lastRow < 6 Then GoTo Done
    Wbs.Range("C6:O" & g_lastRow).Calculate
    LoadTaskValues
    Set g_aliases = Nothing
    If rebuild Then RebuildHierarchy
    BuildTaskIndex
    If rebuild Or linkPredecessors Then BindPredecessors
    before = g_values
    CalculateSchedule
    ' Calculate the entire dependency graph before committing any schedule values.
    For r = 6 To g_lastRow
        If g_levels(r) > 0 Then
            For Each c In Array(6, 7, 8, 14)
                If ValuesDiffer(before(r, c), g_values(r, c)) Then Wbs.Cells(r, c).Value = g_values(r, c)
            Next c
        End If
    Next r
    If rebuild Then
        Wbs.Range("G6:H" & g_lastRow).NumberFormat = "yyyy/m/d"
        Wbs.Range("N6:N" & g_lastRow).NumberFormat = "0%"
        With Wbs.Range("O6:O" & Application.Max(1000, g_lastRow + 200)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$C$6:$C$" & g_lastRow
            .IgnoreBlank = True: .InCellDropdown = True: .ShowError = False
        End With
    End If
Done:
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "WBS", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Done
End Sub

Private Sub LoadTaskValues()
    Dim dates As Variant, r As Long
    g_values = Wbs.Range("A1:Y" & g_lastRow).Value2
    ' Preserve Excel's Date subtype, including serials that resemble mmdd input.
    dates = Wbs.Range("G6:H" & g_lastRow).Value
    For r = 6 To g_lastRow
        g_values(r, 7) = dates(r - 5, 1): g_values(r, 8) = dates(r - 5, 2)
    Next r
End Sub

Private Function ValuesDiffer(ByVal before As Variant, ByVal after As Variant) As Boolean
    ValuesDiffer = (CStr(before) <> CStr(after))
End Function

Private Sub PutTaskValue(ByVal r As Long, ByVal c As Long, ByVal value As Variant)
    If ValuesDiffer(g_values(r, c), value) Then Wbs.Cells(r, c).Value = value
    g_values(r, c) = value
End Sub

Public Function DetectLevel(ByVal value As String) As Long
    Dim n As Long
    n = Len(value) - Len(LTrim$(value))
    DetectLevel = Application.Min(g_max_level, Int(n / 4) + 1)
End Function

Private Function RebuildHierarchyNumber(ByVal level As Long) As String
    Dim i As Long
    g_numbers(level) = g_numbers(level) + 1
    For i = level + 1 To g_max_level: g_numbers(i) = 0: Next i
    RebuildHierarchyNumber = CStr(g_numbers(1))
    For i = 2 To level: RebuildHierarchyNumber = RebuildHierarchyNumber & "." & g_numbers(i): Next i
End Function

Private Sub RebuildHierarchy()
    Dim r As Long, i As Long, level As Long, priorLevel As Long
    Dim value As String, oldName As String, parent As Boolean, nextLevel As Long
    Set g_aliases = CreateObject("Scripting.Dictionary")
    g_aliases.CompareMode = vbTextCompare
    For i = 1 To g_max_level: g_numbers(i) = 0: Next i
    For r = 6 To g_lastRow
        oldName = CStr(g_values(r, 3))
        If Len(Trim$(oldName)) > 0 Then
            level = DetectLevel(oldName)
            If priorLevel = 0 Then level = 1
            If level > priorLevel + 1 Then level = priorLevel + 1
            value = String$((level - 1) * 4, " ") & RebuildHierarchyNumber(level) & " " & ProgressChart.ProgressLabel(oldName)
            PutTaskValue r, 3, value
            PutTaskValue r, 25, level
            PutTaskValue r, 2, r - 5
            Wbs.Cells(r, 3).WrapText = True
            Wbs.Rows(r).RowHeight = Application.Max(24, (Len(value) \ 36 + 1) * 15)
            g_aliases(NormalizeTaskText(oldName)) = value
            priorLevel = level
        Else
            PutTaskValue r, 2, Empty: PutTaskValue r, 24, Empty: PutTaskValue r, 25, Empty
        End If
    Next r
    For r = g_lastRow To 6 Step -1
        If Len(Trim$(CStr(g_values(r, 3)))) > 0 Then
            level = DetectLevel(CStr(g_values(r, 3)))
            parent = nextLevel > level
            PutTaskValue r, 24, IIf(parent, "P", "T")
            Wbs.Range(Wbs.Cells(r, 2), Wbs.Cells(r, 25)).Font.Bold = parent
            nextLevel = level
        End If
    Next r
End Sub

Private Sub BuildTaskIndex()
    Dim r As Long, level As Long, depth As Long, stack(1 To g_max_level) As Long, key As String
    Dim lastChild() As Long, parent As Long
    ReDim g_levels(6 To g_lastRow): ReDim g_parents(6 To g_lastRow)
    ReDim g_firstChild(6 To g_lastRow): ReDim g_nextSibling(6 To g_lastRow)
    ReDim g_subtreeEnd(6 To g_lastRow): ReDim lastChild(6 To g_lastRow)
    Set g_nameRows = CreateObject("Scripting.Dictionary")
    g_nameRows.CompareMode = vbTextCompare
    For r = 6 To g_lastRow
        key = NormalizeTaskText(CStr(g_values(r, 3)))
        If Len(key) > 0 Then
            level = DetectLevel(CStr(g_values(r, 3))): g_levels(r) = level
            Do While depth > 0
                If g_levels(stack(depth)) < level Then Exit Do
                g_subtreeEnd(stack(depth)) = r - 1: depth = depth - 1
            Loop
            If depth > 0 Then
                parent = stack(depth): g_parents(r) = parent
                If lastChild(parent) = 0 Then
                    g_firstChild(parent) = r
                Else
                    g_nextSibling(lastChild(parent)) = r
                End If
                lastChild(parent) = r
            End If
            depth = depth + 1: stack(depth) = r
            If Not g_nameRows.Exists(key) Then g_nameRows.Add key, r
        End If
    Next r
    Do While depth > 0
        g_subtreeEnd(stack(depth)) = g_lastRow: depth = depth - 1
    Loop
End Sub

Private Sub BindPredecessors()
    Dim values As Variant, formulas As Variant, r As Long, candidate As Long, key As String, formula As String
    Wbs.Range("O6:O" & g_lastRow).Calculate
    ' Include the header so these reads always return two-dimensional arrays.
    values = Wbs.Range("O5:O" & g_lastRow).Value2
    formulas = Wbs.Range("O5:O" & g_lastRow).Formula
    For r = 6 To g_lastRow
        g_values(r, 15) = values(r - 4, 1)
        If g_levels(r) > 0 And Not IsError(g_values(r, 15)) Then
            key = NormalizeTaskText(CStr(g_values(r, 15)))
            If Left$(CStr(formulas(r - 4, 1)), 1) <> "=" Then
                If Not g_aliases Is Nothing Then
                    If g_aliases.Exists(key) Then key = NormalizeTaskText(CStr(g_aliases(key)))
                End If
            End If
            candidate = FindPredecessor(key)
            If candidate > 0 Then
                formula = "=C" & candidate
                If CStr(formulas(r - 4, 1)) <> formula Then Wbs.Cells(r, 15).Formula = formula
                g_values(r, 15) = g_values(candidate, 3)
            End If
        End If
    Next r
End Sub

Public Function IsParentTask(ByVal rowNumber As Long) As Boolean
    Dim nextRow As Long, lastRow As Long
    lastRow = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
    If Len(Trim$(CStr(Wbs.Cells(rowNumber, 3).Value2))) = 0 Then Exit Function
    For nextRow = rowNumber + 1 To lastRow
        If Len(Trim$(CStr(Wbs.Cells(nextRow, 3).Value2))) > 0 Then
            IsParentTask = DetectLevel(CStr(Wbs.Cells(nextRow, 3).Value2)) > DetectLevel(CStr(Wbs.Cells(rowNumber, 3).Value2))
            Exit Function
        End If
    Next nextRow
End Function

Private Function SubtreeEnd(ByVal rowNumber As Long) As Long
    Dim r As Long, level As Long, lastRow As Long
    lastRow = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
    level = DetectLevel(CStr(Wbs.Cells(rowNumber, 3).Value2))
    SubtreeEnd = rowNumber
    For r = rowNumber + 1 To lastRow
        If Len(Trim$(CStr(Wbs.Cells(r, 3).Value2))) > 0 Then
            If DetectLevel(CStr(Wbs.Cells(r, 3).Value2)) <= level Then Exit Function
        End If
        SubtreeEnd = r
    Next r
End Function

Private Sub AddDependency(ByVal task As Long, ByVal dependency As Long, ByRef pending() As Long, ByRef heads() As Long, ByRef targets() As Long, ByRef links() As Long, ByRef count As Long)
    count = count + 1: targets(count) = task: links(count) = heads(dependency)
    heads(dependency) = count: pending(task) = pending(task) + 1
End Sub

Private Sub CalculateSchedule()
    Dim pending() As Long, heads() As Long, targets() As Long, links() As Long, queue() As Long
    Dim r As Long, count As Long, edge As Long, first As Long, last As Long, finished As Long, total As Long, dependent As Long
    Dim value As Variant, predecessor As Long
    ReDim pending(6 To g_lastRow): ReDim heads(6 To g_lastRow)
    ReDim targets(1 To 2 * g_lastRow): ReDim links(1 To 2 * g_lastRow): ReDim queue(1 To g_lastRow)
    ReDim g_predecessors(6 To g_lastRow): ReDim g_leafDays(6 To g_lastRow): ReDim g_weightedProgress(6 To g_lastRow)
    For r = 6 To g_lastRow
        If g_levels(r) > 0 Then
            total = total + 1
            If g_parents(r) > 0 Then AddDependency g_parents(r), r, pending, heads, targets, links, count
            If g_firstChild(r) = 0 Then
                value = g_values(r, 15)
                If IsError(value) Then Err.Raise vbObjectError + 502, , "Invalid predecessor at row " & r & "."
                If Len(Trim$(CStr(value))) > 0 Then
                    predecessor = FindPredecessor(CStr(value))
                    If predecessor = 0 Then Err.Raise vbObjectError + 502, , "Predecessor not found at row " & r & ". Select a task from the list."
                    g_predecessors(r) = predecessor
                    AddDependency r, predecessor, pending, heads, targets, links, count
                End If
            End If
        End If
    Next r
    For r = 6 To g_lastRow
        If g_levels(r) > 0 And pending(r) = 0 Then last = last + 1: queue(last) = r
    Next r
    first = 1
    Do While first <= last
        r = queue(first): first = first + 1
        ResolveTaskValues r: finished = finished + 1
        edge = heads(r)
        Do While edge > 0
            dependent = targets(edge): pending(dependent) = pending(dependent) - 1
            If pending(dependent) = 0 Then last = last + 1: queue(last) = dependent
            edge = links(edge)
        Loop
    Loop
    If finished <> total Then
        For r = 6 To g_lastRow
            If pending(r) > 0 Then Err.Raise vbObjectError + 501, , "Circular predecessor or parent dependency at row " & r & "."
        Next r
    End If
End Sub

Private Sub ResolveTaskValues(ByVal r As Long)
    Dim child As Long, predecessor As Long, value As Variant, startDate As Date, endDate As Date
    Dim minDate As Date, maxDate As Date, days As Double, progress As Double
    child = g_firstChild(r)
    If child > 0 Then
        minDate = DateSerial(9999, 12, 31): maxDate = DateSerial(1900, 1, 1)
        Do While child > 0
            If g_leafDays(child) > 0 Then
                startDate = CDate(g_values(child, 7)): endDate = CDate(g_values(child, 8))
                If startDate < minDate Then minDate = startDate
                If endDate > maxDate Then maxDate = endDate
                g_leafDays(r) = g_leafDays(r) + g_leafDays(child)
                g_weightedProgress(r) = g_weightedProgress(r) + g_weightedProgress(child)
            End If
            child = g_nextSibling(child)
        Loop
        If g_leafDays(r) > 0 Then
            g_values(r, 7) = minDate: g_values(r, 8) = maxDate
            g_values(r, 6) = DateDiff("d", minDate, maxDate) + 1
            g_values(r, 14) = g_weightedProgress(r) / g_leafDays(r)
        Else
            g_values(r, 6) = Empty: g_values(r, 7) = Empty: g_values(r, 8) = Empty: g_values(r, 14) = Empty
        End If
    Else
        predecessor = g_predecessors(r)
        If predecessor > 0 Then
            If ProgressChart.TryProgressDate(g_values(predecessor, 8), endDate) Then
                g_values(r, 7) = DateAdd("d", 1, endDate)
            Else
                g_values(r, 7) = Empty: g_values(r, 8) = Empty
            End If
        End If
        value = g_values(r, 6)
        If Len(CStr(value)) = 0 Or Len(CStr(g_values(r, 7))) = 0 Then
            g_values(r, 8) = Empty
        ElseIf ProgressChart.TryProgressDate(g_values(r, 7), startDate) Then
            If Not IsNumeric(value) Then Err.Raise vbObjectError + 505, , "Days must be a positive whole number at row " & r & "."
            days = CDbl(value)
            If days < 1 Or days <> Fix(days) Then Err.Raise vbObjectError + 505, , "Days must be a positive whole number at row " & r & "."
            endDate = CalcWorkDate(startDate, CLng(days))
            g_values(r, 7) = startDate: g_values(r, 8) = endDate
            g_leafDays(r) = days
            value = g_values(r, 14)
            If Not IsError(value) Then
                If IsNumeric(value) Then progress = Application.Max(0, Application.Min(1, CDbl(value)))
            End If
            g_weightedProgress(r) = days * progress
        End If
    End If
End Sub

Private Function FindPredecessor(ByVal name As String) As Long
    Dim key As String
    key = NormalizeTaskText(name)
    If g_nameRows.Exists(key) Then FindPredecessor = CLng(g_nameRows(key))
End Function

Public Function NormalizeTaskText(ByVal value As String) As String
    value = Replace(Replace(Replace(value, " ", ""), vbTab, ""), ChrW(&H3000), "")
    NormalizeTaskText = Replace(Replace(Replace(value, ChrW(&HA0), ""), vbCr, ""), vbLf, "")
End Function

Public Function TryNormalizeWbsDate(ByVal target As Range, ByRef result As Date) As Boolean
    Dim previous As Date, fallbackYear As Long, shortInput As Boolean
    If target.Row > 6 Then
        If ProgressChart.TryProgressDate(target.Worksheet.Cells(target.Row - 1, target.Column).Value, previous) Then fallbackYear = Year(previous)
    End If
    TryNormalizeWbsDate = ProgressChart.TryProgressDateWithYear(target.Value2, fallbackYear, result, shortInput)
End Function

Public Function ReadWbsDate(ByVal rowNumber As Long, ByVal colNumber As Long, ByRef result As Date) As Boolean
    ReadWbsDate = ProgressChart.TryProgressDate(Wbs.Cells(rowNumber, colNumber).Value, result)
End Function

Public Sub AddTask()
    InsertTask False
End Sub

Public Sub AddSubtask()
    InsertTask True
End Sub

Private Sub InsertTask(ByVal asChild As Boolean)
    Dim r As Long, insertRow As Long, level As Long, oldEvents As Boolean, startDate As Date, hasStart As Boolean, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    If Not ActiveSheet Is Wbs Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    r = Selection.Row
    If r < 6 Or Len(Trim$(CStr(Wbs.Cells(r, 3).Value2))) = 0 Then
        r = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
        insertRow = Application.Max(6, r + 1)
        level = 1
    Else
        level = DetectLevel(CStr(Wbs.Cells(r, 3).Value2))
        If asChild Then level = level + 1
        insertRow = SubtreeEnd(r) + 1
        hasStart = ReadWbsDate(r, 7, startDate)
    End If
    If level > g_max_level Then MsgBox "The maximum task level is " & g_max_level & ".", vbInformation: Exit Sub
    Application.EnableEvents = False
    Wbs.Rows(insertRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    With Wbs.Range(Wbs.Cells(insertRow, 2), Wbs.Cells(insertRow, 25))
        .ClearContents: .Font.Bold = False
    End With
    Wbs.Cells(insertRow, 3).Value = String$((level - 1) * 4, " ") & "New task"
    Wbs.Cells(insertRow, 6).Value = 1: Wbs.Cells(insertRow, 14).Value = 0
    If hasStart Then Wbs.Cells(insertRow, 7).Value = startDate
    Wbs.Range("G" & insertRow & ":H" & insertRow).NumberFormat = "yyyy/m/d"
    Wbs.Cells(insertRow, 14).NumberFormat = "0%"
    RefreshParentTasks
    Application.EnableEvents = oldEvents
    Wbs.Cells(insertRow, 3).Select
Done:
    Application.EnableEvents = oldEvents
    If Len(errorText) > 0 Then MsgBox errorText, vbExclamation, "Add task"
    Exit Sub
Failed:
    errorText = Err.Description: Resume Done
End Sub

Public Function SelectedTaskCells(ByVal selected As Range) As Range
    Dim lastRow As Long
    If Not selected.Worksheet Is Wbs Then Exit Function
    lastRow = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
    If lastRow < 6 Then Exit Function
    Set SelectedTaskCells = Intersect(selected.EntireRow, Wbs.Range("C6:C" & lastRow))
End Function

Public Sub SetTaskLevels(ByVal selected As Range, ByVal targetLevel As Long, Optional ByVal delta As Long = 0)
    Dim picked As Range, cell As Range, selectedRows() As Boolean, planned() As Long
    Dim r As Long, child As Long, covered As Long, shift As Long, previous As Long, changed As Boolean
    Dim oldNames As Variant, oldLinks As Variant, errorNumber As Long, errorText As String, committed As Boolean
    If selected Is Nothing Then Err.Raise vbObjectError + 503, , "Select one or more tasks first."
    Set picked = SelectedTaskCells(selected)
    If picked Is Nothing Then Err.Raise vbObjectError + 503, , "Select one or more tasks first."
    g_lastRow = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
    LoadTaskValues: BuildTaskIndex
    ReDim selectedRows(6 To g_lastRow): ReDim planned(6 To g_lastRow)
    For Each cell In picked.Cells: selectedRows(cell.Row) = True: Next cell
    For r = 6 To g_lastRow: planned(r) = g_levels(r): Next r
    ' Use original subtree boundaries; overlapping selections move only once.
    For r = 6 To g_lastRow
        If selectedRows(r) And g_levels(r) > 0 And r > covered Then
            shift = delta
            If delta = 0 Then shift = targetLevel - g_levels(r)
            covered = g_subtreeEnd(r)
            For child = r To covered
                If g_levels(child) > 0 Then
                    planned(child) = g_levels(child) + shift
                    If planned(child) < 1 Or planned(child) > g_max_level Then Err.Raise vbObjectError + 503, , "The selection must stay within levels 1 to " & g_max_level & "."
                    If shift <> 0 Then changed = True
                End If
            Next child
        End If
    Next r
    For r = 6 To g_lastRow
        If planned(r) > 0 Then
            If planned(r) > previous + 1 Then Err.Raise vbObjectError + 503, , "Row " & r & " would skip a parent level. Select its preceding parent too, or choose a shallower level."
            previous = planned(r)
        End If
    Next r
    If Not changed Then Exit Sub
    oldNames = Wbs.Range("C6:C" & g_lastRow).Formula
    oldLinks = Wbs.Range("O6:O" & g_lastRow).Formula
    On Error GoTo Failed
    committed = True
    For r = 6 To g_lastRow
        If planned(r) <> g_levels(r) Then Wbs.Cells(r, 3).Value = String$((planned(r) - 1) * 4, " ") & Trim$(CStr(g_values(r, 3)))
    Next r
    RefreshParentTasks
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If committed Then
        Wbs.Range("C6:C" & g_lastRow).Formula = oldNames
        Wbs.Range("O6:O" & g_lastRow).Formula = oldLinks
        RefreshParentTasks
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "WBS", errorText
End Sub

Private Sub FindTimelineRange()
    Dim r As Long, s As Date, e As Date, found As Boolean
    g_firstDate = DateSerial(9999, 12, 31): g_lastDate = DateSerial(1900, 1, 1)
    For r = 6 To g_lastRow
        If Len(Trim$(CStr(Wbs.Cells(r, 3).Value2))) > 0 Then
            If ReadWbsDate(r, 7, s) And ReadWbsDate(r, 8, e) Then
                If e >= s Then
                    If s < g_firstDate Then g_firstDate = s
                    If e > g_lastDate Then g_lastDate = e
                    found = True
                End If
            End If
        End If
    Next r
    g_lastDateCol = 0
    If found Then
        g_lastDateCol = 26 + DateDiff("d", g_firstDate, g_lastDate)
        If g_lastDateCol >= Wbs.Columns.Count Then Err.Raise vbObjectError + 504, , "The plan exceeds Excel's available daily columns. Narrow the plan date range."
    End If
End Sub

Public Sub ClearGanttChart()
    Dim lastCol As Long, lastRow As Long, used As Range
    Set used = Wbs.UsedRange
    lastCol = Application.Max(26, used.Column + used.Columns.Count - 1)
    lastRow = Application.Max(6, used.Row + used.Rows.Count - 1)
    With Wbs.Range(Wbs.Cells(1, 26), Wbs.Cells(lastRow, lastCol))
        .UnMerge: .Clear
    End With
End Sub

Private Sub BuildTimeline()
    Dim c As Long, day As Date, weekdays As Variant, lastRow As Long
    If g_lastDateCol = 0 Then Exit Sub
    weekdays = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    lastRow = Application.Max(6, g_lastRow)
    For c = 26 To g_lastDateCol
        day = DateAdd("d", c - 26, g_firstDate)
        Wbs.Columns(c).ColumnWidth = 6
        Wbs.Cells(4, c).Value = day: Wbs.Cells(4, c).NumberFormat = "m/d"
        Wbs.Cells(5, c).Value = weekdays(Weekday(day, vbMonday) - 1)
        With Wbs.Range(Wbs.Cells(4, c), Wbs.Cells(lastRow, c))
            .Font.Name = "Arial": .Font.Size = 10
            If Weekday(day, vbMonday) > 5 Then .Interior.Color = RGB(241, 245, 249)
        End With
        Wbs.Range(Wbs.Cells(4, c), Wbs.Cells(5, c)).HorizontalAlignment = xlCenter
    Next c
    Wbs.Range(Wbs.Cells(4, 26), Wbs.Cells(5, g_lastDateCol)).Font.Bold = True
End Sub

Public Sub DrawGanttChart(ByVal r As Long)
    Dim s As Date, e As Date, firstCol As Long, lastCol As Long, bar As Range, p As Double
    If g_lastDateCol = 0 Then Exit Sub
    If Len(Trim$(CStr(Wbs.Cells(r, 3).Value2))) = 0 Then Exit Sub
    If Not ReadWbsDate(r, 7, s) Or Not ReadWbsDate(r, 8, e) Then Exit Sub
    If e < s Then Exit Sub
    firstCol = 26 + DateDiff("d", g_firstDate, s): lastCol = 26 + DateDiff("d", g_firstDate, e)
    Set bar = Wbs.Range(Wbs.Cells(r, firstCol), Wbs.Cells(r, lastCol))
    If IsNumeric(Wbs.Cells(r, 14).Value2) Then p = Wbs.Cells(r, 14).Value2
    If Wbs.Cells(r, 24).Value2 = "P" Then
        With bar.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Color = RGB(139, 0, 0): .Weight = xlMedium: End With
        With bar.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Color = RGB(139, 0, 0): .Weight = xlThin: End With
        bar.Cells(1, 1).Value = ProgressChart.ProgressLabel(CStr(Wbs.Cells(r, 3).Value2))
        bar.Font.Bold = True
    Else
        bar.Interior.Color = IIf(p >= 1, RGB(98, 195, 132), RGB(198, 220, 246))
        bar.Merge: bar.Value = p: bar.NumberFormat = "0%": bar.HorizontalAlignment = xlCenter
    End If
End Sub

Public Sub DrawCurrentDayLine()
    Dim c As Long
    If g_lastDateCol = 0 Then Exit Sub
    If Date < g_firstDate Or Date > g_lastDate Then Exit Sub
    c = 26 + DateDiff("d", g_firstDate, Date)
    With Wbs.Range(Wbs.Cells(6, c), Wbs.Cells(g_lastRow, c)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Color = RGB(37, 99, 235): .Weight = xlMedium
    End With
End Sub
