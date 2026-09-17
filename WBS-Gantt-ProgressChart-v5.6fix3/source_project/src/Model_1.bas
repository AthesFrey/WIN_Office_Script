Attribute VB_Name = "Model_1"
Option Explicit

Public Const g_max_level As Long = 5
Public Const PROGRESS_FORMAT As String = "0.00%"
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
Private g_beforeEdit As Object
Private g_pendingEdit As Object
Private g_undoEdits As Collection
Private g_redoEdits As Collection
Private g_retryState As Object
Private g_retryDirection As Boolean
Private g_replaying As Boolean
Private g_keysBound As Boolean
Private g_editLastRow As Long
Private g_historyDepth As Long
Private g_historyEvents As Boolean
Private g_updatingHistoryUI As Boolean
Private g_capturingSnapshot As Boolean
Private g_historyError As String
Private g_historyErrorShown As Boolean
Private Const HISTORY_LIMIT As Long = 3
Private Const WARNING_TAG As String = "WBS53_WARNING_"

Public Function HistoryMaintenanceActive() As Boolean
    HistoryMaintenanceActive = (g_historyDepth > 0)
End Function

Public Sub BeginHistoryMaintenance()
    If g_historyDepth = 0 Then
        g_historyEvents = Application.EnableEvents
        Application.EnableEvents = False
    End If
    g_historyDepth = g_historyDepth + 1
End Sub

Public Sub EndHistoryMaintenance()
    If g_historyDepth = 0 Then Exit Sub
    g_historyDepth = g_historyDepth - 1
    If g_historyDepth = 0 Then Application.EnableEvents = g_historyEvents
End Sub

Public Function LastHistoryError() As String
    LastHistoryError = g_historyError
End Function

Public Sub RememberHistoryFailure(ByVal context As String, ByVal number As Long, ByVal description As String)
    If Len(g_historyError) > 0 Then Exit Sub
    g_historyError = context & " (" & number & "): " & description
    Debug.Print g_historyError
End Sub

Public Sub ReportHistoryFailure(ByVal context As String, ByVal number As Long, ByVal description As String)
    ' Never start another snapshot or show a dialog on every selection.
    RememberHistoryFailure context, number, description
    If g_historyErrorShown Then Exit Sub
    g_historyErrorShown = True
    MsgBox "History could not be initialized. " & g_historyError, vbExclamation, "WBS history"
End Sub

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
    ClearHistory
    ThisWorkbook.Activate: Wbs.Select Replace:=True
    RefreshParentTasks
    CheckScheduleWarnings
    FindTimelineRange
    ClearGanttChart
    BuildTimeline
    For r = 6 To g_lastRow: DrawGanttChart r: Next r
    DrawCurrentDayLine
    FormatWbsBody
    ResetUndoTracking
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    UpdateHistoryUI
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

Private Sub FormatWbsBody()
    Dim lastCol As Long
    lastCol = Application.Max(25, Wbs.UsedRange.Column + Wbs.UsedRange.Columns.Count - 1)
    With Wbs.Range(Wbs.Cells(6, 2), Wbs.Cells(Application.Max(6, g_lastRow), lastCol)).Font
        .Name = "Microsoft YaHei": .Size = 11
    End With
End Sub

Public Sub CaptureUndoSnapshot(Optional ByVal throughRow As Long = 0)
    ' Selection updates the next edit's baseline, never a committed history entry.
    If g_replaying Then Exit Sub
    Set g_beforeEdit = Nothing
    ValidateHistory
    Set g_beforeEdit = TakeSnapshot(throughRow)
End Sub

Private Function TakeSnapshot(Optional ByVal throughRow As Long = 0) As Object
    Dim snapshot As Object, lastRow As Long
    lastRow = Application.Min(Wbs.Rows.Count, Application.Max(1000, throughRow + 200, _
        Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row + 200, Wbs.UsedRange.Row + Wbs.UsedRange.Rows.Count - 1))
    Set snapshot = CaptureRangeState(Wbs.Range("A6:Y" & lastRow), Wbs.Range("O6:O" & lastRow))
    snapshot.Add "lastRow", lastRow
    snapshot.Add "level", Wbs.Range("D2").Value2
    Set TakeSnapshot = snapshot
End Function

Public Function CaptureRangeState(ByVal area As Range, Optional ByVal validationArea As Range = Nothing) As Object
    Dim snapshot As Object, formats As Collection, numberFormats As Collection
    Dim formulas As Collection, texts As Collection, cell As Range
    Dim values As Variant, formulaValues As Variant
    Dim r As Long, c As Long, fieldId As Long
    Dim errorNumber As Long, errorText As String, errorSource As String
    If g_capturingSnapshot Then Err.Raise vbObjectError + 513, "History snapshot", "A snapshot is already being captured."
    BeginHistoryMaintenance
    g_capturingSnapshot = True
    On Error GoTo Failed
    If area.Areas.Count <> 1 Then Err.Raise vbObjectError + 513, "History snapshot", "Capture one rectangular area at a time."
    Set snapshot = CreateObject("Scripting.Dictionary")
    Set formats = New Collection: Set numberFormats = New Collection
    Set formulas = New Collection: Set texts = New Collection
    snapshot.Add "address", area.Address
    ' Range.Value2/Formula are scalars for one cell, arrays otherwise.
    If area.CountLarge = 1 Then
        ReDim values(1 To 1, 1 To 1): ReDim formulaValues(1 To 1, 1 To 1)
        values(1, 1) = area.Value2: formulaValues(1, 1) = area.Formula
    Else
        values = area.Value2: formulaValues = area.Formula
    End If
    snapshot.Add "content", formulaValues
    ' Bulk replay only raw numeric/Boolean/error values and blanks. Text and
    ' formulas need different write paths; neither should reparse the other.
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If VarType(formulaValues(r, c)) = vbString Then
                If Len(formulaValues(r, c)) > 0 Then
                    Set cell = area.Cells(r, c)
                    If cell.HasFormula Then
                        formulas.Add Array(cell.Address, formulaValues(r, c), CStr(cell.NumberFormat))
                    Else
                        texts.Add Array(cell.Address, values(r, c), CStr(cell.NumberFormat))
                    End If
                End If
                values(r, c) = Empty
            End If
        Next c
    Next r
    snapshot.Add "values", values
    snapshot.Add "formulas", formulas
    snapshot.Add "texts", texts
    CaptureNumberFormats area, numberFormats
    snapshot.Add "numberFormats", numberFormats
    For fieldId = 2 To 12: CaptureFormatRuns area, fieldId, formats: Next fieldId
    If Not validationArea Is Nothing Then CaptureFormatRuns validationArea, 13, formats
    For fieldId = 14 To 21: CaptureFormatRuns area, fieldId, formats: Next fieldId
    snapshot.Add "formats", formats
    Set CaptureRangeState = snapshot
Done:
    On Error GoTo 0
    g_capturingSnapshot = False
    EndHistoryMaintenance
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorText
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description: errorSource = Err.Source
    Resume Done
End Function

Private Sub CaptureNumberFormats(ByVal area As Range, ByVal formats As Collection)
    Dim value As Variant, pending As Collection, current As Range
    Set pending = New Collection: pending.Add area
    Do While pending.Count > 0
        Set current = pending(pending.Count): pending.Remove pending.Count
        value = current.NumberFormat
        If IsNull(value) Then
            QueueFormatParts current, 1, pending
        Else
            If VarType(value) <> vbString Then Err.Raise vbObjectError + 507, "History snapshot", "Invalid number format at " & current.Address
            If Len(value) = 0 Then Err.Raise vbObjectError + 507, "History snapshot", "Missing number format at " & current.Address
            formats.Add Array(current.Address, CStr(value))
        End If
    Loop
End Sub

Private Sub QueueFormatParts(ByVal area As Range, ByVal fieldId As Long, ByVal pending As Collection)
    Dim half As Long, first As Range, second As Range
    If area.Columns.Count > 1 And fieldId <> 12 Then
        half = area.Columns.Count \ 2
        Set first = area.Resize(, half)
        Set second = area.Offset(, half).Resize(, area.Columns.Count - half)
    ElseIf area.Rows.Count > 1 Then
        half = area.Rows.Count \ 2
        Set first = area.Resize(half)
        Set second = area.Offset(half).Resize(area.Rows.Count - half)
    Else
        Err.Raise vbObjectError + 507, "History snapshot", "Cannot read format " & fieldId & " at " & area.Address
    End If
    If first.CountLarge >= area.CountLarge Or second.CountLarge >= area.CountLarge Then
        Err.Raise vbObjectError + 507, "History snapshot", "Cannot split " & area.Address
    End If
    ' LIFO traversal keeps only the unfinished siblings in memory.
    pending.Add second: pending.Add first
End Sub

Private Sub RestoreSnapshotEntry(ByVal ws As Worksheet, ByVal item As Variant, ByVal isFormula As Boolean)
    Dim cell As Range, errorNumber As Long, errorText As String
    Set cell = ws.Range(CStr(item(0)))
    On Error GoTo Failed
    If isFormula Then
        cell.NumberFormat = "General"
        cell.Formula = item(1)
    Else
        cell.NumberFormat = "@"
        cell.Value2 = item(1)
    End If
    cell.NumberFormat = CStr(item(2))
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    cell.NumberFormat = CStr(item(2))
    On Error GoTo 0
    Err.Raise errorNumber, "WBS Undo", errorText
End Sub

Private Function SnapshotProperty(ByVal area As Range, ByVal fieldId As Long) As Variant
    Select Case fieldId
        Case 2: SnapshotProperty = area.Font.Bold
        Case 3: SnapshotProperty = area.Font.Name
        Case 4: SnapshotProperty = area.Font.Size
        Case 5: SnapshotProperty = area.Font.Italic
        Case 6: SnapshotProperty = area.Font.Color
        Case 7: SnapshotProperty = area.Interior.Color
        Case 8: SnapshotProperty = area.Interior.PatternColor
        Case 9: SnapshotProperty = area.Interior.Pattern
        Case 10: SnapshotProperty = area.WrapText
        Case 11: SnapshotProperty = area.HorizontalAlignment
        Case 12: SnapshotProperty = area.RowHeight
        Case 13: SnapshotProperty = SnapshotValidation(area)
        Case 14: SnapshotProperty = area.VerticalAlignment
        Case 15: SnapshotProperty = area.Font.Underline
        Case 16: SnapshotProperty = area.Font.Strikethrough
        Case 17: SnapshotProperty = area.Orientation
        Case 18: SnapshotProperty = area.ShrinkToFit
        Case 19: SnapshotProperty = area.IndentLevel
        Case 20: SnapshotProperty = area.ReadingOrder
        Case 21: SnapshotProperty = SnapshotBorders(area)
    End Select
    ' Color itself returns zero on some mixed ranges; ColorIndex reports Null.
    Select Case fieldId
        Case 6: If IsNull(area.Font.ColorIndex) Then SnapshotProperty = Null
        Case 7: If IsNull(area.Interior.ColorIndex) Then SnapshotProperty = Null
        Case 8: If IsNull(area.Interior.PatternColorIndex) Then SnapshotProperty = Null
    End Select
End Function

Private Function BorderValue(ByVal border As Border) As Variant
    Dim style As Variant, weight As Variant, color As Variant
    style = border.LineStyle
    If IsNull(style) Then BorderValue = Null: Exit Function
    If style = xlLineStyleNone Then BorderValue = Array(style): Exit Function
    weight = border.Weight: color = border.Color
    If IsNull(weight) Or IsNull(color) Or IsNull(border.ColorIndex) Then BorderValue = Null: Exit Function
    BorderValue = Array(style, weight, color)
End Function

Private Function SnapshotBorders(ByVal area As Range) As Variant
    Dim edges As Variant, values(0 To 5) As Variant, i As Long, inside As Variant
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalDown, xlDiagonalUp)
    For i = 0 To 5
        values(i) = BorderValue(area.Borders(edges(i)))
        If IsNull(values(i)) Then SnapshotBorders = Null: Exit Function
    Next i
    ' Collapse only ranges where every cell has the same edge properties.
    ' Checking inside edges prevents a table outline from becoming a grid.
    If area.Columns.Count > 1 Then
        inside = BorderValue(area.Borders(xlInsideVertical))
        If Not SameValue(inside, values(0)) Or Not SameValue(inside, values(3)) Then SnapshotBorders = Null: Exit Function
    End If
    If area.Rows.Count > 1 Then
        inside = BorderValue(area.Borders(xlInsideHorizontal))
        If Not SameValue(inside, values(1)) Or Not SameValue(inside, values(2)) Then SnapshotBorders = Null: Exit Function
    End If
    SnapshotBorders = values
End Function

Private Sub RestoreBorder(ByVal border As Border, ByVal value As Variant)
    If value(0) = xlLineStyleNone Then
        border.LineStyle = xlLineStyleNone
    Else
        border.Weight = value(1): border.Color = value(2)
        border.LineStyle = value(0)
    End If
End Sub

Private Sub RestoreBorders(ByVal area As Range, ByVal values As Variant)
    Dim edges As Variant, i As Long
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalDown, xlDiagonalUp)
    For i = 0 To 5: RestoreBorder area.Borders(edges(i)), values(i): Next i
    If area.Columns.Count > 1 Then RestoreBorder area.Borders(xlInsideVertical), values(0)
    If area.Rows.Count > 1 Then RestoreBorder area.Borders(xlInsideHorizontal), values(1)
End Sub

Private Function SnapshotValidation(ByVal area As Range) As Variant
    Dim value As Variant, field As Variant, validated As Range
    Dim errorNumber As Long, errorText As String
    ' A single-cell SpecialCells query may search the whole sheet, so always
    ' intersect its result. The caller holds the shared event guard.
    On Error Resume Next
    Set validated = area.SpecialCells(xlCellTypeAllValidation)
    errorNumber = Err.Number: errorText = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 And errorNumber <> 1004 Then Err.Raise errorNumber, "History validation", errorText
    If Not validated Is Nothing Then Set validated = Application.Intersect(area, validated)
    If validated Is Nothing Then SnapshotValidation = Array(-1): Exit Function
    If validated.CountLarge <> area.CountLarge Then SnapshotValidation = Null: Exit Function
    On Error GoTo Unreadable
    With area.Validation
        value = Array(.Type, xlValidAlertStop, xlBetween, "", "", .IgnoreBlank, False, .ShowInput, .ShowError, .InputTitle, .InputMessage, .ErrorTitle, .ErrorMessage)
        If IsNull(value(0)) Then SnapshotValidation = Null: Exit Function
        If value(0) <> xlValidateInputOnly Then
            value(1) = .AlertStyle: value(3) = .Formula1
            If value(0) = xlValidateList Then
                value(6) = .InCellDropdown
            ElseIf value(0) <> xlValidateCustom Then
                value(2) = .Operator
                If IsNull(value(2)) Then SnapshotValidation = Null: Exit Function
                If value(2) = xlBetween Or value(2) = xlNotBetween Then value(4) = .Formula2
            End If
        End If
    End With
    For Each field In value
        If IsNull(field) Then SnapshotValidation = Null: Exit Function
    Next field
    SnapshotValidation = value
    Exit Function
Unreadable:
    errorNumber = Err.Number: errorText = Err.Description
    If area.CountLarge > 1 Then SnapshotValidation = Null: Exit Function
    Err.Raise errorNumber, "History validation", "Cannot read validation at " & area.Address & ": " & errorText
End Function

Private Sub RestoreSnapshotValidation(ByVal area As Range, ByVal value As Variant)
    With area.Validation
        .Delete
        If value(0) = -1 Then Exit Sub
        Select Case value(0)
            Case xlValidateInputOnly: .Add Type:=value(0)
            Case xlValidateList, xlValidateCustom
                .Add Type:=value(0), AlertStyle:=value(1), Formula1:=value(3)
            Case Else
                .Add Type:=value(0), AlertStyle:=value(1), Operator:=value(2), Formula1:=value(3), Formula2:=value(4)
        End Select
        .IgnoreBlank = value(5)
        If value(0) = xlValidateList Then .InCellDropdown = value(6)
        .ShowInput = value(7): .ShowError = value(8)
        .InputTitle = value(9): .InputMessage = value(10)
        .ErrorTitle = value(11): .ErrorMessage = value(12)
    End With
End Sub

Private Sub CaptureFormatRuns(ByVal area As Range, ByVal fieldId As Long, ByVal formats As Collection)
    Dim value As Variant, pending As Collection, current As Range
    Set pending = New Collection: pending.Add area
    Do While pending.Count > 0
        Set current = pending(pending.Count): pending.Remove pending.Count
        value = SnapshotProperty(current, fieldId)
        If IsNull(value) Then
            QueueFormatParts current, fieldId, pending
        Else
            formats.Add Array(current.Address, fieldId, value)
        End If
    Loop
End Sub

Private Sub RestoreSnapshotFormat(ByVal ws As Worksheet, ByVal item As Variant)
    With ws.Range(CStr(item(0)))
        Select Case CLng(item(1))
            Case 2: .Font.Bold = item(2)
            Case 3: .Font.Name = item(2)
            Case 4: .Font.Size = item(2)
            Case 5: .Font.Italic = item(2)
            Case 6: .Font.Color = item(2)
            Case 7: .Interior.Color = item(2)
            Case 8: .Interior.PatternColor = item(2)
            Case 9: .Interior.Pattern = item(2)
            Case 10: .WrapText = item(2)
            Case 11: .HorizontalAlignment = item(2)
            Case 12: .RowHeight = item(2)
            Case 13: RestoreSnapshotValidation ws.Range(CStr(item(0))), item(2)
            Case 14: .VerticalAlignment = item(2)
            Case 15: .Font.Underline = item(2)
            Case 16: .Font.Strikethrough = item(2)
            Case 17: .Orientation = item(2)
            Case 18: .ShrinkToFit = item(2)
            Case 19: .IndentLevel = item(2)
            Case 20: .ReadingOrder = item(2)
            Case 21: RestoreBorders ws.Range(CStr(item(0))), item(2)
        End Select
    End With
End Sub

Private Sub EnsureHistory()
    If g_undoEdits Is Nothing Then Set g_undoEdits = New Collection
    If g_redoEdits Is Nothing Then Set g_redoEdits = New Collection
End Sub

Public Function UndoCount() As Long
    EnsureHistory
    UndoCount = g_undoEdits.Count
End Function

Public Function RedoCount() As Long
    EnsureHistory
    RedoCount = g_redoEdits.Count
End Function

Private Function HistoryState() As Object
    Dim entry As Object
    EnsureHistory
    If Not g_retryState Is Nothing Then
        Set HistoryState = g_retryState
    ElseIf g_redoEdits.Count > 0 Then
        Set entry = g_redoEdits(g_redoEdits.Count)
        Set HistoryState = entry("before")
    ElseIf g_undoEdits.Count > 0 Then
        Set entry = g_undoEdits(g_undoEdits.Count)
        Set HistoryState = entry("after")
    End If
End Function

Private Function SameValue(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim i As Long
    If VarType(first) <> VarType(second) Then Exit Function
    If IsArray(first) Then
        If LBound(first) <> LBound(second) Or UBound(first) <> UBound(second) Then Exit Function
        For i = LBound(first) To UBound(first)
            If Not SameValue(first(i), second(i)) Then Exit Function
        Next i
        SameValue = True
    ElseIf IsNull(first) Or IsEmpty(first) Then
        SameValue = True
    ElseIf IsNumeric(first) And VarType(first) <> vbString Then
        SameValue = (first = second)
    Else
        SameValue = (StrComp(CStr(first), CStr(second), vbBinaryCompare) = 0)
    End If
End Function

Private Function OutsideEdit(ByVal area As Range, ByVal changed As Range) As Boolean
    OutsideEdit = True
    If changed Is Nothing Then Exit Function
    OutsideEdit = (Application.Intersect(area, changed) Is Nothing)
End Function

Private Function FormatMatches(ByVal area As Range, ByVal fieldId As Long, ByVal expected As Variant, ByVal changed As Range) As Boolean
    Dim actual As Variant, overlap As Range, current As Range, pending As Collection
    Set pending = New Collection: pending.Add area
    Do While pending.Count > 0
        Set current = pending(pending.Count): pending.Remove pending.Count
        If fieldId = 1 Then actual = current.NumberFormat Else actual = SnapshotProperty(current, fieldId)
        If Not SameValue(actual, expected) Then
            If changed Is Nothing Then Exit Function
            ' A paste can resize the entire destination row.
            If fieldId = 12 Then Set overlap = Application.Intersect(current, changed.EntireRow) Else Set overlap = Application.Intersect(current, changed)
            If overlap Is Nothing Then Exit Function
            If overlap.CountLarge <> current.CountLarge Then QueueFormatParts current, fieldId, pending
        End If
    Loop
    FormatMatches = True
End Function

Private Function SnapshotMatches(ByVal snapshot As Object, Optional ByVal changed As Range = Nothing) As Boolean
    Dim lastRow As Long, lastUsed As Long
    If snapshot Is Nothing Then Exit Function
    lastRow = CLng(snapshot("lastRow"))
    If Not RangeStateMatches(Wbs, snapshot, changed) Then Exit Function
    lastUsed = Wbs.UsedRange.Row + Wbs.UsedRange.Rows.Count - 1
    If lastUsed > lastRow Then
        If Application.CountA(Wbs.Range("A" & lastRow + 1 & ":Y" & lastUsed)) > 0 Then Exit Function
    End If
    ' D2 reflects selection and is deliberately not an edit-consistency check.
    SnapshotMatches = True
End Function

Public Function RangeStateMatches(ByVal ws As Worksheet, ByVal snapshot As Object, Optional ByVal changed As Range = Nothing) As Boolean
    Dim errorNumber As Long, errorText As String, errorSource As String
    BeginHistoryMaintenance
    On Error GoTo Failed
    RangeStateMatches = MatchesRangeState(ws, snapshot, changed)
Done:
    On Error GoTo 0
    EndHistoryMaintenance
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorText
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description: errorSource = Err.Source
    Resume Done
End Function

Private Function MatchesRangeState(ByVal ws As Worksheet, ByVal snapshot As Object, ByVal changed As Range) As Boolean
    Dim actual As Variant, expected As Variant, r As Long, c As Long, area As Range
    Dim items As Collection, item As Variant
    If snapshot Is Nothing Then Exit Function
    Set area = ws.Range(CStr(snapshot("address")))
    If area.CountLarge = 1 Then
        ReDim actual(1 To 1, 1 To 1): actual(1, 1) = area.Formula
    Else
        actual = area.Formula
    End If
    expected = snapshot("content")
    For r = 1 To UBound(expected, 1)
        For c = 1 To UBound(expected, 2)
            If Not SameValue(actual(r, c), expected(r, c)) Then
                If OutsideEdit(area.Cells(r, c), changed) Then Exit Function
            End If
        Next c
    Next r
    ' Formula text and an equal-sign literal can have identical .Formula strings.
    Set items = snapshot("formulas")
    For Each item In items
        If OutsideEdit(ws.Range(CStr(item(0))), changed) Then
            If Not ws.Range(CStr(item(0))).HasFormula Then Exit Function
        End If
    Next item
    Set items = snapshot("texts")
    For Each item In items
        If OutsideEdit(ws.Range(CStr(item(0))), changed) Then
            If ws.Range(CStr(item(0))).HasFormula Then Exit Function
        End If
    Next item
    Set items = snapshot("numberFormats")
    For Each item In items
        If Not FormatMatches(ws.Range(CStr(item(0))), 1, item(1), changed) Then Exit Function
    Next item
    Set items = snapshot("formats")
    For Each item In items
        If Not FormatMatches(ws.Range(CStr(item(0))), CLng(item(1)), item(2), changed) Then Exit Function
    Next item
    MatchesRangeState = True
End Function

Public Sub ValidateHistory()
    Dim expected As Object
    If g_replaying Then Exit Sub
    Set expected = HistoryState
    If expected Is Nothing Then Exit Sub
    If Not SnapshotMatches(expected) Then
        ClearHistory
        Set g_beforeEdit = TakeSnapshot
    End If
End Sub

Public Sub ClearHistory()
    Set g_undoEdits = New Collection: Set g_redoEdits = New Collection
    Set g_pendingEdit = Nothing: Set g_retryState = Nothing
    Set g_beforeEdit = Nothing
End Sub

Public Sub BeginWbsEdit(ByVal lastChangedRow As Long, Optional ByVal changed As Range = Nothing)
    Dim baseline As Object
    EnsureHistory
    Set g_pendingEdit = Nothing
    If Not g_retryState Is Nothing Then
        Set baseline = g_beforeEdit
        ClearHistory
        Set g_beforeEdit = baseline
    End If
    ' Change events arrive after input. Check the untouched cells before the
    ' scheduler writes linked results, so old snapshots never absorb outside edits.
    If Not changed Is Nothing And Not g_beforeEdit Is Nothing Then
        If lastChangedRow <= CLng(g_beforeEdit("lastRow")) Then
            If Not SnapshotMatches(g_beforeEdit, changed) Then ClearHistory
        End If
    End If
    ' Preserve existing stacks until a changed edit has actually completed.
    Set g_pendingEdit = g_beforeEdit
    g_editLastRow = lastChangedRow
End Sub

Public Sub FinishWbsEdit()
    Dim after As Object, entry As Object, warning As String
    On Error GoTo Failed
    If g_pendingEdit Is Nothing Then
        ResetUndoTracking
        Exit Sub
    End If
    ' A paste beyond the pre-edit capture cannot safely invent old formats.
    If g_editLastRow > CLng(g_pendingEdit("lastRow")) Then
        warning = "This edit extends beyond the captured rows. History was cleared. Select the full destination before pasting to retain Undo."
        GoTo FailedEdit
    End If
    If Not SnapshotMatches(g_pendingEdit) Then
        Set after = TakeSnapshot(g_editLastRow)
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "before", g_pendingEdit
        entry.Add "after", after
        Set g_redoEdits = New Collection
        Set g_retryState = Nothing
        g_undoEdits.Add entry
        If g_undoEdits.Count > HISTORY_LIMIT Then g_undoEdits.Remove 1
        Set g_beforeEdit = after
    Else
        Set g_beforeEdit = TakeSnapshot(g_editLastRow)
    End If
    Set g_pendingEdit = Nothing
    UpdateHistoryUI
    Exit Sub
Failed:
    warning = "History could not be recorded: " & Err.Description
FailedEdit:
    On Error Resume Next
    ClearHistory
    Set g_beforeEdit = Nothing
    Set g_beforeEdit = TakeSnapshot
    UpdateHistoryUI
    On Error GoTo 0
    MsgBox warning, vbExclamation, "WBS history"
End Sub

Public Sub ResetUndoTracking()
    ClearHistory
    Set g_beforeEdit = Nothing
    Set g_beforeEdit = TakeSnapshot
    UpdateHistoryUI
End Sub

Public Function HistoryContextActive() As Boolean
    If ActiveWorkbook Is Nothing Then Exit Function
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Function
    If Not ActiveSheet Is Wbs Then Exit Function
    HistoryContextActive = True
End Function

Public Sub ReleaseHistoryKeys()
    If Not g_keysBound Then Exit Sub
    Application.OnKey "^z"
    Application.OnKey "^y"
    g_keysBound = False
End Sub

Public Sub UpdateHistoryUI()
    Dim prefix As String, suffix As String, canUndo As Boolean, canRedo As Boolean
    Dim errorNumber As Long, errorText As String
    If g_updatingHistoryUI Then Exit Sub
    BeginHistoryMaintenance
    g_updatingHistoryUI = True
    On Error GoTo Failed
    EnsureHistory
    canUndo = (g_undoEdits.Count > 0): canRedo = (g_redoEdits.Count > 0)
    If Not g_retryState Is Nothing Then
        canUndo = canUndo And Not g_retryDirection
        canRedo = canRedo And g_retryDirection
    End If
    ' Missing/protected shapes must not prevent keyboard recovery or cleanup.
    On Error Resume Next
    SetHistoryButton "UndoWbs", "Undo", g_undoEdits.Count, canUndo
    SetHistoryButton "RedoWbs", "Redo", g_redoEdits.Count, canRedo
    Sheet4.UpdateHistoryUI
    On Error GoTo Failed
    If Not ActiveWorkbook Is Nothing Then
        If ActiveWorkbook Is ThisWorkbook Then
            If ActiveSheet Is Wbs Then suffix = "Wbs"
            If ActiveSheet Is Sheet4 Then suffix = "Milestones"
        End If
    End If
    If Len(suffix) > 0 Then
        prefix = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ProgressChart."
        Application.OnKey "^z", prefix & "Undo" & suffix
        Application.OnKey "^y", prefix & "Redo" & suffix
        g_keysBound = True
    Else
        ReleaseHistoryKeys
    End If
Done:
    On Error GoTo 0
    g_updatingHistoryUI = False
    EndHistoryMaintenance
    If errorNumber <> 0 Then ReportHistoryFailure "History controls", errorNumber, errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    ReleaseHistoryKeys
    On Error GoTo 0
    GoTo Done
End Sub

Private Sub SetHistoryButton(ByVal name As String, ByVal label As String, ByVal count As Long, ByVal enabled As Boolean)
    Dim shp As Shape, caption As String, color As Long
    Set shp = Wbs.Shapes("PG_Button_" & name)
    caption = label & " (" & count & ")"
    color = IIf(enabled, RGB(37, 99, 235), RGB(148, 163, 184))
    If shp.TextFrame2.TextRange.Text <> caption Then shp.TextFrame2.TextRange.Text = caption
    If shp.Fill.ForeColor.RGB <> color Then shp.Fill.ForeColor.RGB = color
End Sub

Private Sub RestoreSnapshot(ByVal snapshot As Object, ByVal clearThrough As Long)
    Dim lastRow As Long
    lastRow = CLng(snapshot("lastRow"))
    If clearThrough > lastRow Then Wbs.Range("A" & lastRow + 1 & ":Y" & clearThrough).ClearContents
    RestoreRangeState Wbs, snapshot
    Wbs.Range("D2").Value2 = snapshot("level")
    Wbs.Calculate
    ThisWorkbook.Worksheets("Milestones").Calculate
End Sub

Public Sub RestoreRangeState(ByVal ws As Worksheet, ByVal snapshot As Object)
    Dim item As Variant, items As Collection
    Set items = snapshot("numberFormats")
    For Each item In items: ws.Range(CStr(item(0))).NumberFormat = CStr(item(1)): Next item
    ws.Range(CStr(snapshot("address"))).Value2 = snapshot("values")
    Set items = snapshot("texts")
    For Each item In items: RestoreSnapshotEntry ws, item, False: Next item
    Set items = snapshot("formulas")
    For Each item In items: RestoreSnapshotEntry ws, item, True: Next item
    Set items = snapshot("formats")
    For Each item In items: RestoreSnapshotFormat ws, item: Next item
End Sub

Public Sub UndoLastEdit(Optional ByVal propagateErrors As Boolean = False)
    ReplayHistory False, propagateErrors
End Sub

Public Sub RedoLastEdit(Optional ByVal propagateErrors As Boolean = False)
    ReplayHistory True, propagateErrors
End Sub

Private Sub ReplayHistory(ByVal redo As Boolean, ByVal propagateErrors As Boolean)
    Dim source As Collection, destination As Collection, entry As Object, target As Object, original As Object
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation
    Dim errorText As String, clearThrough As Long, started As Boolean, restored As Object
    If g_replaying Or Not HistoryContextActive Then Exit Sub
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    EnsureHistory
    ValidateHistory
    If Not g_retryState Is Nothing Then
        If redo <> g_retryDirection Then GoTo Done
    End If
    If redo Then
        Set source = g_redoEdits: Set destination = g_undoEdits
    Else
        Set source = g_undoEdits: Set destination = g_redoEdits
    End If
    If source.Count = 0 Then GoTo Done
    Set entry = source(source.Count)
    If redo Then Set target = entry("after") Else Set target = entry("before")
    Set original = TakeSnapshot
    If Wbs.ProtectContents Then Err.Raise vbObjectError + 508, , "Unprotect WBS before retrying."
    clearThrough = Application.Max(CLng(target("lastRow")), CLng(original("lastRow")))
    g_replaying = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    started = True
    RestoreSnapshot target, clearThrough
    If Not SnapshotMatches(target) Then Err.Raise vbObjectError + 508, , "The restored body could not be verified."
    Set restored = TakeSnapshot
    ' Commit only after both replay and the next edit baseline have succeeded.
    destination.Add entry
    source.Remove source.Count
    Set g_beforeEdit = restored
    Set g_retryState = Nothing: Set g_pendingEdit = Nothing
Done:
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    g_replaying = False
    UpdateHistoryUI
    On Error GoTo 0
    If Len(errorText) > 0 Then
        If propagateErrors Then Err.Raise vbObjectError + 508, "WBS history", errorText
        MsgBox errorText & vbCrLf & "History is retained. Resolve the error, then retry.", vbExclamation, "WBS history"
    End If
    Exit Sub
Failed:
    errorText = IIf(redo, "Redo", "Undo") & " failed: " & Err.Description
    On Error Resume Next
    If started Then RestoreSnapshot original, clearThrough
    ' Keep the actual failed/rolled-back state so later external edits invalidate retry.
    If Not original Is Nothing Then
        Set g_retryState = TakeSnapshot
        g_retryDirection = redo
        Set g_beforeEdit = g_retryState
    End If
    On Error GoTo 0
    GoTo Done
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
    Wbs.Range("N6:N" & g_lastRow).NumberFormat = PROGRESS_FORMAT
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
        With Wbs.Range("O6:O" & Application.Min(Wbs.Rows.Count, Application.Max(1000, g_lastRow + 200))).Validation
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
    If g_pendingEdit Is Nothing Then ResetUndoTracking
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

Public Sub CheckScheduleWarnings()
    Dim r As Long, parent As Long, previous As Long, lastScan As Long, i As Long
    Dim previousSibling() As Long, startDate As Date, endDate As Date, previousEnd As Date
    Dim value As Variant, rules As FormatConditions
    Wbs.Calculate
    g_lastRow = Wbs.Cells(Wbs.Rows.Count, 3).End(xlUp).Row
    lastScan = Application.Max(6, Wbs.UsedRange.Row + Wbs.UsedRange.Rows.Count - 1)
    Set rules = Wbs.Range("F6:G" & lastScan).FormatConditions
    For i = rules.Count To 1 Step -1
        If rules(i).Type = xlExpression Then
            If InStr(1, rules(i).Formula1, WARNING_TAG, vbBinaryCompare) > 0 Then rules(i).Delete
        End If
    Next i
    If g_lastRow < 6 Then Exit Sub
    LoadTaskValues: BuildTaskIndex
    ReDim previousSibling(0 To g_lastRow)
    For r = 6 To g_lastRow
        If g_levels(r) > 0 Then
            value = Wbs.Cells(r, 6).Value2
            If Not IsError(value) And Not IsEmpty(value) Then
                If IsNumeric(value) Then
                    If ReadWbsDate(r, 7, startDate) And ReadWbsDate(r, 8, endDate) Then
                        If CDbl(value) <> DateDiff("d", startDate, endDate) + 1 Then AddScheduleWarning Wbs.Cells(r, 6), "DAYS", RGB(255, 199, 206)
                    End If
                End If
            End If
            parent = g_parents(r): previous = previousSibling(parent)
            ' VBA And does not short-circuit: never index the sentinel row zero.
            If previous > 0 And r = previous + 1 Then
                If g_levels(previous) = g_levels(r) Then
                    If ReadWbsDate(r, 7, startDate) And ReadWbsDate(previous, 8, previousEnd) Then
                        If CDbl(startDate) - CDbl(previousEnd) <> 1 Then AddScheduleWarning Wbs.Cells(r, 7), "SEQUENCE", RGB(255, 235, 156)
                    End If
                End If
            End If
            previousSibling(parent) = r
        End If
    Next r
End Sub

Private Sub AddScheduleWarning(ByVal cell As Range, ByVal kind As String, ByVal color As Long)
    Dim rule As FormatCondition
    ' Own only this tagged rule. Removing it reveals the user's original fill.
    Set rule = cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=N(""" & WARNING_TAG & kind & """)=0")
    rule.Interior.Color = color
    rule.SetFirstPriority
    rule.StopIfTrue = False
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
    ClearHistory
    Sheet4.ResetUndoTracking
    Wbs.Rows(insertRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    With Wbs.Range(Wbs.Cells(insertRow, 2), Wbs.Cells(insertRow, 25))
        .ClearContents: .Font.Bold = False
        .Font.Name = "Microsoft YaHei": .Font.Size = 11
    End With
    Wbs.Cells(insertRow, 3).Value = String$((level - 1) * 4, " ") & "New task"
    Wbs.Cells(insertRow, 6).Value = 1: Wbs.Cells(insertRow, 14).Value = 0
    If hasStart Then Wbs.Cells(insertRow, 7).Value = startDate
    Wbs.Range("G" & insertRow & ":H" & insertRow).NumberFormat = "yyyy/m/d"
    Wbs.Cells(insertRow, 14).NumberFormat = PROGRESS_FORMAT
    RefreshParentTasks
    Application.EnableEvents = oldEvents
    Wbs.Cells(insertRow, 3).Select
Done:
    Application.EnableEvents = oldEvents
    UpdateHistoryUI
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
    Dim oldState As Object, errorNumber As Long, errorText As String, committed As Boolean
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
    Set oldState = TakeSnapshot
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
        RestoreSnapshot oldState, CLng(oldState("lastRow"))
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
        bar.Merge: bar.Value = p: bar.NumberFormat = PROGRESS_FORMAT: bar.HorizontalAlignment = xlCenter
        bar.WrapText = False: bar.ShrinkToFit = True
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
