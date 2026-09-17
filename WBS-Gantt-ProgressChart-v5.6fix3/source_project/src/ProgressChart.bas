Attribute VB_Name = "ProgressChart"
Option Explicit

Private Const PG_SHEET As String = "Milestones"
Private Const PG_FIRST As Long = 14
Private Const PG_LAST As Long = 213
Private Const PG_LIMIT As Long = 60
Private Const PG_GROUP As String = "WBS_ProgressChart"
Private Const PG_TEMP As String = "WBS_PG_TMP_"
Private Const PG_MODE_CELL As String = "C6"
Private Const PG_SCALE_CELL As String = "C7"

Public Sub OpenProgressChart()
    On Error GoTo Failed
    ThisWorkbook.Worksheets(PG_SHEET).Activate
    Exit Sub
Failed:
    MsgBox "Milestones sheet was not found.", vbExclamation
End Sub

Public Sub RefreshAll(Optional ByVal propagateErrors As Boolean = False)
    Dim oldScreen As Boolean, errorText As String, milestoneError As String, generated As Boolean
    Dim started As Boolean, rollbackError As String
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    Sheet4.BeginEdit: started = True
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets("WBS").Activate
    ThisWorkbook.Worksheets("WBS").Range("C6").Select
    UpdateGantt True
    generated = TryRefreshProgressChart(milestoneError)
    If Not generated Then
        rollbackError = Sheet4.CancelEdit: started = False
        If Len(rollbackError) > 0 Then Err.Raise vbObjectError + 509, "Refresh All", milestoneError & rollbackError
    End If
    ApplyRefreshAllFormatting
    Model_1.ResetUndoTracking
    Sheet4.ResetUndoTracking: started = False
    Application.ScreenUpdating = oldScreen
    If Not generated And Not propagateErrors Then MsgBox "Other content has been refreshed. Milestones skipped after retry: " & milestoneError, vbExclamation, "Refresh All"
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If started Then errorText = errorText & Sheet4.CancelEdit
    Sheet4.ResetUndoTracking
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If propagateErrors Then
        Err.Raise vbObjectError + 509, "Refresh All", errorText
    End If
    MsgBox "Refresh All failed: " & errorText, vbExclamation
End Sub

Public Sub UpdateGantt(Optional ByVal propagateErrors As Boolean = False): Model_1.UpdateGanttChart propagateErrors: End Sub
Public Sub AddTask(): Model_1.AddTask: End Sub
Public Sub AddSubtask(): Model_1.AddSubtask: End Sub
Public Sub PromoteTasks(): Sheet1.ChangeSelectedLevels -1: End Sub
Public Sub DemoteTasks(): Sheet1.ChangeSelectedLevels 1: End Sub
Public Sub UndoWbs(): Model_1.UndoLastEdit: End Sub
Public Sub RedoWbs(): Model_1.RedoLastEdit: End Sub
Public Sub UndoMilestones(): Sheet4.UndoLastEdit: End Sub
Public Sub RedoMilestones(): Sheet4.RedoLastEdit: End Sub
Public Sub DeleteMilestoneLevel1(): DeleteMilestonesByLevel 1: End Sub
Public Sub DeleteMilestoneLevel2(): DeleteMilestonesByLevel 2: End Sub
Public Sub DeleteMilestoneLevel3(): DeleteMilestonesByLevel 3: End Sub
Public Sub DeleteMilestoneLevel4(): DeleteMilestonesByLevel 4: End Sub
Public Sub DeleteMilestoneLevel5(): DeleteMilestonesByLevel 5: End Sub

Public Sub DeleteMilestonesByLevel(ByVal taskLevel As Long, Optional ByVal propagateErrors As Boolean = False)
    Dim ws As Worksheet, source As Worksheet, removed As Range
    Dim r As Long, sourceRow As Long, count As Long, taskValue As Variant
    Dim oldEvents As Boolean, oldScreen As Boolean, started As Boolean, errorText As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(PG_SHEET): Set source = ThisWorkbook.Worksheets("WBS")
    If Not ActiveWorkbook Is ThisWorkbook Or Not ActiveSheet Is ws Then Exit Sub
    If taskLevel < 1 Or taskLevel > Model_1.g_max_level Then Err.Raise vbObjectError + 511, , "Choose a level from 1 to 5."
    source.Calculate: ws.Calculate
    For r = PG_FIRST To PG_LAST
        If TrySourceRow(ws.Cells(r, 3).Value2, source.Rows.Count, sourceRow) Then
            taskValue = source.Cells(sourceRow, 3).Value2
            If Not IsError(taskValue) Then
                If Len(Trim$(CStr(taskValue))) > 0 Then
                    If Model_1.DetectLevel(CStr(taskValue)) = taskLevel Then
                        If removed Is Nothing Then Set removed = ws.Range("B" & r & ":J" & r) Else Set removed = Union(removed, ws.Range("B" & r & ":J" & r))
                        count = count + 1
                    End If
                End If
            End If
        End If
    Next r
    If count = 0 Then
        ws.Range("B10").Value = "No linked level " & taskLevel & " tasks found."
        Exit Sub
    End If
    If ws.ProtectContents Then Err.Raise vbObjectError + 511, , "Unprotect Milestones before deleting tasks."
    Sheet4.BeginEdit: started = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    removed.ClearContents
    ws.Range("B10").Value = "Deleted " & count & " level " & taskLevel & " entries. Click Generate to update the chart."
    Sheet4.FinishEdit: started = False
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    errorText = "Delete failed: " & Err.Description
    If started Then errorText = errorText & Sheet4.CancelEdit
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If propagateErrors Then Err.Raise vbObjectError + 511, "Milestones", errorText
    MsgBox errorText, vbExclamation, "Milestones"
End Sub

Private Function TrySourceRow(ByVal value As Variant, ByVal maxRow As Long, ByRef rowNumber As Long) As Boolean
    Dim number As Double
    On Error GoTo InvalidRow
    rowNumber = 0
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or VarType(value) = vbBoolean Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    number = CDbl(value)
    If number < 6 Or number > maxRow Or number <> Fix(number) Then Exit Function
    rowNumber = CLng(number): TrySourceRow = True
InvalidRow:
End Function

Public Sub ImportSelectedMilestones(Optional ByVal propagateErrors As Boolean = False)
    Dim source As Worksheet, target As Worksheet, picked As Range, part As Range
    Dim rows() As Long, count As Long, r As Long, i As Long, j As Long, swapRow As Long
    Dim lastRow As Long, dest As Long, freeRows As Long, duplicate As Boolean, taskValue As Variant, existingRow As Long
    Dim oldEvents As Boolean, started As Boolean, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Set source = ThisWorkbook.Worksheets("WBS")
    Set target = ThisWorkbook.Worksheets(PG_SHEET)
    If Not ActiveSheet Is source Or TypeName(Selection) <> "Range" Then
        MsgBox "Select task cells or rows on WBS first.", vbInformation
        Exit Sub
    End If
    ' Refresh source formulas before duplicate detection in manual calculation mode.
    source.Calculate: target.Calculate
    lastRow = source.Cells(source.Rows.Count, 3).End(xlUp).Row
    If lastRow < 6 Then Err.Raise vbObjectError + 512, , "There are no WBS tasks to import."
    Set picked = Application.Intersect(Selection, source.Rows("6:" & lastRow))
    If picked Is Nothing Then MsgBox "Select tasks from row 6 onward.", vbInformation: Exit Sub
    ReDim rows(1 To PG_LAST - PG_FIRST + 1)
    For Each part In picked.Areas
        For r = part.Row To part.Row + part.Rows.Count - 1
            taskValue = source.Cells(r, 3).Value
            duplicate = False
            If Not IsError(taskValue) And Len(Trim$(CStr(taskValue))) > 0 Then
                For i = 1 To count: If rows(i) = r Then duplicate = True
                Next i
                For i = PG_FIRST To PG_LAST
                    If TrySourceRow(target.Cells(i, 3).Value2, source.Rows.Count, existingRow) Then If existingRow = r Then duplicate = True
                Next i
                If Not duplicate Then
                    If count >= UBound(rows) Then MsgBox "There are too many selected tasks.", vbExclamation: Exit Sub
                    count = count + 1: rows(count) = r
                End If
            End If
        Next r
    Next part
    If count = 0 Then MsgBox "The selected tasks are already imported or have no names.", vbInformation: Exit Sub
    For i = 2 To count
        swapRow = rows(i): j = i - 1
        Do While j >= 1
            If rows(j) <= swapRow Then Exit Do
            rows(j + 1) = rows(j): j = j - 1
        Loop
        rows(j + 1) = swapRow
    Next i
    For r = PG_FIRST To PG_LAST
        If Application.CountA(target.Range("B" & r & ":J" & r)) = 0 Then freeRows = freeRows + 1
    Next r
    If count > freeRows Then MsgBox "There are not enough empty milestone rows.", vbExclamation: Exit Sub
    If target.ProtectContents Then Err.Raise vbObjectError + 512, , "Unprotect Milestones before importing tasks."
    Sheet4.BeginEdit: started = True
    Application.EnableEvents = False
    dest = PG_FIRST
    For i = 1 To count
        Do While Application.CountA(target.Range("B" & dest & ":J" & dest)) > 0: dest = dest + 1: Loop
        r = rows(i)
        target.Cells(dest, 2).Value = "Yes"
        target.Cells(dest, 3).Formula = "=ROW(WBS!C" & r & ")"
        target.Cells(dest, 4).Formula = "=IF(WBS!C" & r & "="""","""",WBS!C" & r & ")"
        target.Cells(dest, 5).Formula = "=IF(WBS!H" & r & "="""","""",WBS!H" & r & ")"
        target.Cells(dest, 5).NumberFormat = "yyyy/m/d"
        target.Cells(dest, 6).Value = "Linked to WBS"
        target.Cells(dest, 7).Formula = "=IFERROR(WBS!D" & r & ","""")"
        target.Cells(dest, 8).Formula = "=IFERROR(WBS!N" & r & ","""")"
        target.Cells(dest, 8).NumberFormat = Model_1.PROGRESS_FORMAT
        target.Cells(dest, 9).Formula = "=IF(WBS!C" & r & "="""","""",IFERROR(IF(WBS!N" & r & ">=1,""Complete"",IF(WBS!H" & r & "<TODAY(),""Overdue"",IF(WBS!G" & r & ">TODAY(),""Not started"",""In progress""))),""Check""))"
        target.Cells(dest, 10).Formula = "=IFERROR(WBS!X" & r & ","""")"
        dest = dest + 1
    Next i
    target.Calculate
    target.Range("B10").Value = "Imported " & count & ". Click Generate."
    Sheet4.FinishEdit: started = False
    Application.EnableEvents = oldEvents
    target.Activate
    Exit Sub
Failed:
    errorText = "Task import failed: " & Err.Description
    If started Then errorText = errorText & Sheet4.CancelEdit
    Application.EnableEvents = oldEvents
    If propagateErrors Then Err.Raise vbObjectError + 512, "Milestones", errorText
    MsgBox errorText, vbExclamation
End Sub

Public Sub GenerateProgressChart(Optional ByVal propagateErrors As Boolean = False)
    Dim errorText As String, started As Boolean
    On Error GoTo Failed
    Sheet4.BeginEdit: started = True
    If TryRefreshProgressChart(errorText) Then
        Sheet4.FinishEdit
        Exit Sub
    End If
    GoTo ReportFailure
Failed:
    errorText = Err.Description
ReportFailure:
    If started Then errorText = errorText & Sheet4.CancelEdit
    On Error GoTo 0
    If propagateErrors Then Err.Raise vbObjectError + 507, "Milestones", errorText
    MsgBox "Milestones skipped after retry: " & errorText, vbExclamation, "Milestones"
End Sub

Private Function TryRefreshProgressChart(ByRef errorText As String) As Boolean
    Dim ws As Worksheet, cleanupError As String
    If TryBuildProgressChart(errorText) Then TryRefreshProgressChart = True: Exit Function
    ' A failed refresh discards all generated old/partial drawings before retry.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PG_SHEET)
    If Not ws Is Nothing Then DeleteProgressDrawings ws
    On Error GoTo 0
    If TryBuildProgressChart(errorText) Then TryRefreshProgressChart = True: Exit Function
    ' A second failure leaves no stale chart and does not block Refresh All.
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then DeleteProgressDrawings ws
    If Err.Number <> 0 Then cleanupError = "; old chart cleanup failed: " & Err.Description
    errorText = errorText & cleanupError
    If Not ws Is Nothing Then ws.Range("B10").Value = "Milestones skipped after retry: " & errorText
    On Error GoTo 0
End Function

Private Function TryBuildProgressChart(ByRef errorText As String) As Boolean
    On Error GoTo Failed
    BuildProgressChart
    TryBuildProgressChart = True
    Exit Function
Failed:
    errorText = Err.Description
End Function

Private Sub BuildProgressChart()
    Dim ws As Worksheet, names() As String, dates() As Date, inputCount As Long, nodeCount As Long
    Dim r As Long, i As Long, j As Long, mark As String, taskName As String
    Dim value As Variant, milestoneDate As Date, swapDate As Date, swapName As String
    Dim lastYear As Long, shortInput As Boolean, previousDate As Date, previousShort As Boolean
    Dim sourceRow As Long, reason As String, mode As String, span As Double, chartScale As Double
    Dim width As Single, height As Single, gap As Single, left As Single, top As Single, axis As Single
    Dim centers() As Single, widths() As Single, labelLeft() As Single, labelLanes() As Long, dateLanes() As Long
    Dim laneRight() As Single, dateRight() As Single, labelShapes() As Shape, dateShapes() As Shape
    Dim labelHeight As Single, available As Single, dateLaneCount As Long, labelLaneCount As Long, lane As Long
    Dim failureText As String, markerTop As Single
    Dim shapeIndexes() As Variant, shapeCount As Long, shp As Shape, group As Shape, background As Shape, oldScreen As Boolean
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    Set ws = ThisWorkbook.Worksheets(PG_SHEET)
    ws.Range("H" & PG_FIRST & ":H" & PG_LAST).NumberFormat = Model_1.PROGRESS_FORMAT
    ThisWorkbook.Worksheets("WBS").Calculate: ws.Calculate
    ReDim names(1 To PG_LIMIT): ReDim dates(1 To PG_LIMIT)
    For r = PG_FIRST To PG_LAST
        mark = Trim$(CStr(ws.Cells(r, 2).Value2))
        If mark <> "" And mark <> "Yes" And mark <> "No" Then RejectProgressInput "Row " & r & ": Show must be Yes or No."
        If mark = "Yes" Then
            value = ws.Cells(r, 4).Value
            If IsError(value) Then RejectProgressInput "Row " & r & ": milestone name has an error."
            taskName = Trim$(CStr(value))
            sourceRow = 0
            value = ws.Cells(r, 3).Value2
            If IsError(value) Then RejectProgressInput "Row " & r & ": WBS source has an error."
            If Len(Trim$(CStr(value))) > 0 Then
                If Not TrySourceRow(value, ThisWorkbook.Worksheets("WBS").Rows.Count, sourceRow) Then
                    RejectProgressInput "Row " & r & ": WBS source must be a whole-number task row from 6 onward."
                End If
                If Not ValidateImportedTask(sourceRow, reason) Then RejectProgressInput "Row " & r & ": linked WBS row " & sourceRow & " needs review: " & reason
                taskName = ProgressLabel(taskName)
            End If
            If Len(taskName) = 0 Then RejectProgressInput "Row " & r & ": enter a milestone name."
            If Len(taskName) > 120 Then RejectProgressInput "Row " & r & ": milestone name is too long."
            lastYear = 0
            If r > PG_FIRST Then
                If TryProgressDateWithYear(ws.Cells(r - 1, 5).Value2, 0, previousDate, previousShort) Then lastYear = Year(previousDate)
            End If
            If Not TryProgressDateWithYear(ws.Cells(r, 5).Value2, lastYear, milestoneDate, shortInput) Then RejectProgressInput "Row " & r & ": invalid date. Use yyyy/m/d, yyyymmdd, or mmdd with the previous row year."
            If Not ws.Cells(r, 5).HasFormula Then ws.Cells(r, 5).Value = milestoneDate
            ws.Cells(r, 5).NumberFormat = "yyyy/m/d"
            inputCount = inputCount + 1: If inputCount > PG_LIMIT Then RejectProgressInput "A chart supports up to 60 milestones."
            names(inputCount) = taskName: dates(inputCount) = milestoneDate
        End If
    Next r
    If inputCount = 0 Then RejectProgressInput "Add at least one milestone and set Show to Yes."
    For i = 2 To inputCount
        swapDate = dates(i): swapName = names(i): j = i - 1
        Do While j >= 1
            If dates(j) <= swapDate Then Exit Do
            dates(j + 1) = dates(j): names(j + 1) = names(j): j = j - 1
        Loop
        dates(j + 1) = swapDate: names(j + 1) = swapName
    Next i
    nodeCount = 1
    For i = 2 To inputCount
        If dates(i) = dates(nodeCount) Then
            names(nodeCount) = names(nodeCount) & "; " & names(i)
        Else
            nodeCount = nodeCount + 1
            dates(nodeCount) = dates(i): names(nodeCount) = names(i)
        End If
    Next i
    mode = Trim$(CStr(ws.Range(PG_MODE_CELL).Value)): If mode <> "Date proportional" Then mode = "Equal spacing"
    span = CDbl(dates(nodeCount)) - CDbl(dates(1)): chartScale = ProgressChartScale(ws.Range(PG_SCALE_CELL).Value)
    width = Application.Max(690, nodeCount * 115) * chartScale
    gap = (width - 44) / nodeCount
    left = ws.Range("L3").Left: top = ws.Range("L3").Top
    ReDim centers(1 To nodeCount): ReDim widths(1 To nodeCount): ReDim labelLeft(1 To nodeCount)
    ReDim labelLanes(1 To nodeCount): ReDim dateLanes(1 To nodeCount)
    ReDim laneRight(1 To nodeCount): ReDim dateRight(1 To nodeCount)
    ReDim labelShapes(1 To nodeCount): ReDim dateShapes(1 To nodeCount)
    For i = 1 To nodeCount
        If mode = "Date proportional" Then
            centers(i) = left + width / 2
            If span > 0 Then centers(i) = left + 90 + (width - 180) * (CDbl(dates(i)) - CDbl(dates(1))) / span
        Else
            centers(i) = left + 22 + gap * (i - 0.5)
        End If
    Next i
    ' First pass: width follows adjacent spacing; reserve collision-free lanes.
    For i = 1 To nodeCount
        available = gap
        If i > 1 Then available = Application.Min(available, centers(i) - centers(i - 1))
        If i < nodeCount Then available = Application.Min(available, centers(i + 1) - centers(i))
        widths(i) = Application.Max(64, Application.Min(220, available - 12))
        labelLeft(i) = Application.Max(left + 8, Application.Min(left + width - 8 - widths(i), centers(i) - widths(i) / 2))
        lane = 1
        Do While lane <= labelLaneCount
            If labelLeft(i) >= laneRight(lane) + 10 Then Exit Do
            lane = lane + 1
        Loop
        labelLanes(i) = lane: labelLaneCount = Application.Max(labelLaneCount, lane)
        laneRight(lane) = labelLeft(i) + widths(i)
        lane = 1
        Do While lane <= dateLaneCount
            If centers(i) - 26 >= dateRight(lane) + 10 Then Exit Do
            lane = lane + 1
        Loop
        dateLanes(i) = lane: dateLaneCount = Application.Max(dateLaneCount, lane)
        dateRight(lane) = centers(i) + 26
    Next i
    axis = top + 38 + dateLaneCount * 24
    ' Match the visible top of the arrow shaft, whose thickness is 55% of 12 pt.
    markerTop = axis - 12 * 0.55 / 2 - 8
    Application.ScreenUpdating = False
    ThisWorkbook.Activate: ws.Select Replace:=True: ws.Range("B14").Select: DeleteProgressTemps ws
    ReDim shapeIndexes(0 To 1 + 5 * nodeCount)
    ' Second pass: use Excel's text measurements before setting lane height.
    labelHeight = 30
    For i = 1 To nodeCount
        Set labelShapes(i) = AddProgressText(ws, names(i), labelLeft(i), axis + 18, widths(i), Application.Max(400, 14 * (Len(names(i)) + 1)), True)
        labelHeight = Application.Max(labelHeight, labelShapes(i).TextFrame2.TextRange.BoundHeight + 6)
        If labelShapes(i).TextFrame2.TextRange.BoundWidth > widths(i) + 1 Then
            labelShapes(i).TextFrame2.TextRange.Text = WrapProgressLabel(names(i), MaxLabelChars(widths(i)))
            labelHeight = Application.Max(labelHeight, labelShapes(i).TextFrame2.TextRange.BoundHeight + 6)
        End If
        RememberProgressShape labelShapes(i), shapeIndexes, shapeCount
        Set dateShapes(i) = AddProgressText(ws, Format$(dates(i), "m/d"), centers(i) - 26, markerTop - 15 - (dateLanes(i) - 1) * 24, 52, 13, True)
        RememberProgressShape dateShapes(i), shapeIndexes, shapeCount
    Next i
    height = axis - top + 18 + labelLaneCount * (labelHeight + 10)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    RememberProgressShape shp, shapeIndexes, shapeCount
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Visible = msoFalse
    Set background = shp
    Set shp = ws.Shapes.AddShape(msoShapeRightArrow, left + 8, axis - 6, width - 16, 12)
    RememberProgressShape shp, shapeIndexes, shapeCount
    shp.Adjustments.Item(1) = 0.55: shp.Adjustments.Item(2) = 0.6
    shp.Fill.ForeColor.RGB = RGB(139, 0, 0): shp.Line.Visible = msoFalse
    For i = 1 To nodeCount
        labelShapes(i).Height = labelHeight
        labelShapes(i).Top = axis + 18 + (labelLanes(i) - 1) * (labelHeight + 10)
        If labelLanes(i) > 1 Or Abs(labelLeft(i) + widths(i) / 2 - centers(i)) > 1 Then
            Set shp = ws.Shapes.AddLine(centers(i), axis + 9, labelLeft(i) + widths(i) / 2, labelShapes(i).Top - 3)
            RememberProgressShape shp, shapeIndexes, shapeCount
            shp.Line.ForeColor.RGB = RGB(148, 163, 184): shp.Line.Weight = 0.5
            shp.ZOrder msoSendToBack
        End If
        If dateLanes(i) > 1 Then
            Set shp = ws.Shapes.AddLine(centers(i), markerTop - 1, centers(i), dateShapes(i).Top + 15)
            RememberProgressShape shp, shapeIndexes, shapeCount
            shp.Line.ForeColor.RGB = RGB(148, 163, 184): shp.Line.Weight = 0.5
            shp.ZOrder msoSendToBack
        End If
        Set shp = ws.Shapes.AddShape(msoShapeIsoscelesTriangle, centers(i) - 5, markerTop, 10, 16)
        RememberProgressShape shp, shapeIndexes, shapeCount
        shp.Fill.ForeColor.RGB = RGB(242, 160, 0): shp.Line.Visible = msoFalse
    Next i
    background.ZOrder msoSendToBack
    ReDim Preserve shapeIndexes(0 To shapeCount - 1)
    Set group = ws.Shapes.Range(shapeIndexes).Group: group.Name = PG_TEMP & "GROUP": group.Placement = xlFreeFloating
    group.AlternativeText = "Milestones: sorted by date, " & mode & ", " & inputCount & " milestones in " & nodeCount & " dates."
    ws.Range("B10").Value = "Generated " & inputCount & " milestones; " & Format$(dates(1), "yyyy/m/d") & " - " & Format$(dates(nodeCount), "yyyy/m/d")
    For i = ws.Shapes.Count To 1 Step -1: If ws.Shapes(i).Name = PG_GROUP Then ws.Shapes(i).Delete
    Next i
    group.Name = PG_GROUP
    If ActiveSheet Is ws Then group.Select
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    failureText = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then DeleteProgressTemps ws
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise vbObjectError + 507, "Milestones", failureText
End Sub

Private Sub RejectProgressInput(ByVal message As String)
    Err.Raise vbObjectError + 507, "Milestones", message
End Sub

Private Function MaxLabelChars(ByVal width As Single) As Long
    MaxLabelChars = Int(width / 10): If MaxLabelChars < 4 Then MaxLabelChars = 4
End Function
Private Function ProgressChartScale(ByVal value As Variant) As Double
    If IsNumeric(value) Then ProgressChartScale = CDbl(value) Else ProgressChartScale = 1
    If ProgressChartScale < 0.5 Then ProgressChartScale = 0.5
    If ProgressChartScale > 2 Then ProgressChartScale = 2
End Function
Public Sub CopyProgressChart()
    Dim ws As Worksheet, group As Shape
    On Error GoTo Failed: Set ws = ThisWorkbook.Worksheets(PG_SHEET): Set group = ws.Shapes(PG_GROUP)
    ws.Activate: group.CopyPicture Appearance:=xlScreen, Format:=xlPicture: MsgBox "Chart copied. Paste it into Word, PowerPoint or another sheet.", vbInformation: Exit Sub
Failed: MsgBox "Copy failed: " & Err.Description, vbExclamation
End Sub
Public Sub WidenProgressChart(): AdjustProgressChartScale 0.1: End Sub
Public Sub NarrowProgressChart(): AdjustProgressChartScale -0.1: End Sub
Private Sub AdjustProgressChartScale(ByVal delta As Double)
    Dim ws As Worksheet, value As Variant, oldScale As Double, errorText As String, started As Boolean
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(PG_SHEET): value = ws.Range(PG_SCALE_CELL).Value
    If IsNumeric(value) Then oldScale = CDbl(value) Else oldScale = 1
    Sheet4.BeginEdit: started = True
    ws.Range(PG_SCALE_CELL).Value = Application.Max(0.5, Application.Min(2, oldScale + delta))
    ws.Range(PG_SCALE_CELL).NumberFormat = "0%"
    GenerateProgressChart True
    Sheet4.FinishEdit
    Exit Sub
Failed:
    errorText = "Spacing failed: " & Err.Description
    If started Then errorText = errorText & Sheet4.CancelEdit
    MsgBox errorText, vbExclamation, "Milestones"
End Sub

Private Sub ApplyRefreshAllFormatting()
    Dim milestones As Worksheet
    Set milestones = ThisWorkbook.Worksheets(PG_SHEET)
    With milestones.Range("B3:J10").Font
        .Name = "Microsoft YaHei"
        .Size = 11
    End With
    With milestones.Range("B14:J213").Font
        .Name = "Microsoft YaHei"
        .Size = 11
    End With
End Sub

Private Sub DeleteProgressTemps(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1: If Left$(ws.Shapes(i).Name, Len(PG_TEMP)) = PG_TEMP Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub DeleteProgressDrawings(ByVal ws As Worksheet)
    Dim i As Long, name As String, owned As Boolean, errorText As String
    For i = ws.Shapes.Count To 1 Step -1
        name = ws.Shapes(i).Name
        owned = (Left$(name, Len(PG_GROUP)) = PG_GROUP Or Left$(name, Len(PG_TEMP)) = PG_TEMP)
        owned = owned Or name = "PG_Background" Or name = "PG_Axis"
        owned = owned Or Left$(name, 10) = "PG_Marker_" Or Left$(name, 8) = "PG_Date_" Or Left$(name, 9) = "PG_Label_"
        If Left$(ws.Shapes(i).AlternativeText, 11) = "Milestones:" Then owned = True
        If owned Then
            On Error Resume Next
            Err.Clear
            ws.Shapes(i).Delete
            If Err.Number <> 0 Then errorText = Err.Description
            On Error GoTo 0
        End If
    Next i
    If Len(errorText) > 0 Then Err.Raise vbObjectError + 507, "Milestones cleanup", errorText
End Sub
Private Sub RememberProgressShape(ByVal shp As Shape, ByRef indexes() As Variant, ByRef count As Long)
    shp.Name = PG_TEMP & "SHAPE_" & CStr(shp.ID): shp.Placement = xlFreeFloating: indexes(count) = shp.Name: count = count + 1
End Sub
Private Function AddProgressText(ByVal ws As Worksheet, ByVal value As String, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single, ByVal bold As Boolean) As Shape
    Dim shp As Shape: Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height): shp.Name = PG_TEMP & "SHAPE_" & CStr(shp.ID): shp.Fill.Visible = msoTrue: shp.Fill.Solid: shp.Fill.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Visible = msoFalse
    With shp.TextFrame2: .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0: .WordWrap = msoTrue: .AutoSize = msoAutoSizeNone: .VerticalAnchor = msoAnchorTop: .TextRange.Text = value: .TextRange.ParagraphFormat.Alignment = msoAlignCenter: .TextRange.Font.Name = "Arial": .TextRange.Font.Size = 10: .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0): If bold Then .TextRange.Font.Bold = msoTrue
    End With: Set AddProgressText = shp
End Function
Public Function ProgressLabel(ByVal value As String) As String
    Dim p As Long, token As String, numbered As Boolean
    value = Trim$(Replace(value, vbTab, " ")): p = InStr(value, " ")
    If p > 1 Then token = Left$(value, p - 1): numbered = IsTaskNumber(token): If numbered Then value = Trim$(Mid$(value, p + 1))
    ProgressLabel = value
End Function
Private Function IsTaskNumber(ByVal value As String) As Boolean
    Dim i As Long, ch As String: If Len(value) = 0 Then Exit Function
    For i = 1 To Len(value): ch = Mid$(value, i, 1): If (ch < "0" Or ch > "9") And ch <> "." Then Exit Function
    Next i: IsTaskNumber = True
End Function
Public Function WrapProgressLabel(ByVal value As String, ByVal maxChars As Long) As String
    Dim i As Long, n As Long, ch As String, result As String
    value = Replace(Replace(value, vbCrLf, vbLf), vbCr, vbLf)
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch = vbLf Then
            result = result & vbLf
            n = 0
        Else
            If n >= maxChars Then
                result = result & vbLf
                n = 0
            End If
            result = result & ch
            n = n + 1
        End If
    Next i
    WrapProgressLabel = result
End Function
Public Function TryProgressDate(ByVal value As Variant, ByRef result As Date) As Boolean
    Dim ignored As Boolean: TryProgressDate = TryProgressDateWithYear(value, 0, result, ignored)
End Function
Public Function TryProgressDateWithYear(ByVal value As Variant, ByVal fallbackYear As Long, ByRef result As Date, ByRef shortInput As Boolean) As Boolean
    Dim text As String, parts() As String, y As Long, m As Long, d As Long, serial As Double, digits As String
    On Error GoTo InvalidDate
    shortInput = IsShortProgressDate(value)
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or VarType(value) = vbBoolean Then Exit Function
    If VarType(value) = vbDate Then
        result = DateSerial(Year(value), Month(value), Day(value))
        GoTo AcceptDate
    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
        serial = CDbl(value)
        If serial <> Fix(serial) Then Exit Function
        If Fix(serial) >= 10000000# And Fix(serial) <= 99999999# Then
            digits = Format$(Fix(serial), "0")
            y = CLng(Left$(digits, 4)): m = CLng(Mid$(digits, 5, 2)): d = CLng(Right$(digits, 2))
            GoTo BuildDate
        ElseIf shortInput Then
            If fallbackYear = 0 Then Exit Function
            digits = Right$("0000" & Format$(Fix(serial), "0"), 4)
            y = fallbackYear: m = CLng(Left$(digits, 2)): d = CLng(Right$(digits, 2))
            GoTo BuildDate
        Else
            If serial < 1 Or serial >= 2958466# Then Exit Function
            If Fix(serial) = 60 Then Exit Function
            If serial < 60 Then serial = serial + 1
            result = CDate(Fix(serial))
            GoTo AcceptDate
        End If
    Else
        text = Trim$(CStr(value))
        If IsDigitsOnly(text) Then
            If Len(text) = 8 Then
                y = CLng(Left$(text, 4)): m = CLng(Mid$(text, 5, 2)): d = CLng(Right$(text, 2))
                GoTo BuildDate
            ElseIf Len(text) >= 3 And Len(text) <= 4 Then
                If fallbackYear = 0 Then Exit Function
                digits = Right$("0000" & text, 4)
                y = fallbackYear: m = CLng(Left$(digits, 2)): d = CLng(Right$(digits, 2))
                GoTo BuildDate
            Else
                Exit Function
            End If
        End If
        text = Replace(Replace(text, "-", "/"), ".", "/")
        parts = Split(text, "/")
        If UBound(parts) <> 2 Then Exit Function
        If Len(parts(0)) <> 4 Or Len(parts(1)) < 1 Or Len(parts(1)) > 2 Or Len(parts(2)) < 1 Or Len(parts(2)) > 2 Then Exit Function
        If Not IsDigitsOnly(parts(0)) Or Not IsDigitsOnly(parts(1)) Or Not IsDigitsOnly(parts(2)) Then Exit Function
        y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    End If
BuildDate:
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    result = DateSerial(y, m, d)
    If Year(result) <> y Or Month(result) <> m Or Day(result) <> d Then Exit Function
AcceptDate:
    If Year(result) < 1900 Then Exit Function
    TryProgressDateWithYear = True
InvalidDate:
End Function
Private Function IsShortProgressDate(ByVal value As Variant) As Boolean
    Dim text As String, n As Double: On Error GoTo Done: If IsError(value) Or IsEmpty(value) Or IsNull(value) Or VarType(value) = vbBoolean Then Exit Function
    If VarType(value) = vbDate Then Exit Function
    If VarType(value) <> vbString And IsNumeric(value) Then n = CDbl(value): IsShortProgressDate = n = Fix(n) And n >= 100 And n <= 9999 Else text = Trim$(CStr(value)): IsShortProgressDate = IsDigitsOnly(text) And Len(text) >= 3 And Len(text) <= 4
Done:
End Function
Private Function IsDigitsOnly(ByVal value As String) As Boolean
    Dim i As Long, ch As String: If Len(value) = 0 Then Exit Function
    For i = 1 To Len(value): ch = Mid$(value, i, 1): If ch < "0" Or ch > "9" Then Exit Function
    Next i: IsDigitsOnly = True
End Function
Private Function ValidateImportedTask(ByVal rowNumber As Long, ByRef reason As String) As Boolean
    Dim source As Worksheet, value As Variant, startDate As Date, endDate As Date, progress As Double
    reason = ""
    On Error GoTo InvalidTask
    Set source = ThisWorkbook.Worksheets("WBS")
    If rowNumber < 6 Or rowNumber > source.Rows.Count Then reason = "Invalid WBS row.": Exit Function
    value = source.Cells(rowNumber, 3).Value2
    If IsError(value) Then reason = "Task name has an error.": Exit Function
    If Len(Trim$(CStr(value))) = 0 Then reason = "The source task no longer exists.": Exit Function
    value = source.Cells(rowNumber, 7).Value
    If Not TryProgressDate(value, startDate) Then reason = "Invalid planned start date.": Exit Function
    value = source.Cells(rowNumber, 8).Value
    If Not TryProgressDate(value, endDate) Then reason = "Invalid planned end date.": Exit Function
    If startDate > endDate Then reason = "Planned start is after planned end.": Exit Function
    value = source.Cells(rowNumber, 14).Value2
    If IsError(value) Then reason = "Progress has an error.": Exit Function
    If Len(Trim$(CStr(value))) > 0 Then
        If VarType(value) = vbBoolean Or Not IsNumeric(value) Then reason = "Progress is not numeric.": Exit Function
        progress = CDbl(value)
        If progress < 0 Or progress > 1 Then reason = "Progress must be 0% to 100%.": Exit Function
    End If
    ValidateImportedTask = True: Exit Function
InvalidTask: reason = "Unable to read task data: " & Err.Description
End Function
