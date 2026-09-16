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

Public Sub RefreshAll()
    Dim oldScreen As Boolean
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets("WBS").Activate
    ThisWorkbook.Worksheets("WBS").Range("C6").Select
    Model_1.UpdateGanttChart True
    GenerateProgressChart
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    Application.ScreenUpdating = oldScreen
    MsgBox "Refresh All failed (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub

Public Sub UpdateGantt(): Model_1.UpdateGanttChart: End Sub
Public Sub AddTask(): Model_1.AddTask: End Sub
Public Sub AddSubtask(): Model_1.AddSubtask: End Sub
Public Sub PromoteTasks(): Sheet1.ChangeSelectedLevels -1: End Sub
Public Sub DemoteTasks(): Sheet1.ChangeSelectedLevels 1: End Sub

Public Sub ImportSelectedMilestones()
    Dim source As Worksheet, target As Worksheet, picked As Range, part As Range
    Dim rows() As Long, count As Long, r As Long, i As Long, j As Long, swapRow As Long
    Dim lastRow As Long, dest As Long, freeRows As Long, duplicate As Boolean, taskValue As Variant
    On Error GoTo Failed
    Set source = ThisWorkbook.Worksheets("WBS")
    Set target = ThisWorkbook.Worksheets(PG_SHEET)
    If Not ActiveSheet Is source Or TypeName(Selection) <> "Range" Then
        MsgBox "Select task cells or rows on WBS first.", vbInformation
        Exit Sub
    End If
    lastRow = source.Cells(source.Rows.Count, 3).End(xlUp).Row
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
                    If IsNumeric(target.Cells(i, 3).Value) Then If CLng(target.Cells(i, 3).Value) = r Then duplicate = True
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
        target.Cells(dest, 8).NumberFormat = "0%"
        target.Cells(dest, 9).Formula = "=IF(WBS!C" & r & "="""","""",IFERROR(IF(WBS!N" & r & ">=1,""Complete"",IF(WBS!H" & r & "<TODAY(),""Overdue"",IF(WBS!G" & r & ">TODAY(),""Not started"",""In progress""))),""Check""))"
        target.Cells(dest, 10).Formula = "=IFERROR(WBS!X" & r & ","""")"
        dest = dest + 1
    Next i
    target.Activate
    target.Range("B10").Value = "Imported " & count & ". Click Generate."
    Exit Sub
Failed:
    MsgBox "Task import failed: " & Err.Description, vbExclamation
End Sub

Public Sub GenerateProgressChart()
    Dim ws As Worksheet, names() As String, dates() As Date, count As Long
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
    ThisWorkbook.Worksheets("WBS").Calculate: ws.Calculate
    ReDim names(1 To PG_LIMIT): ReDim dates(1 To PG_LIMIT)
    For r = PG_FIRST To PG_LAST
        mark = Trim$(CStr(ws.Cells(r, 2).Value2))
        If mark <> "" And mark <> "Yes" And mark <> "No" Then MsgBox "Row " & r & ": Show must be Yes or No.", vbExclamation: Exit Sub
        If mark = "Yes" Then
            value = ws.Cells(r, 4).Value
            If IsError(value) Then MsgBox "Row " & r & ": milestone name has an error.", vbExclamation: Exit Sub
            taskName = Trim$(CStr(value))
            If Len(CStr(ws.Cells(r, 3).Value2)) > 0 And IsNumeric(ws.Cells(r, 3).Value2) Then taskName = ProgressLabel(taskName)
            If Len(taskName) = 0 Then MsgBox "Row " & r & ": enter a milestone name.", vbExclamation: Exit Sub
            If Len(taskName) > 120 Then MsgBox "Row " & r & ": milestone name is too long.", vbExclamation: Exit Sub
            lastYear = 0
            If r > PG_FIRST Then
                If TryProgressDateWithYear(ws.Cells(r - 1, 5).Value2, 0, previousDate, previousShort) Then lastYear = Year(previousDate)
            End If
            If Not TryProgressDateWithYear(ws.Cells(r, 5).Value2, lastYear, milestoneDate, shortInput) Then MsgBox "Row " & r & ": invalid date. Use yyyy/m/d, yyyymmdd, or mmdd with the previous row year.", vbExclamation: Exit Sub
            If Not ws.Cells(r, 5).HasFormula Then ws.Cells(r, 5).Value = milestoneDate
            ws.Cells(r, 5).NumberFormat = "yyyy/m/d"
            If Len(CStr(ws.Cells(r, 3).Value2)) > 0 And IsNumeric(ws.Cells(r, 3).Value2) Then
                sourceRow = CLng(ws.Cells(r, 3).Value)
                If Not ValidateImportedTask(sourceRow, reason) Then MsgBox "Row " & r & ": linked WBS row " & sourceRow & " needs review: " & reason, vbExclamation: Exit Sub
            End If
            count = count + 1: If count > PG_LIMIT Then MsgBox "A chart supports up to 60 milestones.", vbExclamation: Exit Sub
            names(count) = taskName: dates(count) = milestoneDate
        End If
    Next r
    If count = 0 Then MsgBox "Add at least one milestone and set Show to Yes.", vbInformation: Exit Sub
    For i = 2 To count
        swapDate = dates(i): swapName = names(i): j = i - 1
        Do While j >= 1
            If dates(j) <= swapDate Then Exit Do
            dates(j + 1) = dates(j): names(j + 1) = names(j): j = j - 1
        Loop
        dates(j + 1) = swapDate: names(j + 1) = swapName
    Next i
    mode = Trim$(CStr(ws.Range(PG_MODE_CELL).Value)): If mode <> "Date proportional" Then mode = "Equal spacing"
    span = CDbl(dates(count)) - CDbl(dates(1)): chartScale = ProgressChartScale(ws.Range(PG_SCALE_CELL).Value)
    width = Application.Max(690, count * 115) * chartScale
    gap = (width - 44) / count
    left = ws.Range("L3").Left: top = ws.Range("L3").Top
    ReDim centers(1 To count): ReDim widths(1 To count): ReDim labelLeft(1 To count)
    ReDim labelLanes(1 To count): ReDim dateLanes(1 To count)
    ReDim laneRight(1 To count): ReDim dateRight(1 To count)
    ReDim labelShapes(1 To count): ReDim dateShapes(1 To count)
    For i = 1 To count
        If mode = "Date proportional" Then
            centers(i) = left + width / 2
            If span > 0 Then centers(i) = left + 90 + (width - 180) * (CDbl(dates(i)) - CDbl(dates(1))) / span
        Else
            centers(i) = left + 22 + gap * (i - 0.5)
        End If
    Next i
    ' First pass: width follows adjacent spacing; reserve collision-free lanes.
    For i = 1 To count
        available = gap
        If i > 1 Then available = Application.Min(available, centers(i) - centers(i - 1))
        If i < count Then available = Application.Min(available, centers(i + 1) - centers(i))
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
    ReDim shapeIndexes(0 To 1 + 5 * count)
    ' Second pass: use Excel's text measurements before setting lane height.
    labelHeight = 30
    For i = 1 To count
        Set labelShapes(i) = AddProgressText(ws, names(i), labelLeft(i), axis + 18, widths(i), 400, True)
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
    For i = 1 To count
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
    group.AlternativeText = "Milestones: sorted by date, " & mode & ", " & count & " milestones."
    ws.Range("B10").Value = "Generated " & count & " milestones; " & Format$(dates(1), "yyyy/m/d") & " - " & Format$(dates(count), "yyyy/m/d")
    For i = ws.Shapes.Count To 1 Step -1: If ws.Shapes(i).Name = PG_GROUP Then ws.Shapes(i).Delete
    Next i
    group.Name = PG_GROUP: group.Select: Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    failureText = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then DeleteProgressTemps ws
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    MsgBox "Milestones chart generation failed: " & failureText, vbExclamation
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
    Dim ws As Worksheet, value As Variant, oldScale As Double
    Set ws = ThisWorkbook.Worksheets(PG_SHEET): value = ws.Range(PG_SCALE_CELL).Value
    If IsNumeric(value) Then oldScale = CDbl(value) Else oldScale = 1
    ws.Range(PG_SCALE_CELL).Value = Application.Max(0.5, Application.Min(2, oldScale + delta)): ws.Range(PG_SCALE_CELL).NumberFormat = "0%": GenerateProgressChart
End Sub
Private Sub DeleteProgressTemps(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1: If Left$(ws.Shapes(i).Name, Len(PG_TEMP)) = PG_TEMP Then ws.Shapes(i).Delete
    Next i
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
        If UBound(parts) <> 2 Or Len(parts(0)) <> 4 Then Exit Function
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
    On Error GoTo InvalidTask: Set source = ThisWorkbook.Worksheets("WBS"): If rowNumber < 6 Then reason = "Invalid WBS row.": Exit Function
    value = source.Cells(rowNumber, 7).Value2: If IsError(value) Or Not TryProgressDate(value, startDate) Then reason = "Invalid planned start date.": Exit Function
    value = source.Cells(rowNumber, 8).Value2: If IsError(value) Or Not TryProgressDate(value, endDate) Then reason = "Invalid planned end date.": Exit Function
    If startDate > endDate Then reason = "Planned start is after planned end.": Exit Function
    value = source.Cells(rowNumber, 14).Value2: If Not IsError(value) And Len(Trim$(CStr(value))) > 0 Then If Not IsNumeric(value) Then reason = "Progress is not numeric.": Exit Function Else progress = CDbl(value): If progress < 0 Or progress > 1 Then reason = "Progress must be 0% to 100%.": Exit Function
    ValidateImportedTask = True: Exit Function
InvalidTask: reason = "Unable to read task data: " & Err.Description
End Function
