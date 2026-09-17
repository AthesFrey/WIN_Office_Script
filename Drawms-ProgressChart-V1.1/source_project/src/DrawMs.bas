Attribute VB_Name = "DrawMs"
Option Explicit

Private Const DM_SHEET As String = "Milestones"
Private Const DM_FIRST As Long = 14
Private Const DM_LAST As Long = 213
Private Const DM_LIMIT As Long = 60
Private Const DM_GROUP As String = "DrawMs_Chart"
Private Const DM_TEMP As String = "DRAWMS_TMP_"
Private Const DM_CHART_LEFT As Single = 650

Public Sub GenerateChart()
    Dim errorText As String, spacingScale As Double
    On Error GoTo Failed
    spacingScale = ReadChartScale(ThisWorkbook.Worksheets(DM_SHEET))
    If TryRefreshChart(errorText, spacingScale) Then Exit Sub
Failed:
    If Len(errorText) = 0 Then errorText = Err.Description
    MsgBox "Chart generation failed: " & errorText, vbExclamation, "drawms"
End Sub

Public Sub WidenSpacing()
    AdjustSpacing 0.1
End Sub

Public Sub NarrowSpacing()
    AdjustSpacing -0.1
End Sub

Private Sub AdjustSpacing(ByVal delta As Double)
    Dim ws As Worksheet, spacingScale As Double, errorText As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(DM_SHEET)
    spacingScale = ReadChartScale(ws) + delta
    spacingScale = ClampScale(spacingScale)
    If TryRefreshChart(errorText, spacingScale) Then Exit Sub
Failed:
    If Len(errorText) = 0 Then errorText = Err.Description
    MsgBox "Spacing adjustment failed: " & errorText, vbExclamation, "drawms"
End Sub

Private Function TryRefreshChart(ByRef errorText As String, ByVal spacingScale As Double) As Boolean
    If TryBuildChart(errorText, spacingScale) Then
        TryRefreshChart = True
        Exit Function
    End If
    On Error Resume Next
    DeleteChart ThisWorkbook.Worksheets(DM_SHEET)
    On Error GoTo 0
    If TryBuildChart(errorText, spacingScale) Then TryRefreshChart = True
End Function

Private Function TryBuildChart(ByRef errorText As String, ByVal spacingScale As Double) As Boolean
    On Error GoTo Failed
    BuildChart spacingScale
    TryBuildChart = True
    Exit Function
Failed:
    errorText = Err.Description
End Function

Private Sub BuildChart(ByVal chartScale As Double)
    Dim ws As Worksheet, names() As String, dates() As String
    Dim inputCount As Long, nodeCount As Long, r As Long, i As Long
    Dim valueName As String, valueDate As String
    Dim width As Single, height As Single, gap As Single, left As Single, top As Single, axis As Single
    Dim markerTop As Single, labelHeight As Single, available As Single
    Dim centers() As Single, widths() As Single, dateWidths() As Single, dateHeights() As Single, labelLeft() As Single
    Dim labelLanes() As Long, dateLanes() As Long, laneRight() As Single, dateRight() As Single
    Dim labelShapes() As Shape, dateShapes() As Shape
    Dim dateLaneCount As Long, labelLaneCount As Long, lane As Long
    Dim dateLaneGap As Single, maxDateHeight As Single, dateLines As Long
    Dim shapeIndexes() As Variant, shapeCount As Long, shp As Shape, group As Shape
    Dim background As Shape, oldScreen As Boolean
    Dim failureText As String
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    Set ws = ThisWorkbook.Worksheets(DM_SHEET)
    chartScale = ClampScale(chartScale)
    ReDim names(1 To DM_LIMIT): ReDim dates(1 To DM_LIMIT)
    For r = DM_FIRST To DM_LAST
        valueName = CellText(ws.Cells(r, 2))
        valueDate = CellText(ws.Cells(r, 3))
        If Len(valueName) > 0 Or Len(valueDate) > 0 Then
            inputCount = inputCount + 1
            If inputCount > DM_LIMIT Then RejectInput "A chart supports up to 60 milestones."
            names(inputCount) = valueName
            dates(inputCount) = valueDate
        End If
    Next r
    If inputCount = 0 Then RejectInput "Enter at least one milestone name or date."
    nodeCount = inputCount
    width = Application.Max(690, nodeCount * 115) * chartScale
    gap = (width - 44) / nodeCount
    ' Keep the chart clear of the three fixed controls even when a user
    ' changes the input column widths.
    left = DM_CHART_LEFT
    top = ws.Range("F3").Top
    ReDim centers(1 To nodeCount): ReDim widths(1 To nodeCount)
    ReDim dateWidths(1 To nodeCount): ReDim dateHeights(1 To nodeCount): ReDim labelLeft(1 To nodeCount)
    ReDim labelLanes(1 To nodeCount): ReDim dateLanes(1 To nodeCount)
    ReDim laneRight(1 To nodeCount): ReDim dateRight(1 To nodeCount)
    ReDim labelShapes(1 To nodeCount): ReDim dateShapes(1 To nodeCount)
    ReDim shapeIndexes(0 To 1 + 5 * nodeCount)
    For i = 1 To nodeCount
        centers(i) = left + 22 + gap * (i - 0.5)
        available = gap
        If i > 1 Then available = Application.Min(available, centers(i) - centers(i - 1))
        If i < nodeCount Then available = Application.Min(available, centers(i + 1) - centers(i))
        widths(i) = Application.Max(64, Application.Min(220, available - 12))
        dateWidths(i) = Application.Max(52, Application.Min(180, 8 * Len(dates(i)) + 12))
        dateLines = (Len(dates(i)) + MaxTextChars(dateWidths(i)) - 1) \ MaxTextChars(dateWidths(i))
        If dateLines < 1 Then dateLines = 1
        dateHeights(i) = Application.Max(20, 15 * dateLines)
        maxDateHeight = Application.Max(maxDateHeight, dateHeights(i))
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
            If centers(i) - dateWidths(i) / 2 >= dateRight(lane) + 10 Then Exit Do
            lane = lane + 1
        Loop
        dateLanes(i) = lane: dateLaneCount = Application.Max(dateLaneCount, lane)
        dateRight(lane) = centers(i) + dateWidths(i) / 2
    Next i
    dateLaneGap = Application.Max(24, maxDateHeight + 8)
    axis = top + 38 + dateLaneCount * dateLaneGap
    markerTop = axis - 12 * 0.55 / 2 - 8
    Application.ScreenUpdating = False
    ThisWorkbook.Activate: ws.Select Replace:=True: ws.Range("B14").Select
    DeleteTemporaryShapes ws
    labelHeight = 30
    For i = 1 To nodeCount
        Set labelShapes(i) = AddDrawMsText(ws, names(i), labelLeft(i), axis + 18, widths(i), _
                                           Application.Max(400, 14 * (Len(names(i)) + 1)), True)
        If labelShapes(i).TextFrame2.TextRange.BoundWidth > widths(i) + 1 Then
            labelShapes(i).TextFrame2.TextRange.Text = WrapDrawMsText(names(i), MaxTextChars(widths(i)))
        End If
        labelHeight = Application.Max(labelHeight, labelShapes(i).TextFrame2.TextRange.BoundHeight + 6)
        RememberShape labelShapes(i), shapeIndexes, shapeCount
        Set dateShapes(i) = AddDrawMsText(ws, dates(i), centers(i) - dateWidths(i) / 2, _
                                          markerTop - 2 - dateHeights(i) - (dateLanes(i) - 1) * dateLaneGap, _
                                          dateWidths(i), dateHeights(i), True)
        If dateShapes(i).TextFrame2.TextRange.BoundWidth > dateWidths(i) + 1 Then
            dateShapes(i).TextFrame2.TextRange.Text = WrapDrawMsText(dates(i), MaxTextChars(dateWidths(i)))
        End If
        RememberShape dateShapes(i), shapeIndexes, shapeCount
    Next i
    height = axis - top + 18 + labelLaneCount * (labelHeight + 10)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    RememberShape shp, shapeIndexes, shapeCount
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Visible = msoFalse
    Set background = shp
    Set shp = ws.Shapes.AddShape(msoShapeRightArrow, left + 8, axis - 6, width - 16, 12)
    RememberShape shp, shapeIndexes, shapeCount
    shp.Adjustments.Item(1) = 0.55: shp.Adjustments.Item(2) = 0.6
    shp.Fill.ForeColor.RGB = RGB(139, 0, 0): shp.Line.Visible = msoFalse
    For i = 1 To nodeCount
        labelShapes(i).Height = labelHeight
        labelShapes(i).Top = axis + 18 + (labelLanes(i) - 1) * (labelHeight + 10)
        If labelLanes(i) > 1 Or Abs(labelLeft(i) + widths(i) / 2 - centers(i)) > 1 Then
            Set shp = ws.Shapes.AddLine(centers(i), axis + 9, labelLeft(i) + widths(i) / 2, labelShapes(i).Top - 3)
            RememberShape shp, shapeIndexes, shapeCount
            shp.Line.ForeColor.RGB = RGB(148, 163, 184): shp.Line.Weight = 0.5
            shp.ZOrder msoSendToBack
        End If
        If dateLanes(i) > 1 Then
            Set shp = ws.Shapes.AddLine(centers(i), markerTop - 1, centers(i), dateShapes(i).Top + dateShapes(i).Height)
            RememberShape shp, shapeIndexes, shapeCount
            shp.Line.ForeColor.RGB = RGB(148, 163, 184): shp.Line.Weight = 0.5
            shp.ZOrder msoSendToBack
        End If
        Set shp = ws.Shapes.AddShape(msoShapeIsoscelesTriangle, centers(i) - 5, markerTop, 10, 16)
        RememberShape shp, shapeIndexes, shapeCount
        shp.Fill.ForeColor.RGB = RGB(242, 160, 0): shp.Line.Visible = msoFalse
    Next i
    background.ZOrder msoSendToBack
    ReDim Preserve shapeIndexes(0 To shapeCount - 1)
    Set group = ws.Shapes.Range(shapeIndexes).Group
    group.Name = DM_TEMP & "GROUP": group.Placement = xlFreeFloating
    group.AlternativeText = "drawms: scale=" & Format$(chartScale, "0.0") & "; nodes=" & CStr(inputCount)
    ws.Range("B10").Value = "Generated " & inputCount & " milestones."
    DeleteChart ws
    group.Name = DM_GROUP
    If ActiveSheet Is ws Then group.Select
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    failureText = Err.Description
    On Error Resume Next
    DeleteTemporaryShapes ws
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise vbObjectError + 507, "drawms", failureText
End Sub

Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value2) Or IsEmpty(cell.Value2) Or IsNull(cell.Value2) Then
        CellText = ""
    Else
        CellText = CStr(cell.Value2)
    End If
End Function

Private Sub RejectInput(ByVal message As String)
    Err.Raise vbObjectError + 507, "drawms", message
End Sub

Private Function ClampScale(ByVal value As Double) As Double
    ClampScale = value
    If ClampScale < 0.5 Then ClampScale = 0.5
    If ClampScale > 2 Then ClampScale = 2
End Function

Private Function ReadChartScale(ByVal ws As Worksheet) As Double
    Dim group As Shape, text As String, start As Long, finish As Long, value As String
    ReadChartScale = 1
    On Error GoTo Done
    Set group = ws.Shapes(DM_GROUP)
    text = group.AlternativeText
    start = InStr(1, text, "scale=", vbTextCompare)
    If start = 0 Then GoTo Done
    value = Mid$(text, start + 6)
    finish = InStr(1, value, ";", vbBinaryCompare)
    If finish > 0 Then value = Left$(value, finish - 1)
    If IsNumeric(value) Then ReadChartScale = CDbl(value)
Done:
    ReadChartScale = ClampScale(ReadChartScale)
End Function

Private Sub DeleteTemporaryShapes(ByVal ws As Worksheet)
    Dim i As Long
    If ws Is Nothing Then Exit Sub
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(DM_TEMP)) = DM_TEMP Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub DeleteChart(ByVal ws As Worksheet)
    Dim i As Long, name As String
    If ws Is Nothing Then Exit Sub
    For i = ws.Shapes.Count To 1 Step -1
        name = ws.Shapes(i).Name
        If name = DM_GROUP Then ws.Shapes(i).Delete
        If Left$(name, 8) = "DM_Date_" Or Left$(name, 9) = "DM_Label_" Then ws.Shapes(i).Delete
        If Left$(name, 10) = "DM_Marker_" Or name = "DM_Background" Or name = "DM_Axis" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub RememberShape(ByVal shp As Shape, ByRef indexes() As Variant, ByRef count As Long)
    shp.Name = DM_TEMP & "SHAPE_" & CStr(shp.ID)
    shp.Placement = xlFreeFloating
    indexes(count) = shp.Name: count = count + 1
End Sub

Private Function AddDrawMsText(ByVal ws As Worksheet, ByVal value As String, ByVal left As Single, _
                               ByVal top As Single, ByVal width As Single, ByVal height As Single, _
                               ByVal bold As Boolean) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
    shp.Name = DM_TEMP & "SHAPE_" & CStr(shp.ID)
    shp.Fill.Visible = msoTrue: shp.Fill.Solid: shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoTrue: .AutoSize = msoAutoSizeNone: .VerticalAnchor = msoAnchorTop
        .TextRange.Text = value: .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Name = "Arial": .TextRange.Font.Size = 10
        .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        If bold Then .TextRange.Font.Bold = msoTrue
    End With
    Set AddDrawMsText = shp
End Function

Private Function MaxTextChars(ByVal width As Single) As Long
    MaxTextChars = Int(width / 10)
    If MaxTextChars < 4 Then MaxTextChars = 4
End Function

Private Function WrapDrawMsText(ByVal value As String, ByVal maxChars As Long) As String
    Dim i As Long, count As Long, ch As String, result As String
    value = Replace(Replace(value, vbCrLf, vbLf), vbCr, vbLf)
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch = vbLf Then
            result = result & vbLf: count = 0
        Else
            If count >= maxChars Then result = result & vbLf: count = 0
            result = result & ch: count = count + 1
        End If
    Next i
    WrapDrawMsText = result
End Function
