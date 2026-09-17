Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private mSelectedTasks As Range
Private mInitializing As Boolean

Public Sub InitializeEditor(Optional ByVal resetHistory As Boolean = False)
    Dim errorNumber As Long, errorText As String
    If mInitializing Then Exit Sub
    Model_1.BeginHistoryMaintenance
    mInitializing = True
    On Error GoTo Failed
    If ActiveSheet Is Me Then
        If TypeName(Selection) = "Range" Then Set mSelectedTasks = SelectedTaskCells(Selection)
    End If
    If resetHistory Then Model_1.ClearHistory
    Model_1.CaptureUndoSnapshot
Done:
    On Error Resume Next
    Model_1.UpdateHistoryUI
    mInitializing = False
    Model_1.EndHistoryMaintenance
    On Error GoTo 0
    If errorNumber <> 0 Then Model_1.ReportHistoryFailure "WBS editor", errorNumber, errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Model_1.RememberHistoryFailure "WBS editor", errorNumber, errorText
    Model_1.ClearHistory
    Resume Done
End Sub

Private Sub Worksheet_Activate()
    If Model_1.HistoryMaintenanceActive Then Exit Sub
    InitializeEditor
End Sub

Private Sub Worksheet_Deactivate()
    If Model_1.HistoryMaintenanceActive Then Exit Sub
    Model_1.ReleaseHistoryKeys
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim part As Range, throughRow As Long, errorNumber As Long, errorText As String
    If Model_1.HistoryMaintenanceActive Or mInitializing Then Exit Sub
    Model_1.BeginHistoryMaintenance
    On Error GoTo Failed
    If Target.Address = "$D$2" Then GoTo Done
    Set mSelectedTasks = SelectedTaskCells(Target)
    ShowSelectedLevel
    For Each part In Target.Areas
        If part.Row >= 6 And part.Rows.Count < Me.Rows.Count - 5 Then throughRow = Application.Max(throughRow, part.Row + part.Rows.Count - 1)
    Next part
    Model_1.CaptureUndoSnapshot throughRow
Done:
    On Error Resume Next
    Model_1.UpdateHistoryUI
    Model_1.EndHistoryMaintenance
    On Error GoTo 0
    If errorNumber <> 0 Then Model_1.ReportHistoryFailure "WBS selection", errorNumber, errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Model_1.RememberHistoryFailure "WBS selection", errorNumber, errorText
    Model_1.ClearHistory
    Resume Done
End Sub

Private Sub ShowSelectedLevel()
    Dim cell As Range, level As Long, firstLevel As Long
    If mSelectedTasks Is Nothing Then Me.Range("D2").ClearContents: Exit Sub
    For Each cell In mSelectedTasks.Cells
        If Len(Trim$(CStr(cell.Value2))) > 0 Then
            level = DetectLevel(CStr(cell.Value2))
            If firstLevel = 0 Then firstLevel = level
            If level <> firstLevel Then Me.Range("D2").ClearContents: Exit Sub
        End If
    Next cell
    If firstLevel > 0 Then Me.Range("D2").Value = firstLevel Else Me.Range("D2").ClearContents
End Sub

Public Sub ChangeSelectedLevels(ByVal delta As Long)
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation, warning As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    If Not ActiveSheet Is Me Then Exit Sub
    If TypeName(Selection) = "Range" Then
        If Selection.Address <> "$D$2" Then Set mSelectedTasks = SelectedTaskCells(Selection)
    End If
    Model_1.CaptureUndoSnapshot
    Model_1.BeginWbsEdit 6
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    SetTaskLevels mSelectedTasks, 0, delta
Done:
    On Error Resume Next
    Application.EnableEvents = False
    ShowSelectedLevel
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If Len(warning) > 0 Then MsgBox warning, vbExclamation, "Task level"
    Model_1.FinishWbsEdit
    Exit Sub
Failed:
    warning = Err.Description: Resume Done
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watched As Range, dateCells As Range, progressCells As Range, cell As Range, rowsChanged As Object, key As Variant, part As Range
    Dim parsed As Date, startDate As Date, endDate As Date, oldEvents As Boolean, r As Long, value As Variant
    Dim warning As String, days As Double, lastRow As Long, rebuild As Boolean, relink As Boolean, levelChange As Boolean
    Dim oldScreen As Boolean, oldCalculation As XlCalculation, undoable As Boolean, editLastRow As Long
    If Model_1.HistoryMaintenanceActive Then Exit Sub
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    lastRow = Application.Max(6, Me.Cells(Me.Rows.Count, 3).End(xlUp).Row)
    rebuild = Not Intersect(Target, Me.Range("C6:C" & Me.Rows.Count)) Is Nothing
    Set watched = Intersect(Target, Union(Me.Range("C6:C" & lastRow), Me.Range("F6:H" & lastRow), Me.Range("N6:O" & lastRow), Me.Range("D2")))
    undoable = True
    For Each part In Target.Areas
        If part.Row < 6 Or part.Column + part.Columns.Count - 1 > 25 Then undoable = False
        editLastRow = Application.Max(editLastRow, part.Row + part.Rows.Count - 1)
    Next part
    If Target.Address = "$D$2" Then undoable = True
    If undoable Then Model_1.BeginWbsEdit editLastRow, Target
    If watched Is Nothing And Not rebuild Then
        If undoable Then Model_1.FinishWbsEdit Else Model_1.ResetUndoTracking
        Exit Sub
    End If
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    levelChange = Not Intersect(Target, Me.Range("D2")) Is Nothing
    Set progressCells = Intersect(Target, Me.Range("N6:N" & lastRow))
    If Not progressCells Is Nothing Then progressCells.NumberFormat = Model_1.PROGRESS_FORMAT
    If levelChange Then
        value = Me.Range("D2").Value2
        If Not IsNumeric(value) Then warning = "Choose a whole-number level from 1 to 5.": GoTo Done
        days = CDbl(value)
        If days <> Fix(days) Or days < 1 Or days > g_max_level Then warning = "Choose a whole-number level from 1 to 5.": GoTo Done
        SetTaskLevels mSelectedTasks, CLng(days)
        GoTo Done
    End If
    If watched Is Nothing Then RefreshParentTasks: GoTo Done
    relink = Not Intersect(watched, Me.Range("O6:O" & lastRow)) Is Nothing
    Set rowsChanged = CreateObject("Scripting.Dictionary")
    Set dateCells = Intersect(watched, Me.Range("G6:H" & lastRow))
    If Not dateCells Is Nothing Then
        For Each cell In dateCells.Cells
            If Not IsParentTask(cell.Row) And Len(CStr(cell.Value2)) > 0 Then
                If TryNormalizeWbsDate(cell, parsed) Then
                    If Not cell.HasFormula Then cell.Value = parsed
                    cell.NumberFormat = "yyyy/m/d"
                Else
                    value = cell.Value2
                    cell.NumberFormat = "@": cell.Value = CStr(value)
                    warning = "Invalid date in " & cell.Address(False, False) & ". Use yyyy/m/d or yyyymmdd. Short dates need a valid date immediately above."
                    GoTo Done
                End If
            End If
        Next cell
    End If
    For Each part In watched.Areas
        For r = part.Row To part.Row + part.Rows.Count - 1: rowsChanged(CStr(r)) = True: Next r
    Next part
    For Each key In rowsChanged.Keys
        r = CLng(key)
        If Not IsParentTask(r) Then
            If Not Intersect(Target, Me.Cells(r, 8)) Is Nothing Then
                If Len(CStr(Me.Cells(r, 8).Value2)) = 0 Then
                    Me.Cells(r, 6).ClearContents
                ElseIf ReadWbsDate(r, 7, startDate) And ReadWbsDate(r, 8, endDate) Then
                    If endDate < startDate Then
                        warning = "Planned end must be on or after planned start at row " & r & ".": GoTo Done
                    End If
                    Me.Cells(r, 6).Value = DateDiff("d", startDate, endDate) + 1
                End If
            End If
            If Not Intersect(Target, Union(Me.Cells(r, 6), Me.Cells(r, 7))) Is Nothing Then
                value = Me.Cells(r, 6).Value2
                If Len(CStr(value)) = 0 Or Len(CStr(Me.Cells(r, 7).Value2)) = 0 Then
                    Me.Cells(r, 8).ClearContents
                Else
                    If Not IsNumeric(value) Then warning = "Days must be a positive whole number at row " & r & ".": GoTo Done
                    days = CDbl(value)
                    If days < 1 Or days <> Fix(days) Then warning = "Days must be a positive whole number at row " & r & ".": GoTo Done
                End If
            End If
        End If
    Next key
    RefreshParentTasks rebuild, relink
Done:
    On Error Resume Next
    Application.EnableEvents = False
    If levelChange Then ShowSelectedLevel
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If Len(warning) > 0 Then MsgBox warning, vbExclamation, "WBS input"
    If undoable Then Model_1.FinishWbsEdit Else Model_1.ResetUndoTracking
    Exit Sub
Failed:
    warning = Err.Description: Resume Done
End Sub
