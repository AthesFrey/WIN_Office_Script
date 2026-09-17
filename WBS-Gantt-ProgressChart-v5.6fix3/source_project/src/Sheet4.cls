Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const HISTORY_LIMIT As Long = 3
Private mBefore As Object
Private mPending As Object
Private mUndo As Collection
Private mRedo As Collection
Private mRetry As Object
Private mRetryDirection As Boolean
Private mReplaying As Boolean
Private mDepth As Long
Private mInitializing As Boolean
Private mUpdatingUI As Boolean

Private Function TrackedCells() As Range
    Set TrackedCells = Union(Me.Range("B14:J213"), Me.Range("C6:C7"))
End Function

Private Function TakeSnapshot() As Object
    Dim snapshot As Object
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.Add "body", Model_1.CaptureRangeState(Me.Range("B14:J213"), Me.Range("B14:J213"))
    snapshot.Add "settings", Model_1.CaptureRangeState(Me.Range("C6:C7"), Me.Range("C6:C7"))
    Set TakeSnapshot = snapshot
End Function

Private Function SnapshotMatches(ByVal snapshot As Object, Optional ByVal changed As Range = Nothing) As Boolean
    If snapshot Is Nothing Then Exit Function
    If Not Model_1.RangeStateMatches(Me, snapshot("body"), changed) Then Exit Function
    SnapshotMatches = Model_1.RangeStateMatches(Me, snapshot("settings"), changed)
End Function

Private Sub RestoreSnapshot(ByVal snapshot As Object)
    Model_1.RestoreRangeState Me, snapshot("body")
    Model_1.RestoreRangeState Me, snapshot("settings")
    ThisWorkbook.Worksheets("WBS").Calculate
    Me.Calculate
End Sub

Private Sub EnsureHistory()
    If mUndo Is Nothing Then Set mUndo = New Collection
    If mRedo Is Nothing Then Set mRedo = New Collection
End Sub

Private Sub ClearHistory()
    Set mUndo = New Collection: Set mRedo = New Collection
    Set mBefore = Nothing: Set mPending = Nothing: Set mRetry = Nothing
    mDepth = 0
End Sub

Private Function HistoryState() As Object
    Dim entry As Object
    EnsureHistory
    If Not mRetry Is Nothing Then
        Set HistoryState = mRetry
    ElseIf mRedo.Count > 0 Then
        Set entry = mRedo(mRedo.Count): Set HistoryState = entry("before")
    ElseIf mUndo.Count > 0 Then
        Set entry = mUndo(mUndo.Count): Set HistoryState = entry("after")
    End If
End Function

Private Sub ValidateHistory()
    Dim expected As Object
    If mReplaying Or mDepth > 0 Then Exit Sub
    Set expected = HistoryState
    If expected Is Nothing Then Exit Sub
    If Not SnapshotMatches(expected) Then ClearHistory
End Sub

Public Sub CaptureUndoSnapshot()
    If mReplaying Or mDepth > 0 Then Exit Sub
    ValidateHistory
    Set mBefore = Nothing
    Set mBefore = TakeSnapshot
End Sub

Public Sub ResetUndoTracking()
    ClearHistory
    Set mBefore = TakeSnapshot
    Model_1.UpdateHistoryUI
End Sub

Public Sub InitializeEditor(Optional ByVal resetHistory As Boolean = False)
    Dim errorNumber As Long, errorText As String
    If mInitializing Or mReplaying Or mDepth > 0 Then Exit Sub
    Model_1.BeginHistoryMaintenance
    mInitializing = True
    On Error GoTo Failed
    If resetHistory Then ClearHistory
    CaptureUndoSnapshot
Done:
    On Error Resume Next
    Model_1.UpdateHistoryUI
    mInitializing = False
    Model_1.EndHistoryMaintenance
    On Error GoTo 0
    If errorNumber <> 0 Then Model_1.ReportHistoryFailure "Milestones editor", errorNumber, errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Model_1.RememberHistoryFailure "Milestones editor", errorNumber, errorText
    ClearHistory
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
    If Model_1.HistoryMaintenanceActive Then Exit Sub
    If mReplaying Or mDepth > 0 Then Exit Sub
    ' Snapshot cleanup never calls this event or initialization again.
    RefreshSelectionSnapshot
End Sub

Private Sub RefreshSelectionSnapshot()
    Dim errorNumber As Long, errorText As String
    If mInitializing Then Exit Sub
    Model_1.BeginHistoryMaintenance
    mInitializing = True
    On Error GoTo Failed
    CaptureUndoSnapshot
Done:
    On Error Resume Next
    Model_1.UpdateHistoryUI
    mInitializing = False
    Model_1.EndHistoryMaintenance
    On Error GoTo 0
    If errorNumber <> 0 Then Model_1.ReportHistoryFailure "Milestones selection", errorNumber, errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Model_1.RememberHistoryFailure "Milestones selection", errorNumber, errorText
    ClearHistory
    Resume Done
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watched As Range, errorText As String
    If Model_1.HistoryMaintenanceActive Then Exit Sub
    If mReplaying Or mDepth > 0 Then Exit Sub
    On Error GoTo Failed
    Set watched = Application.Intersect(Target, TrackedCells)
    If watched Is Nothing Then
        InitializeEditor
        Exit Sub
    End If
    ' Whole rows/columns and pastes outside the captured areas are boundaries.
    If watched.CountLarge <> Target.CountLarge Then
        ResetUndoTracking
        Exit Sub
    End If
    BeginEdit Target
    FinishEdit
    Exit Sub
Failed:
    errorText = "Milestones history failed: " & Err.Description
    errorText = errorText & CancelEdit
    MsgBox errorText, vbExclamation, "Milestones"
End Sub

Public Sub BeginEdit(Optional ByVal changed As Range = Nothing)
    If mReplaying Then Exit Sub
    If mDepth > 0 Then mDepth = mDepth + 1: Exit Sub
    EnsureHistory
    If changed Is Nothing Then
        CaptureUndoSnapshot
    ElseIf Not mBefore Is Nothing Then
        ' Worksheet_Change runs after input; only untouched cells can be checked.
        If Not SnapshotMatches(mBefore, changed) Then ClearHistory
    End If
    Set mPending = mBefore
    mDepth = 1
End Sub

Public Sub FinishEdit()
    Dim after As Object, entry As Object
    If mReplaying Or mDepth = 0 Then Exit Sub
    mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub
    If mPending Is Nothing Then ResetUndoTracking: Exit Sub
    If Not SnapshotMatches(mPending) Then
        Set after = TakeSnapshot
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "before", mPending: entry.Add "after", after
        If Not mRetry Is Nothing Then Set mUndo = New Collection
        mUndo.Add entry
        If mUndo.Count > HISTORY_LIMIT Then mUndo.Remove 1
        Set mRedo = New Collection: Set mRetry = Nothing
        Set mBefore = after
    Else
        Set mBefore = mPending
    End If
    Set mPending = Nothing
    Model_1.UpdateHistoryUI
End Sub

Public Function CancelEdit() As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    If Not mPending Is Nothing Then
        mReplaying = True
        Application.EnableEvents = False: Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        If Not SnapshotMatches(mPending) Then RestoreSnapshot mPending
        If Not SnapshotMatches(mPending) Then Err.Raise vbObjectError + 510, , "The original cells could not be verified."
        Set mBefore = TakeSnapshot
    End If
Done:
    Set mPending = Nothing: mDepth = 0
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    mReplaying = False
    Model_1.UpdateHistoryUI
    On Error GoTo 0
    Exit Function
Failed:
    CancelEdit = vbCrLf & "Rollback failed: " & Err.Description & ". History was cleared; review the cells."
    ClearHistory
    Resume Done
End Function

Public Function UndoCount() As Long
    EnsureHistory
    UndoCount = mUndo.Count
End Function

Public Function RedoCount() As Long
    EnsureHistory
    RedoCount = mRedo.Count
End Function

Public Sub UpdateHistoryUI()
    Dim canUndo As Boolean, canRedo As Boolean
    If mUpdatingUI Then Exit Sub
    Model_1.BeginHistoryMaintenance
    mUpdatingUI = True
    On Error GoTo Done
    EnsureHistory
    canUndo = (mUndo.Count > 0): canRedo = (mRedo.Count > 0)
    If Not mRetry Is Nothing Then
        canUndo = canUndo And Not mRetryDirection
        canRedo = canRedo And mRetryDirection
    End If
    ' Shape protection must not prevent shortcuts or error cleanup.
    On Error Resume Next
    SetHistoryButton "UndoMilestones", "Undo", mUndo.Count, canUndo
    SetHistoryButton "RedoMilestones", "Redo", mRedo.Count, canRedo
Done:
    On Error Resume Next
    mUpdatingUI = False
    Model_1.EndHistoryMaintenance
    On Error GoTo 0
End Sub

Private Sub SetHistoryButton(ByVal name As String, ByVal label As String, ByVal count As Long, ByVal enabled As Boolean)
    Dim shp As Shape, caption As String, color As Long
    Set shp = Me.Shapes("PG_Button_" & name)
    caption = label & " (" & count & ")"
    color = IIf(enabled, RGB(37, 99, 235), RGB(148, 163, 184))
    If shp.TextFrame2.TextRange.Text <> caption Then shp.TextFrame2.TextRange.Text = caption
    If shp.Fill.ForeColor.RGB <> color Then shp.Fill.ForeColor.RGB = color
End Sub

Public Sub UndoLastEdit(Optional ByVal propagateErrors As Boolean = False)
    ReplayHistory False, propagateErrors
End Sub

Public Sub RedoLastEdit(Optional ByVal propagateErrors As Boolean = False)
    ReplayHistory True, propagateErrors
End Sub

Private Sub ReplayHistory(ByVal redo As Boolean, ByVal propagateErrors As Boolean)
    Dim source As Collection, destination As Collection, entry As Object, target As Object, original As Object, restored As Object
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalculation As XlCalculation
    Dim errorText As String, rollbackError As String, started As Boolean
    If mReplaying Or mDepth > 0 Then Exit Sub
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Or Not ActiveSheet Is Me Then Exit Sub
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    EnsureHistory
    ValidateHistory
    If Not mRetry Is Nothing Then
        If redo <> mRetryDirection Then GoTo Done
    End If
    If redo Then
        Set source = mRedo: Set destination = mUndo
    Else
        Set source = mUndo: Set destination = mRedo
    End If
    If source.Count = 0 Then GoTo Done
    Set entry = source(source.Count)
    If redo Then Set target = entry("after") Else Set target = entry("before")
    Set original = TakeSnapshot
    If Me.ProtectContents Then Err.Raise vbObjectError + 510, , "Unprotect Milestones before retrying."
    mReplaying = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    started = True
    RestoreSnapshot target
    If Not SnapshotMatches(target) Then Err.Raise vbObjectError + 510, , "The restored cells could not be verified."
    Set restored = TakeSnapshot
    Me.Range("B10").Value = IIf(redo, "Redo", "Undo") & " completed. Click Generate to update the chart."
    destination.Add entry
    source.Remove source.Count
    Set mBefore = restored: Set mRetry = Nothing: Set mPending = Nothing
Done:
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    mReplaying = False
    Model_1.UpdateHistoryUI
    On Error GoTo 0
    If Len(errorText) > 0 Then
        If propagateErrors Then Err.Raise vbObjectError + 510, "Milestones history", errorText
        MsgBox errorText & vbCrLf & "History is retained. Resolve the error, then retry.", vbExclamation, "Milestones history"
    End If
    Exit Sub
Failed:
    errorText = IIf(redo, "Redo", "Undo") & " failed: " & Err.Description
    On Error Resume Next
    If started Then
        Err.Clear
        RestoreSnapshot original
        If Err.Number <> 0 Then rollbackError = vbCrLf & "Rollback failed: " & Err.Description & ". Review the cells."
    End If
    If Not original Is Nothing Then
        Set mRetry = TakeSnapshot
        mRetryDirection = redo: Set mBefore = mRetry
    End If
    errorText = errorText & rollbackError
    On Error GoTo 0
    GoTo Done
End Sub
