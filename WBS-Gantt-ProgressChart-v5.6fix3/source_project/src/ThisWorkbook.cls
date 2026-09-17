Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private mOpening As Boolean

Private Sub Workbook_Open()
    Dim errorNumber As Long, errorText As String
    If mOpening Then Exit Sub
    Model_1.BeginHistoryMaintenance
    mOpening = True
    On Error GoTo Failed
    ' Each editor captures its baseline once; internal events remain disabled.
    Sheet1.InitializeEditor True
    Sheet4.InitializeEditor True
Done:
    On Error Resume Next
    mOpening = False
    Model_1.EndHistoryMaintenance
    On Error GoTo 0
    If errorNumber <> 0 Then Model_1.ReportHistoryFailure "Workbook open", errorNumber, errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Done
End Sub

Private Sub Workbook_Activate()
    InitializeActiveEditor
End Sub

Private Sub InitializeActiveEditor()
    If mOpening Or Model_1.HistoryMaintenanceActive Then Exit Sub
    If ActiveSheet Is Sheet1 Then Sheet1.InitializeEditor
    If ActiveSheet Is Sheet4 Then Sheet4.InitializeEditor
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim part As Range
    If Model_1.HistoryMaintenanceActive Then Exit Sub
    If Not Sh Is Sheet1 Then Exit Sub
    For Each part In Target.Areas
        If part.Columns.Count = Sh.Columns.Count Or part.Rows.Count = Sh.Rows.Count Then
            ' Deleted milestone entries have no live formulas for Excel to
            ' retarget. Never replay their old links after a source row move.
            Sheet4.ResetUndoTracking
            Exit Sub
        End If
    Next part
End Sub

Private Sub Workbook_Deactivate()
    ' Also runs on actual close. Do not release in BeforeClose: Excel may
    ' cancel that close at the save prompt while this workbook stays active.
    If Not Model_1.HistoryMaintenanceActive Then Model_1.ReleaseHistoryKeys
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    ' Save As can change the workbook-qualified keyboard callback name.
    If Success Then Model_1.UpdateHistoryUI
End Sub
