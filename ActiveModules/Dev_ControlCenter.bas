Attribute VB_Name = "Dev_ControlCenter"
Option Explicit

' Dev Control Center: one-click helpers to control the workflow safely

Public Sub RUN_Dev_Compile()
    On Error Resume Next
    Application.VBE.ActiveVBProject.BuildFile ""
    If Err.Number <> 0 Then Err.Clear ' BuildFile not available in all hosts; ignore
    On Error GoTo 0
    Debug.Print "Compile requested"
End Sub

Public Sub RUN_Dev_SyncModules()
    On Error GoTo EH
    Call SyncModules_FromActiveFolder
    Exit Sub
EH: MsgBox "Sync failed: " & Err.Description, vbCritical
End Sub

Public Sub RUN_Dev_ReplaceAllModules()
    On Error GoTo EH
    Call ReplaceAllModules_FromActiveFolder
    Exit Sub
EH: MsgBox "Replace failed: " & Err.Description, vbCritical
End Sub

Public Sub RUN_Dev_ExportModules()
    On Error GoTo EH
    Call ExportModulesToActiveFolder
    Exit Sub
EH: MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

Public Sub RUN_Dev_SmokeTest()
    On Error GoTo EH
    Call RUN_SmokeTest_Workbook
    Exit Sub
EH: MsgBox "Smoke test failed: " & Err.Description, vbCritical
End Sub

Public Sub RUN_Dev_AddDashboardButtons()
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = SheetByName("Dashboard")
    If ws Is Nothing Then
        MsgBox "Dashboard sheet not found.", vbExclamation
        Exit Sub
    End If
    Dim shp As Shape
    ' Add Test+Export button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 60, 180, 28)
    shp.TextFrame2.TextRange.Text = "Test + Export"
    shp.OnAction = "Dev_Exports.RUN_Test_And_Export"
    ' Add Export only button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 210, 60, 180, 28)
    shp.TextFrame2.TextRange.Text = "Export Snapshot"
    shp.OnAction = "Dev_Exports.RUN_Export_ProjectSnapshot"
    MsgBox "Buttons added to Dashboard.", vbInformation
    Exit Sub
EH:
    MsgBox "Failed to add buttons: " & Err.Description, vbCritical
End Sub
