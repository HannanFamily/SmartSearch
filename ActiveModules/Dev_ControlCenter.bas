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
