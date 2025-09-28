Attribute VB_Name = "Dev_SmokeTests"
Option Explicit

Public Sub RUN_SmokeTest_Workbook()
    On Error GoTo EH

    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Then Err.Raise 5, , "DataTable not found"
    If dataLo.DataBodyRange Is Nothing Then Err.Raise 5, , "DataTable empty"

    Dim needHeaders As Variant
    needHeaders = Array("Functional System Category", "Functional System", "Equipment Description", "Object Type", "SAP Equipment ID")
    Dim i As Long
    For i = LBound(needHeaders) To UBound(needHeaders)
        If HeaderIndexByText(dataLo, CStr(needHeaders(i))) = 0 Then
            Err.Raise 5, , "Missing header: " & CStr(needHeaders(i))
        End If
    Next i

    ' Simple ModeConfig sanity
    Dim modeLo As ListObject: Set modeLo = lo("ModeConfigTable")
    If modeLo Is Nothing Or modeLo.DataBodyRange Is Nothing Then Err.Raise 5, , "ModeConfigTable not found or empty"

    ' Try a minimal OutputAllVisible (will throw if config is broken)
    OutputAllVisible

    MsgBox "Smoke test passed.", vbInformation
    Exit Sub
EH:
    MsgBox "Smoke test failed: " & Err.Description, vbCritical
End Sub
