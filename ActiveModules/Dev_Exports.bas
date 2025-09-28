Attribute VB_Name = "Dev_Exports"
Option Explicit

' One-click project export for reproducible debugging and delivery
'
' Exports:
' - Core tables (DataTable, ConfigTable, ModeConfigTable) as CSV to Data_Exports/
' - Environment snapshot (references, workbook path, user/computer, VBA trust) to env.txt
' - VBA modules export via ExportModulesToActiveFolder into ActiveModules/
' - Optional: logs snapshots
'
Private Function TimestampFolder(ByVal root As String) As String
    Dim path As String
    path = root & Application.PathSeparator & "Project_Export_" & Format(Now, "yyyymmdd_hhnnss")
    If Len(Dir(root, vbDirectory)) = 0 Then MkDir root
    MkDir path
    TimestampFolder = path
End Function

Private Sub EnsureFolder(ByVal p As String)
    If Len(Dir(p, vbDirectory)) = 0 Then MkDir p
End Sub

Private Sub ExportListObjectCsv(ByVal loName As String, ByVal filePath As String)
    On Error GoTo EH
    Dim t As ListObject: Set t = lo(loName)
    If t Is Nothing Or t.DataBodyRange Is Nothing Then Exit Sub
    Dim f As Integer: f = FreeFile
    Open filePath For Output As #f
    Dim c As Long, r As Long
    ' headers
    For c = 1 To t.HeaderRowRange.Columns.Count
        If c > 1 Then Print #f, ",";
        Print #f, EscapeCsv(CStr(t.HeaderRowRange.Cells(1, c).Value));
    Next c
    Print #f, ""
    ' rows
    For r = 1 To t.DataBodyRange.Rows.Count
        For c = 1 To t.DataBodyRange.Columns.Count
            If c > 1 Then Print #f, ",";
            Print #f, EscapeCsv(CStr(t.DataBodyRange.Cells(r, c).Value));
        Next c
        Print #f, ""
    Next r
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Private Function EscapeCsv(ByVal s As String) As String
    Dim needsQuote As Boolean
    needsQuote = (InStr(1, s, ",") > 0) Or (InStr(1, s, Chr$(10)) > 0) Or (InStr(1, s, Chr$(13)) > 0) Or (InStr(1, s, '"') > 0)
    If needsQuote Then
        s = '"' & Replace$(s, '"', '""') & '"'
    End If
    EscapeCsv = s
End Function

Private Sub ExportEnvironment(ByVal folder As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open folder & Application.PathSeparator & "env.txt" For Output As #f
    Print #f, "Timestamp: " & Now
    Print #f, "Workbook: " & ThisWorkbook.FullName
    Print #f, "User: " & Application.UserName
    Print #f, "Computer: " & Environ("COMPUTERNAME")
    Print #f, "VBATrust: " & IIf(HasVBATrustAccess(), "Yes", "No")
    ' References
    Dim ref As Reference
    On Error GoTo SkipRefs
    For Each ref In ThisWorkbook.VBProject.References
        Print #f, "REF: " & ref.Name & " | " & ref.Description
    Next ref
SkipRefs:
    Close #f
End Sub

Public Sub RUN_Export_ProjectSnapshot()
    On Error GoTo EH
    Dim root As String: root = ThisWorkbook.Path & Application.PathSeparator & "logs"
    Dim out As String: out = TimestampFolder(root)

    ' Subfolders
    Dim dataDir As String: dataDir = out & Application.PathSeparator & "Data_Exports"
    Dim modDir As String:  modDir = out & Application.PathSeparator & "ActiveModules"
    EnsureFolder dataDir: EnsureFolder modDir

    ' Export tables
    ExportListObjectCsv "DataTable", dataDir & Application.PathSeparator & "DataTable.csv"
    ExportListObjectCsv "ConfigTable", dataDir & Application.PathSeparator & "ConfigTable.csv"
    ExportListObjectCsv "ModeConfigTable", dataDir & Application.PathSeparator & "ModeConfigTable.csv"

    ' Export modules (to ActiveModules within snapshot)
    Dim origFolder As String: origFolder = GetActiveModulesFolder()
    ' temporarily point Export to snapshot folder by copying files
    ExportModulesToActiveFolder
    ' Copy exported modules into snapshot
    CopyFolderSafe origFolder, modDir

    ' Environment
    ExportEnvironment out

    MsgBox "Project snapshot exported to:" & vbCrLf & out, vbInformation, "Export Complete"
    Exit Sub
EH:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

Public Sub RUN_Test_And_Export()
    On Error GoTo EH
    ' 1) Sync modules (safe)
    SyncModules_FromActiveFolder
    ' 2) Smoke test
    RUN_SmokeTest_Workbook
    ' 3) Export snapshot
    RUN_Export_ProjectSnapshot
    Exit Sub
EH:
    MsgBox "Test+Export failed: " & Err.Description, vbCritical
End Sub

Private Sub CopyFolderSafe(ByVal src As String, ByVal dst As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(Dir(dst, vbDirectory)) = 0 Then MkDir dst
    If Not fso Is Nothing Then fso.CopyFolder src & "*", dst & Application.PathSeparator, True
End Sub
