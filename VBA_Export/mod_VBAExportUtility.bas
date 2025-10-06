Attribute VB_Name = "mod_VBAExportUtility"
'Attribute VB_Name = "mod_VBAExportUtility"  ' Commented out for manual copy/paste convenience
Option Explicit

' ============================================================================
' Module:        mod_VBAExportUtility
' Purpose:       Reliable export of every VBA component (modules, classes,
'                userforms, sheet / workbook modules) to text files so the
'                codebase can be version controlled and reviewed / refactored
'                externally (e.g. in GitHub + AI tooling).
' ----------------------------------------------------------------------------
' Why this exists:
'   You requested a complete review & repair of the search system without
'   assuming existing code quality. To do that we first need authoritative
'   source of truth outside the binary XLSM. This exporter creates that.
' ----------------------------------------------------------------------------
' Key Features:
'   * Exports ALL components, choosing correct extension (.bas/.cls/.frm)
'   * Creates a metadata manifest summarizing component types & timestamps
'   * Optional filtering (future enhancement hook)
'   * Safe overwrite (deletes prior file first to avoid stale content issues)
'   * Basic error logging to Immediate Window and optional Diagnostics sheet
' ----------------------------------------------------------------------------
' Security / Trust Notes:
'   Requires: Trust access to the VBA project object model
'             (File > Options > Trust Center > Trust Center Settings >
'              Macro Settings > CHECK "Trust access to the VBA project...")
' ----------------------------------------------------------------------------
' Integration Points / Future:
'   - A companion importer could reconstitute modules (not needed now)
'   - Can be extended to hash file contents for change detection
'   - Can be wired into a Workbook_BeforeClose event for automatic sync
' ----------------------------------------------------------------------------
' Usage:
'   Call: ExportAllVba  (defaults to folder of the active workbook)
'   or:   ExportAllVba "C:\\Path\\To\\Repo"
' ============================================================================

Private Const DIAG_SHEET_NAME As String = "__Diagnostics"

Public Sub ExportAllVba(Optional ByVal TargetPath As String = "")
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ThisWorkbook

    If Len(TargetPath) = 0 Then
        TargetPath = wb.path
    End If
    If Len(TargetPath) = 0 Then
        Err.Raise vbObjectError + 513, , "Workbook not saved yet – cannot derive export path. Please save the workbook first."
    End If

    If Right$(TargetPath, 1) <> Application.PathSeparator Then
        TargetPath = TargetPath & Application.PathSeparator
    End If

    Dim exportFolder As String
    exportFolder = TargetPath & "VBA_Export" & Application.PathSeparator
    EnsureFolder exportFolder

    Dim manifestLines As Collection
    Set manifestLines = New Collection
    manifestLines.Add "ComponentName,Type,FileName,ExportedUTC"

    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        ExportComponent vbComp, exportFolder, manifestLines
    Next vbComp

    ' Write manifest
    Dim manifestPath As String
    manifestPath = exportFolder & "_manifest.csv"
    WriteTextFile manifestPath, JoinCollection(manifestLines, vbCrLf)

    Debug.Print "[ExportAllVba] Completed. Files in: " & exportFolder
    UpdateDiagnosticsSheet "VBA Export completed to: " & exportFolder
    Exit Sub
EH:
    Debug.Print "[ExportAllVba][ERROR] " & Err.Number & " - " & Err.Description
    UpdateDiagnosticsSheet "VBA Export ERROR: " & Err.Number & " - " & Err.Description
End Sub

Private Sub ExportComponent(ByVal vbComp As VBIDE.VBComponent, _
                            ByVal exportFolder As String, _
                            ByRef manifestLines As Collection)
    On Error GoTo EH

    Dim ext As String
    Select Case vbComp.Type
        Case vbext_ct_ClassModule: ext = ".cls"
        Case vbext_ct_MSForm:      ext = ".frm"
        Case vbext_ct_StdModule:   ext = ".bas"
        Case vbext_ct_Document:    ext = ".cls"   ' Sheet / ThisWorkbook
        Case Else:                 ext = ".txt"
    End Select

    Dim fileName As String
    fileName = SanitizeFileName(vbComp.name) & ext
    Dim fullPath As String
    fullPath = exportFolder & fileName

    ' Remove existing to avoid stale code lines persisting
    If Len(Dir$(fullPath, vbNormal)) > 0 Then
        VBA.Kill fullPath
    End If

    vbComp.Export fullPath

    manifestLines.Add vbComp.name & "," & ComponentTypeString(vbComp.Type) & "," & fileName & "," & Format$(Now, "yyyy-mm-ddThh:nn:ssZ")
    Debug.Print "[Export] " & vbComp.name & " -> " & fileName
    Exit Sub
EH:
    Debug.Print "[ExportComponent][ERROR] " & vbComp.name & " - " & Err.Number & " - " & Err.Description
End Sub

Private Function ComponentTypeString(t As VBIDE.vbext_ComponentType) As String
    Select Case t
        Case vbext_ct_ClassModule: ComponentTypeString = "ClassModule"
        Case vbext_ct_MSForm:      ComponentTypeString = "UserForm"
        Case vbext_ct_StdModule:   ComponentTypeString = "StdModule"
        Case vbext_ct_Document:    ComponentTypeString = "Document"
        Case Else:                 ComponentTypeString = "Other"
    End Select
End Function

Private Sub EnsureFolder(ByVal path As String)
    ' Error 52 (Bad file name or number) can occur if MkDir is called with a trailing
    ' path separator (e.g., "C:\Folder\\") or an otherwise malformed path.
    ' We normalize by trimming any trailing separators before checking/creating.
    On Error GoTo EH
    Dim cleanPath As String
    cleanPath = path
    Do While Right$(cleanPath, 1) = Application.PathSeparator And Len(cleanPath) > 3
        cleanPath = Left$(cleanPath, Len(cleanPath) - 1)
    Loop
    If Len(Dir$(cleanPath, vbDirectory)) = 0 Then
        MkDir cleanPath
    End If
    Exit Sub
EH:
    Debug.Print "[EnsureFolder][ERROR] " & Err.Number & " - " & Err.Description & " Path=" & cleanPath
    ' Propagate so caller can decide; silently failing would hide export failures.
    Err.Raise Err.Number, "EnsureFolder", Err.Description
End Sub

Private Function SanitizeFileName(ByVal rawName As String) As String
    Dim invalidChars As Variant
    ' NOTE: Previous version had a malformed string for the double quote character ("\""), causing a syntax error.
    ' We explicitly insert a double quote via Chr$(34) to avoid quote escaping confusion.
    invalidChars = Array("\\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    Dim c As Variant
    For Each c In invalidChars
        rawName = Replace(rawName, c, "_")
    Next c
    SanitizeFileName = rawName
End Function

Private Sub WriteTextFile(ByVal fullPath As String, ByVal content As String)
    Dim f As Integer
    f = FreeFile
    Open fullPath For Output As #f
    Print #f, content;
    Close #f
End Sub

Private Function JoinCollection(col As Collection, delim As String) As String
    Dim tmp() As String
    ReDim tmp(1 To col.Count)
    Dim i As Long
    For i = 1 To col.Count
        tmp(i) = CStr(col(i))
    Next i
    JoinCollection = Join(tmp, delim)
End Function

' ================================= Diagnostics Helpers ======================
Public Sub QuickDiagnostics_RunAll()
    ' Purpose: Provide immediate, lightweight snapshot of key config tables
    ' so we can reason about the search engine without opening every sheet.
    On Error GoTo EH
    Dim ws As Worksheet
    Dim lo As ListObject

    UpdateDiagnosticsSheet "Diagnostics started @ " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    ' Attempt to locate ModeConfigTable across all sheets (robust to renames)
    Dim found As Boolean
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.name) = LCase$("ModeConfigTable") Then
                DumpListObject lo, "ModeConfigTable"
                found = True
                Exit For
            End If
        Next lo
        If found Then Exit For
    Next ws
    If Not found Then
        UpdateDiagnosticsSheet "WARNING: ModeConfigTable not found in any worksheet."
    End If

    UpdateDiagnosticsSheet "Diagnostics complete." & vbCrLf
    Debug.Print "[Diagnostics] Complete"
    Exit Sub
EH:
    UpdateDiagnosticsSheet "Diagnostics ERROR: " & Err.Number & " - " & Err.Description
    Debug.Print "[Diagnostics][ERROR] " & Err.Number & " - " & Err.Description
End Sub

Private Sub DumpListObject(ByVal lo As ListObject, ByVal label As String)
    On Error Resume Next
    Dim arr
    arr = lo.DataBodyRange.Value
    Dim headers
    headers = lo.HeaderRowRange.Value

    UpdateDiagnosticsSheet "-- Table: " & label & " (" & lo.name & ") Rows=" & lo.DataBodyRange.Rows.Count & " Cols=" & lo.DataBodyRange.Columns.Count

    Dim r As Long, c As Long
    Dim line As String
    ' Header line
    line = "# "
    For c = 1 To lo.DataBodyRange.Columns.Count
        line = line & headers(1, c) & IIf(c < lo.DataBodyRange.Columns.Count, "|", "")
    Next c
    UpdateDiagnosticsSheet line

    Const MAX_ROWS As Long = 25 ' Avoid huge dumps
    For r = 1 To Application.Min(UBound(arr, 1), MAX_ROWS)
        line = ""
        For c = 1 To UBound(arr, 2)
            line = line & arr(r, c) & IIf(c < UBound(arr, 2), "|", "")
        Next c
        UpdateDiagnosticsSheet line
    Next r
    If lo.DataBodyRange.Rows.Count > MAX_ROWS Then
        UpdateDiagnosticsSheet "(Truncated after " & MAX_ROWS & " rows)"
    End If
End Sub

Private Sub UpdateDiagnosticsSheet(ByVal msg As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DIAG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = DIAG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden ' Keep out of general user view
    End If

    With ws
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = msg
    End With
End Sub

' ============================================================================
' End of module
' ============================================================================


