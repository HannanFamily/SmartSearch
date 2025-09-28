Attribute VB_Name = "QuickSearchDiagnostics"
Option Explicit
'
' QuickSearchDiagnostics.bas
' Purpose: One-click diagnostics to help understand why search is not returning results.
' It summarizes key named ranges, data table presence, visible row counts,
' and runs the built-in diagnostics found in mod_PrimaryConsolidatedModule:
'   - RunConfigDiagnostics
'   - DiagnosticTrace_PerformSearch
'
Public Sub RunQuickSearchDiagnostics()
    On Error GoTo EH
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = EnsureSheet("Diagnostics_Summary")
    ws.Cells.Clear

    Dim r As Long: r = 1
    Title ws, r, "Search Diagnostics Summary": r = r + 2

    InfoRow ws, r, "Workbook Path", ThisWorkbook.path: r = r + 1
    InfoRow ws, r, "Workbook Name", ThisWorkbook.name: r = r + 2

    ' Named search input ranges from config
    Dim nm As String, exists As String

    nm = GetConfigValueSafe("InputCell_DescripSearch"): InfoRow ws, r, "Config: InputCell_DescripSearch", nm & ExistsSuffix(ExistsName(nm)): r = r + 1
    nm = GetConfigValueSafe("InputCell_ValveNumSearch"): InfoRow ws, r, "Config: InputCell_ValveNumSearch", nm & ExistsSuffix(ExistsName(nm)): r = r + 1
    nm = GetConfigValueSafe("ResultsStartCell"):      InfoRow ws, r, "Config: ResultsStartCell", nm & ExistsSuffix(ExistsName(nm)): r = r + 1
    nm = GetConfigValueSafe("StatusCell"):            InfoRow ws, r, "Config: StatusCell", nm & ExistsSuffix(ExistsName(nm)): r = r + 2

    ' Try to resolve the primary Data Table
    Dim dataLo As ListObject
    Set dataLo = ResolveDataTable()
    If dataLo Is Nothing Then
        InfoRow ws, r, "Data Table", "NOT FOUND": r = r + 2
    Else
        InfoRow ws, r, "Data Table", dataLo.name & " on sheet '" & dataLo.Parent.name & "'": r = r + 1
        InfoRow ws, r, "Data Table Rows", CStr(dataLo.DataBodyRange.Rows.count): r = r + 1
        InfoRow ws, r, "Data Table Cols", CStr(dataLo.ListColumns.count): r = r + 2

        ' Visible row count
        Dim idxs As Variant
        idxs = VisibleRowIndexesSafe(dataLo)
        InfoRow ws, r, "Visible Row Count", CStr(VariantCount(idxs)): r = r + 2
    End If

    ' Attempt to get description column index (if helper exists)
    Dim descIdx As Long
    On Error Resume Next
    If Not dataLo Is Nothing Then descIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)
    On Error GoTo 0
    If descIdx > 0 Then
        InfoRow ws, r, "Description Column Index", CStr(descIdx): r = r + 2
    End If

    ' Run built-in diagnostics to produce detailed sheets
    InfoRow ws, r, "Action", "Running RunConfigDiagnostics...": r = r + 1
    On Error Resume Next: RunConfigDiagnostics: On Error GoTo EH

    InfoRow ws, r, "Action", "Running DiagnosticTrace_PerformSearch...": r = r + 1
    On Error Resume Next: DiagnosticTrace_PerformSearch: On Error GoTo EH

    r = r + 1
    Note ws, r, "Outputs": r = r + 1
    Bullet ws, r, "ConfigDiagnostics sheet lists your ConfigTable keys, named ranges, and header matches": r = r + 1
    Bullet ws, r, "SearchDiagnostics sheet traces the search flow and regex build": r = r + 1

    r = r + 1
    InfoRow ws, r, "Next Steps", "If search still returns nothing:"
    Bullet ws, r + 1, "Verify the InputCell_DescripSearch and InputCell_ValveNumSearch named ranges point to cells you edit":
    Bullet ws, r + 2, "Ensure ResultsStartCell and StatusCell are valid cells on your dashboard":
    Bullet ws, r + 3, "Confirm the Data Table name matches your configuration and has visible rows":

    ' Auto-export diagnostics to timestamped files in Diagnostic_Notes
    ExportDiagnosticsToFiles True

    Application.ScreenUpdating = True
    MsgBox "Diagnostics complete and exported to Diagnostic_Notes. See 'Diagnostics_Summary', 'ConfigDiagnostics', and 'SearchDiagnostics' sheets.", vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Diagnostics error: " & Err.DESCRIPTION, vbExclamation
End Sub

'============================================================
' Export diagnostics sheets to files we can read here
'============================================================
Public Sub ExportDiagnosticsToFiles(Optional ByVal silent As Boolean = False)
    On Error GoTo EH
    Dim sep As String: sep = Application.PathSeparator
    Dim outDir As String
    If Len(Trim$(ThisWorkbook.path)) > 0 Then
        outDir = ThisWorkbook.path & sep & "Diagnostic_Notes"
    Else
        ' Fallback for unsaved workbook: user's Documents
        outDir = Environ$("USERPROFILE") & sep & "Documents" & sep & "Diagnostic_Notes"
    End If
    EnsureFolder outDir

    Dim ts As String
    ts = Format(Now, "yyyymmdd_hhnnss")

    Dim exported As String
    exported = ""

    exported = exported & ExportSheetTSV("Diagnostics_Summary", outDir & sep & "Diagnostics_Summary_" & ts & ".tsv")
    exported = exported & ExportSheetTSV("ConfigDiagnostics", outDir & sep & "ConfigDiagnostics_" & ts & ".tsv")
    exported = exported & ExportSheetTSV("SearchDiagnostics", outDir & sep & "SearchDiagnostics_" & ts & ".tsv")

    ' Log result onto Diagnostics_Summary sheet for traceability
    LogExportResult outDir, exported, ts

    If Not silent Then
        If Len(Trim$(exported)) = 0 Then
            MsgBox "No diagnostics sheets found to export.", vbExclamation
        Else
            MsgBox "Diagnostics exported to: " & outDir & vbCrLf & exported, vbInformation
        End If
    End If
    Exit Sub
EH:
    MsgBox "Export error: " & Err.DESCRIPTION, vbExclamation
End Sub

Private Function ExportSheetTSV(ByVal sheetName As String, ByVal filePath As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim f As Integer: f = FreeFile
    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    lastRow = ur.row + ur.Rows.count - 1
    lastCol = ur.Column + ur.Columns.count - 1

    On Error Resume Next
    Open filePath For Output As #f
    On Error GoTo 0
    If Err.Number <> 0 Then Exit Function

    For r = ur.row To lastRow
        Dim line As String
        line = ""
        For c = ur.Column To lastCol
            Dim v As String
            v = CStr(ws.Cells(r, c).value)
            ' Escape tabs and newlines
            v = Replace(v, vbTab, " ")
            v = Replace(v, vbCr, " ")
            v = Replace(v, vbLf, " ")
            If c > ur.Column Then line = line & vbTab
            line = line & v
        Next c
        Print #f, line
    Next r
    Close #f

    ExportSheetTSV = filePath & vbCrLf
End Function

Private Sub EnsureFolder(ByVal path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Sub LogExportResult(ByVal outDir As String, ByVal exported As String, ByVal ts As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Diagnostics_Summary")
    If ws Is Nothing Then Exit Sub
    Dim r As Long
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 2
    ws.Cells(r, 1).value = "Export Log (" & ts & ")"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r + 1, 1).value = "Output Folder:"
    ws.Cells(r + 1, 2).value = outDir
    ws.Cells(r + 2, 1).value = "Files:"
    ws.Cells(r + 2, 2).value = IIf(Len(Trim$(exported)) = 0, "<none>", exported)
End Sub

Private Function ResolveDataTable() As ListObject
    ' Try known names, then pick largest table as fallback
    Dim ws As Worksheet, l As ListObject
    Dim best As ListObject
    Dim bestRows As Long

    Dim candidates As Variant
    candidates = Array("tbl_Equipment")

    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        Set best = loSafe(CStr(candidates(i)))
        If Not best Is Nothing Then Set ResolveDataTable = best: Exit Function
    Next i

    ' Fallback: largest ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each l In ws.ListObjects
            If Not l.DataBodyRange Is Nothing Then
                If l.DataBodyRange.Rows.count > bestRows Then
                    bestRows = l.DataBodyRange.Rows.count
                    Set best = l
                End If
            End If
        Next l
    Next ws
    Set ResolveDataTable = best
End Function

Private Function loSafe(ByVal name As String) As ListObject
    On Error Resume Next
    Set loSafe = lo(name)
    On Error GoTo 0
End Function

Private Function VisibleRowIndexesSafe(ByVal dataLo As ListObject) As Variant
    On Error Resume Next
    VisibleRowIndexesSafe = VisibleRowIndexes(dataLo)
    On Error GoTo 0
End Function

Private Function ExistsName(ByVal nm As String) As Boolean
    On Error Resume Next
    If Len(Trim$(nm)) = 0 Then Exit Function
    Dim r As Range: Set r = ThisWorkbook.names(nm).RefersToRange
    ExistsName = Not r Is Nothing
    On Error GoTo 0
End Function

Private Function GetConfigValueSafe(ByVal key As String) As String
    On Error Resume Next
    GetConfigValueSafe = GetConfigValue(key)
    On Error GoTo 0
End Function

Private Function VariantCount(ByVal v As Variant) As Long
    On Error Resume Next
    If IsEmpty(v) Then Exit Function
    VariantCount = UBound(v) - LBound(v) + 1
    On Error GoTo 0
End Function

' ==== small formatting helpers ====
Private Sub Title(ws As Worksheet, ByVal r As Long, ByVal text As String)
    With ws.Cells(r, 1)
        .value = text
        .Font.Bold = True
        .Font.Size = 14
    End With
End Sub

Private Sub InfoRow(ws As Worksheet, ByVal r As Long, ByVal k As String, ByVal v As String)
    ws.Cells(r, 1).value = k
    ws.Cells(r, 2).value = v
    ws.Cells(r, 1).Font.Bold = True
End Sub

Private Sub Note(ws As Worksheet, ByVal r As Long, ByVal text As String)
    With ws.Cells(r, 1)
        .value = text
        .Font.Bold = True
        .Font.Size = 12
    End With
End Sub

Private Sub Bullet(ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Cells(r, 1).value = "- " & text
End Sub

Private Function ExistsSuffix(ByVal exists As Boolean) As String
    If exists Then
        ExistsSuffix = " (exists)"
    Else
        ExistsSuffix = " (missing)"
    End If
End Function

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add
        EnsureSheet.name = nm
    End If
End Function
