'Attribute VB_Name = "mod_MetadataSnapshot"  ' Commented out for manual copy/paste convenience
Option Explicit

' ============================================================================
' Module:        mod_MetadataSnapshot
' Purpose:       Generate a structured textual snapshot of workbook metadata
'                (worksheets, tables, columns, named ranges) to support the
'                refactor & repair of the search system without guesswork.
' ----------------------------------------------------------------------------
' Output:        Writes to the hidden __Diagnostics sheet (appends) AND returns
'                the path of a generated snapshot file (if requested).
' ----------------------------------------------------------------------------
' Usage:
'   Call Snapshot_WorkbookMetadata            ' Immediate window only
'   Call Snapshot_WorkbookMetadata(True)      ' Also write to Metadata file
' ----------------------------------------------------------------------------
' Notes:
'   Keeps logic separate so SearchEngineCore stays focused on search.
'   Safe to run repeatedly; each snapshot is timestamped.
' ============================================================================

Public Function Snapshot_WorkbookMetadata(Optional WriteFile As Boolean = False, _
                                          Optional ByVal TargetFolder As String = "") As String
    On Error GoTo EH
    Dim lines As Collection
    Set lines = New Collection
    lines.Add "# Workbook Metadata Snapshot @ " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lines.Add "Workbook: " & ThisWorkbook.Name

    Dim ws As Worksheet, lo As ListObject, nm As Name
    lines.Add "-- Worksheets --"
    For Each ws In ThisWorkbook.Worksheets
        lines.Add "WS: " & ws.Index & ": " & ws.Name & " Visible=" & SheetVisibility(ws)
        If ws.ListObjects.Count > 0 Then
            For Each lo In ws.ListObjects
                lines.Add "  Table: " & lo.Name & _
                    " Rows=" & SafeCountRows(lo) & _
                    " Cols=" & lo.ListColumns.Count & _
                    " DataR1C1=" & lo.DataBodyRange.Address(ReferenceStyle:=xlR1C1)
                lines.Add "    Columns: " & TableColumnList(lo)
            Next lo
        End If
    Next ws

    lines.Add "-- Named Ranges (Workbook Scope Only) --"
    For Each nm In ThisWorkbook.Names
        On Error Resume Next
        Dim refAddr As String
        refAddr = "(invalid)"
        If Not nm.RefersToRange Is Nothing Then
            refAddr = nm.RefersToRange.Address(External:=False)
        End If
        On Error GoTo 0
        lines.Add "Name: " & nm.Name & " -> " & nm.RefersTo & " | Addr=" & refAddr
    Next nm

    ' Append to diagnostics sheet for easy in-workbook review
    Dim i As Long
    For i = 1 To lines.Count
        DiagnosticsAppend CStr(lines(i))
    Next i

    If WriteFile Then
        If Len(TargetFolder) = 0 Then TargetFolder = ThisWorkbook.Path
        If Right$(TargetFolder, 1) <> Application.PathSeparator Then _
            TargetFolder = TargetFolder & Application.PathSeparator
        Dim outPath As String
        outPath = TargetFolder & "Workbook_Metadata_Snapshot_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
        WriteText outPath, JoinCollection(lines, vbCrLf)
        Snapshot_WorkbookMetadata = outPath
        DiagnosticsAppend "Metadata snapshot written: " & outPath
    End If
    Exit Function
EH:
    DiagnosticsAppend "Snapshot ERROR: " & Err.Number & " - " & Err.Description
End Function

Private Function SheetVisibility(ws As Worksheet) As String
    Select Case ws.Visible
        Case xlSheetVisible: SheetVisibility = "Visible"
        Case xlSheetHidden: SheetVisibility = "Hidden"
        Case xlSheetVeryHidden: SheetVisibility = "VeryHidden"
        Case Else: SheetVisibility = CStr(ws.Visible)
    End Select
End Function

Private Function TableColumnList(lo As ListObject) As String
    Dim parts() As String
    ReDim parts(1 To lo.ListColumns.Count)
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        parts(i) = lo.ListColumns(i).Name
    Next i
    TableColumnList = Join(parts, ", ")
End Function

Private Function SafeCountRows(lo As ListObject) As Long
    On Error Resume Next
    If lo.DataBodyRange Is Nothing Then
        SafeCountRows = 0
    Else
        SafeCountRows = lo.DataBodyRange.Rows.Count
    End If
End Function

Private Sub DiagnosticsAppend(ByVal msg As String)
    Const SHEET As String = "__Diagnostics"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET
        ws.Visible = xlSheetVeryHidden
    End If
    With ws
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = msg
    End With
End Sub

Private Function JoinCollection(col As Collection, delim As String) As String
    Dim arr() As String
    ReDim arr(1 To col.Count)
    Dim i As Long
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i
    JoinCollection = Join(arr, delim)
End Function

Private Sub WriteText(ByVal fullPath As String, ByVal content As String)
    Dim f As Integer: f = FreeFile
    Open fullPath For Output As #f
    Print #f, content;
    Close #f
End Sub

' ============================================================================
' End Module
' ============================================================================
