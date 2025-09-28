Attribute VB_Name = "Dev_Tools_MOD"
Option Explicit
'
' Dev Tools: Archival and Removal Utilities
' ------------------------------------------------------------
' ArchiveAndRemoveMarkedRows
' - Scans the primary DataTable for rows flagged in a Remove column
' - Appends those rows to an external archive workbook/table
' - Merges schema (adds any missing columns in the archive)
' - Deletes flagged rows from the source table
'
' Config keys (optional; with sensible defaults if absent):
'   - RemoveFlagColumn: Header name of the flag column in DataTable (default: "Remove")
'   - ArchiveWorkbookName: File name for the archive workbook (default: "Archived_Equipment.xlsx")
'   - ArchiveSheetName: Worksheet to host the archive table (default: "Archive")
'   - ArchiveTableName: ListObject name for the archive table (default: "ArchiveTable")
'   - ArchiveTimestampColumn: Name of the timestamp column in archive (default: "ArchivedAt")
'   - ArchiveSourceColumn: Name of the source workbook column (default: "SourceWorkbook")
'
' Dependencies:
'   - Requires helper functions from mod_PrimaryConsolidatedModule: lo, GetConfigValue, HeaderIndexByText
'   - Uses DATA_TABLE_NAME constant from mod_PrimaryConsolidatedModule
'
Private Const DEFAULT_REMOVE_FLAG_HEADER As String = "Remove"
Private Const DEFAULT_ARCHIVE_WB_NAME As String = "Archived_Equipment.xlsx"
Private Const DEFAULT_ARCHIVE_SHEET As String = "Archive"
Private Const DEFAULT_ARCHIVE_TABLE As String = "ArchiveTable"
Private Const DEFAULT_ARCHIVE_TS_COL As String = "ArchivedAt"
Private Const DEFAULT_ARCHIVE_SRC_COL As String = "SourceWorkbook"

Public Sub ArchiveAndRemoveMarkedRows()
    On Error GoTo EH

    ' Ensure no filters are active to avoid "Can't move cells in a filtered range or table"
    ClearAllTableFilters

    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim removeHeader As String
    removeHeader = NzStr(GetConfigValue("RemoveFlagColumn"), DEFAULT_REMOVE_FLAG_HEADER)

    Dim removeColIdx As Long: removeColIdx = HeaderIndexByText(dataLo, removeHeader)
    If removeColIdx = 0 Then
        MsgBox "Remove flag column not found: '" & removeHeader & "'", vbExclamation
        Exit Sub
    End If

    ' Collect rows to remove (scan all, not only visible)
    Dim rng As Range: Set rng = dataLo.DataBodyRange
    Dim r As Long, c As Long
    Dim markVal As Variant

    Dim srcHdr() As String, colCount As Long
    colCount = dataLo.ListColumns.Count
    ReDim srcHdr(1 To colCount)
    For c = 1 To colCount
        srcHdr(c) = CStr(dataLo.HeaderRowRange.Cells(1, c).Value)
    Next c

    ' Build a list of row indices flagged for removal
    Dim delIdx() As Long, delCount As Long
    ReDim delIdx(1 To rng.Rows.Count)
    For r = 1 To rng.Rows.Count
        markVal = rng.Cells(r, removeColIdx).Value
        If IsMarkedForRemoval(markVal) Then
            delCount = delCount + 1
            delIdx(delCount) = r
        End If
    Next r
    If delCount = 0 Then
        MsgBox "No rows are marked for removal (column '" & removeHeader & "').", vbInformation
        Exit Sub
    End If
    ReDim Preserve delIdx(1 To delCount)

    ' Confirm destructive action
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Archive and remove " & delCount & " row(s)?" & vbCrLf & _
                  "They will be appended to the archive workbook and deleted from the source.", _
                  vbQuestion + vbOKCancel, "Archive & Remove")
    If resp <> vbOK Then Exit Sub

    ' Prepare archive workbook/table
    Dim archiveWb As Workbook, archiveWs As Worksheet, archiveLo As ListObject
    Dim tsColName As String, srcColName As String
    tsColName = NzStr(GetConfigValue("ArchiveTimestampColumn"), DEFAULT_ARCHIVE_TS_COL)
    srcColName = NzStr(GetConfigValue("ArchiveSourceColumn"), DEFAULT_ARCHIVE_SRC_COL)

    Set archiveWb = EnsureArchiveWorkbook()
    If archiveWb Is Nothing Then
        MsgBox "Unable to create/open archive workbook.", vbCritical
        Exit Sub
    End If
    Set archiveWs = EnsureArchiveSheet(archiveWb)
    If archiveWs Is Nothing Then
        MsgBox "Unable to create/find archive sheet.", vbCritical
        Exit Sub
    End If
    Set archiveLo = EnsureArchiveTable(archiveWs)
    If archiveLo Is Nothing Then
        MsgBox "Unable to create/find archive table.", vbCritical
        Exit Sub
    End If

    ' Ensure archive schema includes: all source headers + timestamp + source-workbook
    Dim arcHdr() As String
    arcHdr = CurrentHeaders(archiveLo)
    Dim superset() As String: superset = UnionHeaders(arcHdr, srcHdr, tsColName, srcColName)
    If Not HeadersEqual(arcHdr, superset) Then
        EnsureTableHasHeaders archiveLo, superset
    End If

    ' Append rows to archive (row-by-row to preserve table growth semantics)
    Dim i As Long, j As Long, srcVal As Variant, tgtColIdx As Long
    Application.ScreenUpdating = False
    For i = 1 To delCount
        Dim lr As ListRow
        Set lr = archiveLo.ListRows.Add
        For j = LBound(superset) To UBound(superset)
            tgtColIdx = HeaderIndexByName(archiveLo, superset(j))
            If tgtColIdx > 0 Then
                If StrComp(superset(j), tsColName, vbTextCompare) = 0 Then
                    lr.Range.Cells(1, tgtColIdx).Value = Now
                ElseIf StrComp(superset(j), srcColName, vbTextCompare) = 0 Then
                    lr.Range.Cells(1, tgtColIdx).Value = ThisWorkbook.Name
                Else
                    Dim srcIdx As Long: srcIdx = HeaderIndexInArray(srcHdr, superset(j))
                    If srcIdx > 0 Then
                        srcVal = rng.Cells(delIdx(i), srcIdx).Value
                        lr.Range.Cells(1, tgtColIdx).Value = srcVal
                    Else
                        ' Column not present in source -> leave blank
                    End If
                End If
            End If
        Next j
    Next i

    ' Save archive workbook
    On Error Resume Next
    archiveWb.Save
    On Error GoTo 0

    ' Delete from source (bottom-up)
    Application.ScreenUpdating = False
    For i = delCount To 1 Step -1
        dataLo.ListRows(delIdx(i)).Delete
    Next i
    Application.ScreenUpdating = True

    MsgBox "Archived and removed " & delCount & " row(s)." & vbCrLf & _
           "Archive: " & FullArchivePath(), vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Error in ArchiveAndRemoveMarkedRows: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ---------- Helpers ----------
Private Function NzStr(ByVal s As String, ByVal fallback As String) As String
    If Len(Trim$(s)) = 0 Then NzStr = fallback Else NzStr = s
End Function

Private Function IsMarkedForRemoval(ByVal v As Variant) As Boolean
    ' Policy: ANY non-blank value (after Trim) in the Remove column means "mark for removal".
    ' Still respects Boolean True when the cell is a checkbox/boolean.
    On Error Resume Next
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        IsMarkedForRemoval = CBool(v)
        Exit Function
    End If
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) > 0 Then IsMarkedForRemoval = True
End Function

Private Function EnsureArchiveWorkbook() As Workbook
    Dim p As String: p = FullArchivePath()
    On Error Resume Next
    Dim wb As Workbook
    Set wb = GetOpenWorkbookByFullName(p)
    On Error GoTo 0
    If wb Is Nothing Then
        If FileExists(p) Then
            Set wb = Application.Workbooks.Open(Filename:=p)
        Else
            ' Create new archive workbook
            Set wb = Application.Workbooks.Add
            On Error Resume Next
            wb.SaveAs Filename:=p, FileFormat:=xlOpenXMLWorkbook ' .xlsx
            On Error GoTo 0
        End If
    End If
    Set EnsureArchiveWorkbook = wb
End Function

Private Function FullArchivePath() As String
    Dim wbName As String
    wbName = NzStr(GetConfigValue("ArchiveWorkbookName"), DEFAULT_ARCHIVE_WB_NAME)

    Dim basePath As String: basePath = ThisWorkbook.Path
    If Len(Trim$(basePath)) = 0 Then
        basePath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    End If
    If Right$(basePath, 1) = Application.PathSeparator Then
        FullArchivePath = basePath & wbName
    Else
        FullArchivePath = basePath & Application.PathSeparator & wbName
    End If
End Function

Private Function EnsureArchiveSheet(ByVal wb As Workbook) As Worksheet
    Dim sheetName As String
    sheetName = NzStr(GetConfigValue("ArchiveSheetName"), DEFAULT_ARCHIVE_SHEET)
    On Error Resume Next
    Set EnsureArchiveSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureArchiveSheet Is Nothing Then
        Set EnsureArchiveSheet = wb.Worksheets.Add
        EnsureArchiveSheet.Name = sheetName
    End If
End Function

Private Function EnsureArchiveTable(ByVal ws As Worksheet) As ListObject
    Dim tableName As String
    tableName = NzStr(GetConfigValue("ArchiveTableName"), DEFAULT_ARCHIVE_TABLE)

    Dim loT As ListObject
    For Each loT In ws.ListObjects
        If StrComp(loT.Name, tableName, vbTextCompare) = 0 Then Set EnsureArchiveTable = loT: Exit Function
    Next loT

    ' Create a minimal 1-row header and convert to table
    Dim hdrCell As Range
    Set hdrCell = ws.Range("A1")
    If Len(Trim$(CStr(hdrCell.Value))) = 0 Then hdrCell.Value = "ID"
    Set EnsureArchiveTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(1, 1), , xlYes)
    EnsureArchiveTable.Name = tableName
End Function

Private Function CurrentHeaders(ByVal loT As ListObject) As String()
    Dim n As Long, i As Long
    Dim arr() As String
    n = loT.HeaderRowRange.Columns.Count
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CStr(loT.HeaderRowRange.Cells(1, i).Value)
    Next i
    CurrentHeaders = arr
End Function

Private Function HeadersEqual(ByRef a() As String, ByRef b() As String) As Boolean
    Dim na As Long, nb As Long, i As Long
    na = UBound(a) - LBound(a) + 1
    nb = UBound(b) - LBound(b) + 1
    If na <> nb Then Exit Function
    For i = LBound(a) To UBound(a)
        If StrComp(a(i), b(i), vbTextCompare) <> 0 Then Exit Function
    Next i
    HeadersEqual = True
End Function

Private Sub EnsureTableHasHeaders(ByVal loT As ListObject, ByRef headers() As String)
    ' Resize/overwrite the header row to match exactly the headers array
    Dim ws As Worksheet: Set ws = loT.Parent
    Dim cCnt As Long: cCnt = UBound(headers) - LBound(headers) + 1

    ' If the table is smaller or larger, resize its Range to fit new header width
    Dim firstCell As Range: Set firstCell = loT.HeaderRowRange.Cells(1, 1)
    Dim newHeaderRange As Range: Set newHeaderRange = firstCell.Resize(1, cCnt)

    ' Ensure the underlying cells hold the header text
    Dim i As Long
    For i = 1 To cCnt
        newHeaderRange.Cells(1, i).Value = headers(i)
    Next i

    ' Resize the ListObject to cover the new header width + existing rows
    Dim rowCnt As Long
    rowCnt = 0
    On Error Resume Next
    rowCnt = loT.DataBodyRange.Rows.Count
    On Error GoTo 0
    If rowCnt < 0 Then rowCnt = 0
    loT.Resize newHeaderRange.Resize(WorksheetFunction.Max(1, rowCnt + 1), cCnt)
End Sub

Private Function UnionHeaders(ByRef arcHdr() As String, ByRef srcHdr() As String, _
                              ByVal tsCol As String, ByVal srcCol As String) As String()
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(arcHdr) To UBound(arcHdr)
        If Len(Trim$(arcHdr(i))) > 0 Then If Not d.Exists(arcHdr(i)) Then d.Add arcHdr(i), d.Count + 1
    Next i
    For i = LBound(srcHdr) To UBound(srcHdr)
        If Len(Trim$(srcHdr(i))) > 0 Then If Not d.Exists(srcHdr(i)) Then d.Add srcHdr(i), d.Count + 1
    Next i
    If Len(Trim$(tsCol)) > 0 Then If Not d.Exists(tsCol) Then d.Add tsCol, d.Count + 1
    If Len(Trim$(srcCol)) > 0 Then If Not d.Exists(srcCol) Then d.Add srcCol, d.Count + 1

    Dim arr() As String
    ReDim arr(1 To d.Count)
    Dim k As Variant
    For Each k In d.Keys
        arr(d(k)) = CStr(k)
    Next k
    UnionHeaders = arr
End Function

Private Function HeaderIndexByName(ByVal loT As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To loT.HeaderRowRange.Columns.Count
        If StrComp(CStr(loT.HeaderRowRange.Cells(1, i).Value), headerName, vbTextCompare) = 0 Then
            HeaderIndexByName = i: Exit Function
        End If
    Next i
End Function

Private Function HeaderIndexInArray(ByRef arr() As String, ByVal headerName As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(arr(i), headerName, vbTextCompare) = 0 Then HeaderIndexInArray = i: Exit Function
    Next i
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(fullPath, vbNormal)) > 0)
End Function

Private Function GetOpenWorkbookByFullName(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then Set GetOpenWorkbookByFullName = wb: Exit Function
    Next wb
End Function

' Clear AutoFilters on all tables in the workbook (safe no-op if none are filtered)
Private Sub ClearAllTableFilters()
    On Error Resume Next
    Dim ws As Worksheet, t As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each t In ws.ListObjects
            If Not t.AutoFilter Is Nothing Then
                If t.AutoFilter.FilterMode Then t.AutoFilter.ShowAllData
            End If
        Next t
    Next ws
    On Error GoTo 0
End Sub
