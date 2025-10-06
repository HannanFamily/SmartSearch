Attribute VB_Name = "Diagnostic"
Public Sub DiagnosticTrace_PerformSearch()
    Dim wsDiag As Worksheet
    Dim NextRow As Long
    Dim dataLo As ListObject
    Dim searchTxt As String, tagTxt As String
    Dim searchColIdx As Long
    Dim rxArr As Variant, synIndex As Object
    Dim idxs As Variant
    Dim kept As Long, i As Long, ri As Long, p As Long
    Dim descText As String, keep As Boolean
    Dim mapLo As ListObject

    On Error Resume Next
    Set wsDiag = ThisWorkbook.Worksheets("SearchDiagnostics")
    If wsDiag Is Nothing Then
        Set wsDiag = ThisWorkbook.Worksheets.Add
        wsDiag.name = "SearchDiagnostics"
    End If
    On Error GoTo 0

    wsDiag.Cells.Clear
    NextRow = 1
    wsDiag.Cells(NextRow, 1).Value = "Step"
    wsDiag.Cells(NextRow, 2).Value = "Detail"
    NextRow = NextRow + 1

    searchTxt = ReadLeftCell(GetConfigValue("InputCell_DescripSearch"))
    tagTxt = ReadLeftCell(GetConfigValue("InputCell_ValveNumSearch"))

    wsDiag.Cells(NextRow, 1).Value = "Search Text": wsDiag.Cells(NextRow, 2).Value = searchTxt: NextRow = NextRow + 1
    wsDiag.Cells(NextRow, 1).Value = "Tag Text": wsDiag.Cells(NextRow, 2).Value = tagTxt: NextRow = NextRow + 1

    Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Then
        wsDiag.Cells(NextRow, 1).Value = "Error"
        wsDiag.Cells(NextRow, 2).Value = "DataTable not found"
        Exit Sub
    End If

    searchColIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)
    wsDiag.Cells(NextRow, 1).Value = "Description Column Index": wsDiag.Cells(NextRow, 2).Value = searchColIdx: NextRow = NextRow + 1

    Set mapLo = lo(MappingTableName())
    Set synIndex = BuildSynonymIndex(mapLo)
    If Len(Trim$(searchTxt)) > 0 Then
        rxArr = BuildSearchRegexes(searchTxt, synIndex)
    Else
        rxArr = Array()
    End If

    wsDiag.Cells(NextRow, 1).Value = "Regex Array Count"
    If IsArray(rxArr) Then
        wsDiag.Cells(NextRow, 2).Value = UBound(rxArr) - LBound(rxArr) + 1
    Else
        wsDiag.Cells(NextRow, 2).Value = "Not an array"
    End If
    NextRow = NextRow + 1

    ' Reflect gating: if inputs are empty/inactive, we would OutputNoResults
    Dim tagActive As Boolean: tagActive = (Len(Trim$(tagTxt)) >= TAG_SEARCH_MIN_LEN)
    If Not IsArray(rxArr) Or (IsArray(rxArr) And (UBound(rxArr) < LBound(rxArr))) Then
        If Not tagActive Then
            wsDiag.Cells(NextRow, 1).Value = "Gating"
            wsDiag.Cells(NextRow, 2).Value = "Empty inputs; would OutputNoResults"
            Exit Sub
        End If
    End If

    idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        wsDiag.Cells(NextRow, 1).Value = "Visible Rows"
        wsDiag.Cells(NextRow, 2).Value = "None"
        Exit Sub
    Else
        wsDiag.Cells(NextRow, 1).Value = "Visible Rows Count"
        wsDiag.Cells(NextRow, 2).Value = UBound(idxs)
        NextRow = NextRow + 1
    End If

    kept = 0
    For i = 1 To UBound(idxs)
        ri = idxs(i)
        keep = True
        If IsArray(rxArr) And (UBound(rxArr) - LBound(rxArr) + 1) > 0 Then
            descText = SafeCellText(dataLo.DataBodyRange.Cells(ri, searchColIdx).Value)
            For p = LBound(rxArr) To UBound(rxArr)
                If Not rxArr(p).Test(descText) Then keep = False: Exit For
            Next p
        End If
        If keep Then kept = kept + 1
    Next i

    wsDiag.Cells(NextRow, 1).Value = "Matched Rows": wsDiag.Cells(NextRow, 2).Value = kept
End Sub

'============================================================
' DEV DIAGNOSTICS (config/search tracing)
'============================================================
Public Sub RunConfigDiagnostics()
    Dim wsCfg As Worksheet, loCfg As ListObject, r As Range
    Dim diagSheet As Worksheet
    Dim NextRow As Long
    Dim key As String, val As String
    Dim namedRng As Range
    Dim foundHeader As Boolean
    Dim dataLo As ListObject
    Dim colIdx As Long
    Dim ws As Worksheet, l As ListObject

    On Error Resume Next
    Set wsCfg = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = wsCfg.ListObjects("ConfigTable")
    If loCfg Is Nothing Then
        MsgBox "ConfigTable not found on ConfigSheet.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Set diagSheet = Nothing
    On Error Resume Next
    Set diagSheet = ThisWorkbook.Worksheets("ConfigDiagnostics")
    On Error GoTo 0
    If diagSheet Is Nothing Then
        Set diagSheet = ThisWorkbook.Worksheets.Add
        diagSheet.name = "ConfigDiagnostics"
    Else
        diagSheet.Cells.Clear
    End If

    diagSheet.Cells(1, 1).Value = "Config Key"
    diagSheet.Cells(1, 2).Value = "Config Value"
    diagSheet.Cells(1, 3).Value = "Named Range Exists"
    diagSheet.Cells(1, 4).Value = "Header Exists in DataTable"
    NextRow = 2

    Set dataLo = Nothing
    For Each ws In ThisWorkbook.Worksheets
        For Each l In ws.ListObjects
            If StrComp(l.name, DATA_TABLE_NAME, vbTextCompare) = 0 Then
                Set dataLo = l
                Exit For
            End If
        Next l
        If Not dataLo Is Nothing Then Exit For
    Next ws

    For Each r In loCfg.DataBodyRange.Rows
        key = Trim$(CStr(r.Cells(1, 1).Value))
        val = Trim$(CStr(r.Cells(1, 2).Value))

        diagSheet.Cells(NextRow, 1).Value = key
        diagSheet.Cells(NextRow, 2).Value = val

        Set namedRng = Nothing
        On Error Resume Next
        Set namedRng = ThisWorkbook.Names(val).RefersToRange
        On Error GoTo 0
        diagSheet.Cells(NextRow, 3).Value = IIf(namedRng Is Nothing, "No", "Yes")

        foundHeader = False
        If Not dataLo Is Nothing Then
            For colIdx = 1 To dataLo.ListColumns.Count
                If StrComp(CStr(dataLo.HeaderRowRange.Cells(1, colIdx).Value), val, vbTextCompare) = 0 Then
                    foundHeader = True
                    Exit For
                End If
            Next colIdx
        End If
        diagSheet.Cells(NextRow, 4).Value = IIf(foundHeader, "Yes", "No")

        NextRow = NextRow + 1
    Next r

    MsgBox "Diagnostics complete. See 'ConfigDiagnostics' sheet.", vbInformation
End Sub

