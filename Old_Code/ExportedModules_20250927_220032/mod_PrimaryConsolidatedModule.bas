Attribute VB_Name = "mod_PrimaryConsolidatedModule"
'============================================================
' MODE-DRIVEN SEARCH & DISPLAY (Extensible Modes)
'============================================================
' This section implements support for dashboard modes defined in the ModeConfig table.
' Each mode specifies:
'   - FilterFormula: Excel formula (as string) to select rows
'   - OutputColumns: Comma-separated DataTable columns to display
'   - DisplayType: "Table" (default) or "Popup" (custom UserForm)
'   - CustomHandler: (optional) VBA Sub/Function for advanced display
' The active mode is selected via a dashboard dropdown (named range: ModeSelector).
'============================================================
'' Attribute VB_Name = "mod_SmartSearch_FINAL"  ' (commented out for portability)


Option Explicit
'
'============================================================
' EQUIPMENT SEARCH ENGINE (Dev) � Consolidated Single-Module Layout
'------------------------------------------------------------
' Purpose
' - Keep the runtime engine, config access, utilities, and diagnostics
'   together for easy transfer between workbooks during development.
' - Each section is self-contained so code can later be split back into
'   separate modules without renaming or hidden dependencies.
' - Output columns are driven by ConfigTable (Out_Column1..8).
' - Description search is synonym-aware (word-boundary regex) with
'   AND across tokens; Valve/Tag is exact match on "Valve Number".
'------------------------------------------------------------
' Organization (Sections)
' - ENTRYPOINTS (guarded)
' - CORES (PerformSearch / OutputAllVisible)
' - PULSE (slicer/filter change detector)
' - HELPERS (names, tables, output, sorting)
' - REGEX SYNONYMS (description engine)
' - TAG PARSE / RANK (available, not active in filter)
' - LOGGING (local dev)
' - CONFIG ACCESS (from ConfigSheet/ConfigTable)
' - FILTERS / CLEAR (clear inputs & table filters)
' - UTILITIES (misc dev helpers)
' - DEV DIAGNOSTICS (config/search tracing)
'============================================================

'==============================
' SMART SEARCH  PRODUCTION (Description + Tag ID)
'==============================

' Tables & names you already use
Public Const DATA_TABLE_NAME As String = "DataTable"
Public Const MAPPING_TABLE_NAME As String = "tbl_Mapping"

' Tag search behavior
Public Const TAG_SEARCH_MIN_LEN As Long = 3

' Policy: when inputs are empty, do NOT auto-show all results.
' This is enforced in RefreshResults, Safe_PerformSearch, and Safe_OutputAllVisible.
' The Show All button remains available to intentionally display everything.

' Re-entrancy guard
Public gBusy As Boolean

'==============================
' ENTRYPOINTS (guarded)
'==============================

Public Sub RefreshResults()
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo EH

    Dim descTxt As String: descTxt = ReadLeftCell(GetConfigValue("InputCell_DescripSearch"))
    Dim tagTxt  As String: tagTxt = ReadLeftCell(GetConfigValue("InputCell_ValveNumSearch"))

    Dim tagActive As Boolean
    tagActive = (Len(Trim$(tagTxt)) >= TAG_SEARCH_MIN_LEN)

    If Len(Trim$(descTxt)) > 0 Or tagActive Then
        PerformSearch
    Else
        OutputNoResults
    End If

CleanExit:
    gBusy = False
    Exit Sub
EH:
    LogErrorLocal "RefreshResults", Err.Number, Err.DESCRIPTION
    Resume CleanExit
End Sub

Public Sub btn_Search():        Safe_PerformSearch:     End Sub
Public Sub btn_ShowAll():       Safe_OutputAllVisible:  End Sub

Public Sub Safe_PerformSearch()
    ' Gate like RefreshResults so an empty query does NOT show all rows
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo EH

    Dim descTxt As String: descTxt = ReadLeftCell(GetConfigValue("InputCell_DescripSearch"))
    Dim tagTxt  As String: tagTxt = ReadLeftCell(GetConfigValue("InputCell_ValveNumSearch"))
    Dim tagActive As Boolean: tagActive = (Len(Trim$(tagTxt)) >= TAG_SEARCH_MIN_LEN)

    If Len(Trim$(descTxt)) > 0 Or tagActive Then
        PerformSearch
    Else
        OutputNoResults
    End If

CleanExit: gBusy = False: Exit Sub
EH: LogErrorLocal "Safe_PerformSearch", Err.Number, Err.DESCRIPTION: Resume CleanExit
End Sub

Public Sub Safe_OutputAllVisible()
    ' Only show all when explicitly requested AND not violating the empty-inputs rule
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo EH

    Dim descTxt As String: descTxt = ReadLeftCell(GetConfigValue("InputCell_DescripSearch"))
    Dim tagTxt  As String: tagTxt = ReadLeftCell(GetConfigValue("InputCell_ValveNumSearch"))
    Dim tagActive As Boolean: tagActive = (Len(Trim$(tagTxt)) >= TAG_SEARCH_MIN_LEN)

    ' If both inputs are empty/inactive, honor the "no results by default" policy
    If Len(Trim$(descTxt)) = 0 And Not tagActive Then
        OutputNoResults
    Else
        OutputAllVisible
    End If

CleanExit: gBusy = False: Exit Sub
EH: LogErrorLocal "Safe_OutputAllVisible", Err.Number, Err.DESCRIPTION: Resume CleanExit
End Sub

Public Sub PerformSearch()
On Error GoTo EH

Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
    MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
    Exit Sub
End If

Dim resultsStart As Range, statusRng As Range
Set resultsStart = nr(GetConfigValue("ResultsStartCell"))
Set statusRng = nr(GetConfigValue("StatusCell"))
If resultsStart Is Nothing Then
    MsgBox "Named range 'ResultsStartCell' not found.", vbExclamation
    Exit Sub
End If

Dim searchTxt As String: searchTxt = ReadLeftCell(GetConfigValue("InputCell_DescripSearch"))
Dim valveTxt As String: valveTxt = ReadLeftCell(GetConfigValue("InputCell_ValveNumSearch"))

Dim rxArr As Variant
If Len(Trim$(searchTxt)) > 0 Then
    Dim mapLo As ListObject: Set mapLo = lo(MAPPING_TABLE_NAME)
    Dim synIndex As Object: Set synIndex = BuildSynonymIndex(mapLo)
    rxArr = BuildSearchRegexes(searchTxt, synIndex)
Else
    rxArr = Array()
End If
Dim descActive As Boolean: descActive = IsArrayNonEmpty(rxArr)

Dim valveColIdx As Long: valveColIdx = HeaderIndexByText(dataLo, "Valve Number")
Dim valveActive As Boolean: valveActive = (Len(Trim$(valveTxt)) > 0 And valveColIdx > 0)

' Resolve the description column (for matching and default sort)
Dim descColIdx As Long: descColIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)

' Clear old results dynamically once we know our width later

' Collect output column headers from ConfigTable
Dim outCols() As Long, outKeys(1 To 8) As String
Dim i As Long, key As String, headerName As String
ReDim outCols(1 To 8)
outKeys(1) = "Out_Column1": outKeys(2) = "Out_Column2": outKeys(3) = "Out_Column3"
outKeys(4) = "Out_Column4": outKeys(5) = "Out_Column5": outKeys(6) = "Out_Column6"
outKeys(7) = "Out_Column7": outKeys(8) = "Out_Column8"

Dim colCount As Long: colCount = 0
For i = 1 To 8
    key = outKeys(i)
    headerName = GetConfigValue(key)
    If Len(Trim$(headerName)) > 0 Then
        Dim idx As Long: idx = HeaderIndexByText(dataLo, headerName)
        If idx > 0 Then
            colCount = colCount + 1
            outCols(colCount) = idx
        End If
    End If
Next i
If colCount = 0 Then
    MsgBox "No valid output columns found in ConfigTable.", vbExclamation
    Exit Sub
End If
ReDim Preserve outCols(1 To colCount)

' Write headers

Dim hdr() As Variant
ReDim hdr(1 To 1, 1 To colCount)
For i = 1 To colCount
    hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).value)
Next i
resultsStart.Resize(1, colCount).value = hdr

' Now that we know width, clear prior output area accordingly
ClearOldResults resultsStart, colCount

Dim idxs As Variant: idxs = VisibleRowIndexes(dataLo)
If IsEmpty(idxs) Then
    WriteStatus statusRng, "No visible rows (check slicers/filters).", ""
    Exit Sub
End If

Dim maxRows As Long: maxRows = MaxOutputRows()
Dim outArr() As Variant, kept As Long, ri As Long, j As Long
ReDim outArr(1 To UBound(idxs), 1 To colCount)

For i = 1 To UBound(idxs)
    ri = idxs(i)
    Dim keep As Boolean: keep = True

    If descActive Then
        Dim descText As String
        If descColIdx > 0 Then
            descText = SafeCellText(dataLo.DataBodyRange.Cells(ri, descColIdx).value)
        Else
            ' Fallback to first output column if config is missing (best-effort)
            descText = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(1)).value)
        End If
        For j = LBound(rxArr) To UBound(rxArr)
            If Not rxArr(j).Test(descText) Then keep = False: Exit For
        Next j
        If Not keep Then GoTo NextRow
    End If

    If valveActive Then
        Dim valveText As String: valveText = SafeCellText(dataLo.DataBodyRange.Cells(ri, valveColIdx).value)
        If StrComp(valveText, valveTxt, vbTextCompare) <> 0 Then GoTo NextRow
    End If

    If keep Then
        kept = kept + 1
        For j = 1 To colCount
            outArr(kept, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).value)
        Next j
        If maxRows > 0 And kept >= maxRows Then Exit For
    End If
NextRow:
Next i

If kept = 0 Then
    WriteStatus statusRng, "Found 0 results.", "Query: " & Trim$(searchTxt)
    Exit Sub
End If

ReDim finalArr(1 To kept, 1 To colCount)
For i = 1 To kept
    For j = 1 To colCount
        finalArr(i, j) = outArr(i, j)
    Next j
Next i

' Sort results by the description column if it is among the output columns
Dim descOutPos As Long: descOutPos = FindOutputPosForDataColumn(outCols, colCount, descColIdx)
If descOutPos > 0 And kept > 1 Then
    QuickSort2D_N finalArr, 1, kept, descOutPos
End If

resultsStart.Offset(1, 0).Resize(kept, colCount).value = finalArr
WriteStatus statusRng, "Found " & kept & " rows.", "Query: " & Trim$(searchTxt)
Exit Sub
EH:
LogErrorLocal "PerformSearch", Err.Number, Err.DESCRIPTION
End Sub


Public Sub OutputAllVisible()
    On Error GoTo EH

    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = nr(GetConfigValue("ResultsStartCell"))
    Set statusRng = nr(GetConfigValue("StatusCell"))
    If resultsStart Is Nothing Then
        MsgBox "Named range 'ResultsStartCell' not found.", vbExclamation
        Exit Sub
    End If

    ' Build dynamic output columns from Config
    Dim outCols() As Long, outKeys(1 To 8) As String
    Dim i As Long, key As String, headerName As String
    ReDim outCols(1 To 8)
    outKeys(1) = "Out_Column1": outKeys(2) = "Out_Column2": outKeys(3) = "Out_Column3"
    outKeys(4) = "Out_Column4": outKeys(5) = "Out_Column5": outKeys(6) = "Out_Column6"
    outKeys(7) = "Out_Column7": outKeys(8) = "Out_Column8"

    Dim colCount As Long: colCount = 0
    For i = 1 To 8
        key = outKeys(i)
        headerName = GetConfigValue(key)
        If Len(Trim$(headerName)) > 0 Then
            Dim idx As Long: idx = HeaderIndexByText(dataLo, headerName)
            If idx > 0 Then
                colCount = colCount + 1
                outCols(colCount) = idx
            End If
        End If
    Next i
    If colCount = 0 Then
        MsgBox "No valid output columns found in ConfigTable.", vbExclamation
        Exit Sub
    End If
    ReDim Preserve outCols(1 To colCount)

    ' Headers
    Dim hdr() As Variant
    ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount
        hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).value)
    Next i
    resultsStart.Resize(1, colCount).value = hdr

    ' Clear old results with dynamic width
    ClearOldResults resultsStart, colCount

    Dim idxs As Variant: idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        WriteStatus statusRng, "No visible rows (check slicers/filters).", ""
        Exit Sub
    End If

    Dim maxRows As Long: maxRows = MaxOutputRows()
    Dim maxOut As Long
    If maxRows > 0 Then
        If UBound(idxs) < maxRows Then maxOut = UBound(idxs) Else maxOut = maxRows
    Else
        maxOut = UBound(idxs)
    End If

    Dim outArr() As Variant, ri As Long, j As Long
    ReDim outArr(1 To maxOut, 1 To colCount)
    For i = 1 To maxOut
        ri = idxs(i)
        For j = 1 To colCount
            outArr(i, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).value)
        Next j
    Next i

    ' Sort visible output by description column if present among output
    Dim descColIdx As Long: descColIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)
    Dim descOutPos As Long: descOutPos = FindOutputPosForDataColumn(outCols, colCount, descColIdx)
    If descOutPos > 0 And maxOut > 1 Then
        QuickSort2D_N outArr, 1, maxOut, descOutPos
    End If

    resultsStart.Offset(1, 0).Resize(maxOut, colCount).value = outArr
    WriteStatus statusRng, "Displayed " & maxOut & " visible row(s).", ""
    Exit Sub
EH:
    LogErrorLocal "OutputAllVisible", Err.Number, Err.DESCRIPTION
End Sub

Public Sub OutputNoResults()
    On Error GoTo EH

    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = nr(GetConfigValue("ResultsStartCell"))
    Set statusRng = nr(GetConfigValue("StatusCell"))
    If resultsStart Is Nothing Then
        MsgBox "Named range 'ResultsStartCell' not found.", vbExclamation
        Exit Sub
    End If

    ' Build dynamic output columns from Config for header row
    Dim outCols() As Long, outKeys(1 To 8) As String
    Dim i As Long, key As String, headerName As String
    ReDim outCols(1 To 8)
    outKeys(1) = "Out_Column1": outKeys(2) = "Out_Column2": outKeys(3) = "Out_Column3"
    outKeys(4) = "Out_Column4": outKeys(5) = "Out_Column5": outKeys(6) = "Out_Column6"
    outKeys(7) = "Out_Column7": outKeys(8) = "Out_Column8"

    Dim colCount As Long: colCount = 0
    For i = 1 To 8
        key = outKeys(i)
        headerName = GetConfigValue(key)
        If Len(Trim$(headerName)) > 0 Then
            Dim idx As Long: idx = HeaderIndexByText(dataLo, headerName)
            If idx > 0 Then
                colCount = colCount + 1
                outCols(colCount) = idx
            End If
        End If
    Next i
    If colCount = 0 Then
        MsgBox "No valid output columns found in ConfigTable.", vbExclamation
        Exit Sub
    End If
    ReDim Preserve outCols(1 To colCount)

    ' Write headers only
    Dim hdr() As Variant
    ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount
        hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).value)
    Next i
    resultsStart.Resize(1, colCount).value = hdr

    ' Clear old results with dynamic width (show no data rows)
    ClearOldResults resultsStart, colCount
    
    WriteStatus statusRng, "Enter search criteria to display results.", ""
    Exit Sub
EH:
    LogErrorLocal "OutputNoResults", Err.Number, Err.DESCRIPTION
End Sub

'==============================
' PULSE (slicers/filters change detector)
'==============================

Public Sub EnsurePulseCell(Optional recalcNow As Boolean = False)
Dim dash As Worksheet, dataLo As ListObject, pulseCell As Range
Dim searchColIdx As Long, addr As String

Set dash = SheetByName(GetConfigValue("DASHBOARD_SHEET"))
If dash Is Nothing Then Exit Sub

Set dataLo = lo(DATA_TABLE_NAME)
If dataLo Is Nothing Then Exit Sub
If dataLo.ListColumns.count = 0 Or dataLo.DataBodyRange Is Nothing Then Exit Sub

searchColIdx = GetColumnIndex("DataTable_TagID", dataLo)
If searchColIdx < 1 Or searchColIdx > dataLo.ListColumns.count Then searchColIdx = 1

Set pulseCell = dash.Range("AA1")
dash.Columns("AA").Hidden = True

addr = dataLo.ListColumns(searchColIdx).DataBodyRange.Address(True, True, xlA1, True)
Application.EnableEvents = False
pulseCell.formula = "=SUBTOTAL(103," & addr & ")"
NameOrUpdate "SlicerPulseCell", "=" & pulseCell.Address(True, True, xlA1, True)
NameOrUpdate "SlicerPulse", "=" & pulseCell.Address(True, True, xlA1, True)
Application.EnableEvents = True

If recalcNow Then Application.Calculate
End Sub

'==============================
' HELPERS (generic)
'==============================

' Visible rows as 1-based indices relative to DataBodyRange
Public Function VisibleRowIndexes(ByVal dataLo As ListObject) As Variant
    Dim firstCol As Range, vis As Range, c As Range
    Dim idx() As Long, n As Long, startRow As Long
    On Error Resume Next
    Set firstCol = dataLo.ListColumns(1).DataBodyRange
    If firstCol Is Nothing Then Exit Function
    startRow = firstCol.row
    Set vis = firstCol.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If vis Is Nothing Then Exit Function
    ReDim idx(1 To vis.count)
    For Each c In vis.Cells
        n = n + 1
        idx(n) = c.row - startRow + 1
    Next c
    If n = 0 Then Exit Function
    ReDim Preserve idx(1 To n)
    VisibleRowIndexes = idx
End Function

' Validate indices exist
Public Function IndicesValid(ByVal dataLo As ListObject, _
ByVal idx1 As Long, ByVal idx2 As Long, ByVal idx3 As Long) As Boolean
    Dim n As Long: n = dataLo.ListColumns.count
    If idx1 < 1 Or idx1 > n Or idx2 < 1 Or idx2 > n Or idx3 < 1 Or idx3 > n Then
        MsgBox "Config column index/indices exceed the DataTable column count.", vbExclamation
        IndicesValid = False
    Else
        IndicesValid = True
    End If
End Function

Public Function lo(ByVal name As String) As ListObject
    Dim ws As Worksheet, l As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each l In ws.ListObjects
            If StrComp(l.name, name, vbTextCompare) = 0 Then Set lo = l: Exit Function
        Next l
    Next ws
End Function

Public Function nr(ByVal nm As String) As Range
    On Error Resume Next
    Set nr = ThisWorkbook.names(nm).RefersToRange
End Function

Public Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
End Function





Public Function ReadLeftCell(ByVal nm As String) As String
    On Error Resume Next
    Dim rng As Range: Set rng = nr(nm)
    If Not rng Is Nothing Then ReadLeftCell = CStr(rng.Cells(1, 1).value)
End Function

Public Function SafeCellText(ByVal v As Variant) As String
    If IsError(v) Then SafeCellText = "" Else SafeCellText = CStr(v)
End Function

Public Function HeaderIndexByText(ByVal dataLo As ListObject, ByVal headerText As String) As Long
    Dim i As Long
    If dataLo Is Nothing Then Exit Function
    If Len(Trim$(headerText)) = 0 Then Exit Function
    For i = 1 To dataLo.ListColumns.count
        If StrComp(CStr(dataLo.HeaderRowRange.Cells(1, i).value), headerText, vbTextCompare) = 0 Then
            HeaderIndexByText = i: Exit Function
        End If
    Next i
End Function

Public Sub WriteHeaders(ByVal startCell As Range, ByVal dataLo As ListObject, _
                         ByVal codeIdx As Long, ByVal sortIdx As Long, ByVal searchIdx As Long)
    ' Legacy 3-col header writer kept for compatibility (unused in new dynamic flow)
    Dim hdr(1 To 1, 1 To 3) As Variant
    hdr(1, 1) = CStr(dataLo.HeaderRowRange.Cells(1, codeIdx).value)
    hdr(1, 2) = CStr(dataLo.HeaderRowRange.Cells(1, sortIdx).value)
    hdr(1, 3) = CStr(dataLo.HeaderRowRange.Cells(1, searchIdx).value)
    startCell.Resize(1, 3).value = hdr
End Sub

Public Sub ClearOldResults(ByVal startCell As Range, Optional ByVal colCount As Long = 3)
    startCell.Offset(1, 0).Resize(100000, colCount).ClearContents
End Sub

Public Sub WriteStatus(ByVal statusRng As Range, ByVal line1 As String, ByVal line2 As String)
    On Error Resume Next
    If Not statusRng Is Nothing Then
        statusRng.Cells(1, 1).value = line1
        If statusRng.Rows.count >= 2 Then statusRng.Cells(2, 1).value = line2
    End If
End Sub

Public Function MaxOutputRows() As Long
Dim v As String: v = GetConfigValue("MAX_OUTPUT_ROWS")
MaxOutputRows = CLngSafe(v)
End Function

Public Sub NameOrUpdate(ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    If Not ThisWorkbook.names(nm) Is Nothing Then
        ThisWorkbook.names(nm).refersTo = refersTo
    Else
        ThisWorkbook.names.Add name:=nm, refersTo:=refersTo
    End If
End Sub

' 2D quicksort by column sortCol (string comparison)
Public Sub QuickSort2D(ByRef arr As Variant, ByVal loIdx As Long, ByVal hiIdx As Long, ByVal sortCol As Long)
    ' Legacy 3-col sorter retained for compatibility (not used in new flow)
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim t1 As Variant, t2 As Variant, t3 As Variant
    i = loIdx: j = hiIdx
    pivot = arr((loIdx + hiIdx) \ 2, sortCol)
    Do While i <= j
        Do While CStr(arr(i, sortCol)) < CStr(pivot): i = i + 1: Loop
        Do While CStr(arr(j, sortCol)) > CStr(pivot): j = j - 1: Loop
        If i <= j Then
            t1 = arr(i, 1): t2 = arr(i, 2): t3 = arr(i, 3)
            arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2): arr(i, 3) = arr(j, 3)
            arr(j, 1) = t1: arr(j, 2) = t2: arr(j, 3) = t3
            i = i + 1: j = j - 1
        End If
    Loop
    If loIdx < j Then QuickSort2D arr, loIdx, j, sortCol
    If i < hiIdx Then QuickSort2D arr, i, hiIdx, sortCol
End Sub

Public Sub QuickSort2D_N(ByRef arr As Variant, ByVal loIdx As Long, ByVal hiIdx As Long, ByVal sortCol As Long)
    ' Generic quicksort for 2D arrays [rows, cols]; swaps all columns
    Dim i As Long, j As Long, k As Long, colN As Long
    Dim pivot As Variant, tmp As Variant
    If hiIdx <= loIdx Then Exit Sub
    colN = UBound(arr, 2)
    i = loIdx: j = hiIdx
    pivot = arr((loIdx + hiIdx) \ 2, sortCol)
    Do While i <= j
        Do While CStr(arr(i, sortCol)) < CStr(pivot): i = i + 1: Loop
        Do While CStr(arr(j, sortCol)) > CStr(pivot): j = j - 1: Loop
        If i <= j Then
            For k = 1 To colN
                tmp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = tmp
            Next k
            i = i + 1: j = j - 1
        End If
    Loop
    If loIdx < j Then QuickSort2D_N arr, loIdx, j, sortCol
    If i < hiIdx Then QuickSort2D_N arr, i, hiIdx, sortCol
End Sub

Private Function FindOutputPosForDataColumn(ByRef outCols() As Long, ByVal colCount As Long, ByVal dataColIdx As Long) As Long
    Dim p As Long
    If dataColIdx <= 0 Then Exit Function
    For p = 1 To colCount
        If outCols(p) = dataColIdx Then FindOutputPosForDataColumn = p: Exit Function
    Next p
End Function

Public Function IsArrayNonEmpty(ByVal v As Variant) As Boolean
    On Error GoTo Nope
    If IsArray(v) Then
        If (UBound(v) - LBound(v) + 1) > 0 Then IsArrayNonEmpty = True
    End If
    Exit Function
Nope:
End Function

'==============================
' REGEX SYNONYM ENGINE (Description)
'==============================

Public Function EscapeRegex(ByVal s As String) As String
    Dim specials As Variant, ch As Variant
    specials = Array("\", ".", "+", "*", "?", "|", "{", "}", "[", "]", "(", ")", "^", "$")
    For Each ch In specials
        If InStr(s, ch) > 0 Then s = Replace(s, ch, "\" & ch)
    Next ch
    EscapeRegex = s
End Function

' key (lowercase RawTerm or StandardTerm) -> array of synonyms (inc. the StandardTerm)
Public Function BuildSynonymIndex(ByVal mapLo As ListObject) As Object
    Dim syn As Object: Set syn = CreateObject("Scripting.Dictionary")
    syn.CompareMode = vbTextCompare
    If mapLo Is Nothing Or mapLo.DataBodyRange Is Nothing Then Set BuildSynonymIndex = syn: Exit Function

    Dim groupDict As Object: Set groupDict = CreateObject("Scripting.Dictionary")
    groupDict.CompareMode = vbTextCompare

    Dim r As Range, raw As String, std As String
    For Each r In mapLo.DataBodyRange.Rows
        raw = LCase(Trim$(CStr(r.Cells(1, 1).value)))
        std = LCase(Trim$(CStr(r.Cells(1, 2).value)))
        If Len(std) = 0 Then std = raw
        If Len(raw) > 0 Then
            If Not groupDict.exists(std) Then groupDict.Add std, CreateObject("Scripting.Dictionary")
            groupDict(std).CompareMode = vbTextCompare
            groupDict(std)(std) = True
            groupDict(std)(raw) = True
        End If
    Next r

    Dim k As Variant, d As Object, arr() As String, i As Long, t As Variant
    For Each k In groupDict.Keys
        Set d = groupDict(k)
        ReDim arr(0 To d.count - 1): i = 0
        For Each t In d.Keys
            arr(i) = CStr(t): i = i + 1
        Next t
        syn(k) = arr
        For Each t In d.Keys
            syn(CStr(t)) = arr
        Next t
    Next k
    Set BuildSynonymIndex = syn
End Function

Public Function BuildSearchRegexes(ByVal searchText As String, ByVal synIndex As Object) As Variant
    Dim tokens As Variant, rxArr() As Object, rx As Object
    Dim i As Long, j As Long
    Dim term As String, alts As Variant, patt As String

    searchText = Trim$(searchText)
    If Len(searchText) = 0 Then BuildSearchRegexes = Array(): Exit Function

    tokens = Split(searchText, " ")
    ReDim rxArr(0 To 0)
    Dim cnt As Long: cnt = -1

    For i = LBound(tokens) To UBound(tokens)
        term = Trim$(CStr(tokens(i)))
        If Len(term) > 0 Then
            term = LCase(term)
            If synIndex Is Nothing Or Not synIndex.exists(term) Then
                alts = Array(term)
            Else
                alts = synIndex(term)
            End If
            For j = LBound(alts) To UBound(alts)
                alts(j) = EscapeRegex(CStr(alts(j)))
            Next j
            If UBound(alts) > LBound(alts) Then
                patt = "\b(" & Join(alts, "|") & ")\b"
            Else
                patt = "\b" & alts(LBound(alts)) & "\b"
            End If

            Set rx = CreateObject("VBScript.RegExp")
            rx.Global = False
            rx.IgnoreCase = True
            rx.pattern = patt

            cnt = cnt + 1
            If cnt = 0 Then
                ReDim rxArr(0 To 0)
            Else
                ReDim Preserve rxArr(0 To cnt)
            End If
            Set rxArr(cnt) = rx
        End If
    Next i

    If cnt >= 0 Then BuildSearchRegexes = rxArr Else BuildSearchRegexes = Array()
End Function

'==============================


' Load allowed code sets (flexible: only constants change if you move things)
Public Function CodeSet(ByVal sheetName As String, ByVal headerName As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    On Error GoTo Clean
    Dim ws As Worksheet: Set ws = SheetByName(sheetName)
    If ws Is Nothing Then GoTo Clean
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then GoTo Clean
    Dim hdrRow As Range: Set hdrRow = ur.Rows(1)
    Dim c As Range, col As Long: col = 0
    For Each c In hdrRow.Cells
        If StrComp(Trim$(CStr(c.value)), headerName, vbTextCompare) = 0 Then col = c.Column: Exit For
    Next c
    If col = 0 Then GoTo Clean
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).row
    Dim r As Long, v As String
    For r = 2 To lastRow
        v = Trim$(UCase$(CStr(ws.Cells(r, col).value)))
        If Len(v) > 0 Then If Not d.exists(v) Then d.Add v, True
    Next r
Clean:
    Set CodeSet = d
End Function

' Parse user Tag ID query into sys/obj/num (only sys &/or num are required)
Public Sub ParseTagQuery(ByVal q As String, ByVal sysSet As Object, ByVal objSet As Object, _
                          ByRef qSys As String, ByRef qObj As String, ByRef qNum As String)
    qSys = "": qObj = "": qNum = ""
    Dim toks As Variant, i As Long, t As String
    q = Trim$(q)
    If Len(q) = 0 Then Exit Sub

    ' Split and inspect tokens
    toks = Split(q, " ")
    For i = LBound(toks) To UBound(toks)
        t = UCase$(Trim$(CStr(toks(i))))
        If Len(t) = 0 Then GoTo NextTok
        If HasDigit(t) Then
            qNum = LastDigits(t)                      ' last contiguous run of digits in this token
        ElseIf IsLetters(t) And (Len(t) >= 2 And Len(t) <= 3) Then
            If sysSet.exists(t) Then
                qSys = t
            ElseIf Len(t) = 2 And objSet.exists(t) Then
                qObj = t
            End If
        End If
NextTok:
    Next i

    ' Fallback once: if no digits found per-token, try entire string
    If qNum = "" Then qNum = LastDigits(UCase$(q))
End Sub

' Compute rank for a row Tag ID against the query.
' Returns:
'   0 = best (Sys+Num and Obj match if provided)
'   1 = Sys+Num match but Obj differs/omitted
'   2 = Num-only match (or Sys-only when Num not provided)
'  -1 = no match / exclude (or non-conforming when Tag filter active)
Public Function TagMatchRank(ByVal qSys As String, ByVal qObj As String, ByVal qNum As String, _
                              ByVal rowTagText As String) As Long
    Dim rSys As String, rObj As String, rNum As String
    rSys = "": rObj = "": rNum = ""
    ParseRowTag rowTagText, rSys, rObj, rNum

    ' If row is nonconforming and a number/system was asked for, exclude
    If (qSys <> "" Or qNum <> "") And (rSys = "" And rNum = "") Then
        TagMatchRank = -1
        Exit Function
    End If

    ' System exact if provided
    If qSys <> "" Then
        If rSys = "" Or StrComp(rSys, qSys, vbTextCompare) <> 0 Then TagMatchRank = -1: Exit Function
    End If

    ' Number ends-with if provided
    If qNum <> "" Then
        If rNum = "" Then TagMatchRank = -1: Exit Function
        If Right$(rNum, Len(qNum)) <> qNum Then TagMatchRank = -1: Exit Function
    End If

    ' Ranking
    If qSys <> "" And qNum <> "" Then
        If qObj <> "" Then
            If rObj <> "" And StrComp(rObj, qObj, vbTextCompare) = 0 Then
                TagMatchRank = 0
            Else
                TagMatchRank = 1
            End If
        Else
            TagMatchRank = 0
        End If
    ElseIf qSys <> "" And qNum = "" Then
        TagMatchRank = 2
    ElseIf qSys = "" And qNum <> "" Then
        TagMatchRank = 2
    Else
        TagMatchRank = -1
    End If
End Function

' Extract row sys/obj/number robustly; accept tokens like "C3003*P".
' Prefer digits among the first up-to-three tokens (Sys / Obj / Tag token).
Public Sub ParseRowTag(ByVal s As String, ByRef rSys As String, ByRef rObj As String, ByRef rNum As String)
    rSys = "": rObj = "": rNum = ""
    Dim u As String: u = UCase$(Trim$(CStr(s)))
    If Len(u) = 0 Then Exit Sub

    Dim toks As Variant: toks = Split(u, " ")
    Dim maxTok As Long: maxTok = WorksheetFunction.Min(UBound(toks), 2)

    ' Sys = first token if 2-3 letters
    If UBound(toks) >= 0 Then
        If IsLetters(toks(0)) And (Len(toks(0)) >= 2 And Len(toks(0)) <= 3) Then rSys = toks(0)
    End If
    ' Obj = second token if exactly 2 letters
    If UBound(toks) >= 1 Then
        If IsLetters(toks(1)) And Len(toks(1)) = 2 Then rObj = toks(1)
    End If
    ' Number: prefer digits found within the first 3 tokens (handles C3003*P)
    Dim i As Long, cand As String
    For i = 0 To maxTok
        cand = LastDigits(toks(i))
        If Len(cand) > 0 Then rNum = cand     ' keep updating to the last seen among first 3 tokens
    Next i
    ' If still blank, last fallback = last digits anywhere (rare but safe)
    If rNum = "" Then rNum = LastDigits(u)
End Sub

Public Function LastDigits(ByVal s As String) As String
    Dim rx As Object, ms As Object, m As Object
    Dim v As String
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.pattern = "\d+"
    Set ms = rx.Execute(s)
    If ms Is Nothing Or ms.count = 0 Then Exit Function
    For Each m In ms
        v = m.value
    Next m
    LastDigits = v
End Function

Public Function HasDigit(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then HasDigit = True: Exit Function
    Next i
End Function

Public Function IsLetters(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "A" Or ch > "Z" Then IsLetters = False: Exit Function
    Next i
    IsLetters = True
End Function

' Rank-aware sort key helpers
Public Function MakeSortKey(ByVal rankVal As Long, ByVal sortText As String) As String
    MakeSortKey = Right$("00" & CStr(rankVal), 2) & "|" & sortText
End Function

Public Function StripSortKey(ByVal s As String) As String
    Dim p As Long: p = InStr(1, s, "|", vbBinaryCompare)
    If p > 0 Then StripSortKey = Mid$(s, p + 1) Else StripSortKey = s
End Function

'==============================
' LOGGING (local)
'==============================

Public Sub LogErrorLocal(ByVal procName As String, ByVal errNum As Long, ByVal errDesc As String)
    Dim logSheet As Worksheet
    Dim NextRow As Long
    Dim searchText As String, tagText As String

    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("SearchErrorLog")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add
        logSheet.name = "SearchErrorLog"
        logSheet.Cells(1, 1).value = "Timestamp"
        logSheet.Cells(1, 2).value = "Procedure"
        logSheet.Cells(1, 3).value = "Error Number"
        logSheet.Cells(1, 4).value = "Error Description"
        logSheet.Cells(1, 5).value = "SearchBox"
        logSheet.Cells(1, 6).value = "TagID"
    End If
    On Error GoTo 0

    NextRow = logSheet.Cells(logSheet.Rows.count, 1).End(xlUp).row + 1
    searchText = SafeReadNameLocal("SearchBox")
    tagText = SafeReadNameLocal("ValveNumSearchBox")

    logSheet.Cells(NextRow, 1).value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    logSheet.Cells(NextRow, 2).value = procName
    logSheet.Cells(NextRow, 3).value = errNum
    logSheet.Cells(NextRow, 4).value = errDesc
    logSheet.Cells(NextRow, 5).value = searchText
    logSheet.Cells(NextRow, 6).value = tagText

    MsgBox "Unexpected error in " & procName & "." & vbCrLf & "Logged to worksheet: SearchErrorLog", vbExclamation
End Sub


Public Function SafeReadNameLocal(ByVal nm As String) As String
On Error Resume Next
Dim r As Range: Set r = nr(nm)
If Not r Is Nothing Then SafeReadNameLocal = CStr(r.Cells(1, 1).value)
End Function


'============================================================
' CONFIG ACCESS (from ConfigSheet/ConfigTable)
'============================================================
Public Function GetConfigValue(ByVal key As String) As String
    Dim ws As Worksheet, loCfg As ListObject, r As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Or loCfg.DataBodyRange Is Nothing Then Exit Function
    On Error GoTo 0
    For Each r In loCfg.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).value), key, vbTextCompare) = 0 Then
            GetConfigValue = CStr(r.Cells(1, 2).value)
            Exit Function
        End If
    Next r
End Function

Public Function CLngSafe(ByVal s As String) As Long
    If Len(Trim$(s)) = 0 Then
        CLngSafe = 0
    Else
        CLngSafe = CLng(val(s))
    End If
End Function

Public Function GetColumnIndex(ByVal configKey As String, ByVal dataLo As ListObject) As Long
    Dim headerName As String
    headerName = GetConfigValue(configKey)
    If Len(Trim$(headerName)) = 0 Then Exit Function
    GetColumnIndex = HeaderIndexByText(dataLo, headerName)
End Function

'============================================================
' FILTERS / CLEAR (clear inputs & table filters)
'============================================================
Public Sub Clearfilters()
    Dim ws As Worksheet
    Dim tbl As ListObject
    ClearSearchBoxes
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.name = DATA_TABLE_NAME Then
                If Not tbl.AutoFilter Is Nothing Then
                    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
                End If
                ' Refresh results to show no results after clearing all filters and search boxes
                RefreshResults
                Exit Sub
            End If
        Next tbl
    Next ws
    MsgBox "Table '" & DATA_TABLE_NAME & "' not found.", vbExclamation
End Sub

Public Sub ClearSearchBoxes()
    On Error Resume Next
    Dim nm As String
    ' Clear config-based search boxes
    nm = GetConfigValue("InputCell_DescripSearch"): If Len(nm) > 0 Then ThisWorkbook.names(nm).RefersToRange.ClearContents
    nm = GetConfigValue("InputCell_ValveNumSearch"): If Len(nm) > 0 Then ThisWorkbook.names(nm).RefersToRange.ClearContents
    
    ' Clear additional search inputs if they exist
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("SearchResults")
    If Not ws Is Nothing Then
        If Not ws.Range("SearchInput") Is Nothing Then ws.Range("SearchInput").ClearContents
        If Not ws.Range("LocationFilter") Is Nothing Then ws.Range("LocationFilter").ClearContents
    End If
    
    ' Clear any other named search boxes
    If Not ThisWorkbook.names("SearchBox") Is Nothing Then ThisWorkbook.names("SearchBox").RefersToRange.ClearContents
    If Not ThisWorkbook.names("ValveNumSearchBox") Is Nothing Then ThisWorkbook.names("ValveNumSearchBox").RefersToRange.ClearContents
    
    On Error GoTo 0
End Sub

'============================================================
' UTILITIES (misc dev helpers)
'============================================================
Public Sub InsertStaticDateTime()
    With ActiveCell
        .value = Now
        .NumberFormat = "mm/dd/yyyy hh:mm"
    End With
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

    diagSheet.Cells(1, 1).value = "Config Key"
    diagSheet.Cells(1, 2).value = "Config Value"
    diagSheet.Cells(1, 3).value = "Named Range Exists"
    diagSheet.Cells(1, 4).value = "Header Exists in DataTable"
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
        key = Trim$(CStr(r.Cells(1, 1).value))
        val = Trim$(CStr(r.Cells(1, 2).value))

        diagSheet.Cells(NextRow, 1).value = key
        diagSheet.Cells(NextRow, 2).value = val

        Set namedRng = Nothing
        On Error Resume Next
        Set namedRng = ThisWorkbook.names(val).RefersToRange
        On Error GoTo 0
        diagSheet.Cells(NextRow, 3).value = IIf(namedRng Is Nothing, "No", "Yes")

        foundHeader = False
        If Not dataLo Is Nothing Then
            For colIdx = 1 To dataLo.ListColumns.count
                If StrComp(CStr(dataLo.HeaderRowRange.Cells(1, colIdx).value), val, vbTextCompare) = 0 Then
                    foundHeader = True
                    Exit For
                End If
            Next colIdx
        End If
        diagSheet.Cells(NextRow, 4).value = IIf(foundHeader, "Yes", "No")

        NextRow = NextRow + 1
    Next r

    MsgBox "Diagnostics complete. See 'ConfigDiagnostics' sheet.", vbInformation
End Sub

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
    wsDiag.Cells(NextRow, 1).value = "Step"
    wsDiag.Cells(NextRow, 2).value = "Detail"
    NextRow = NextRow + 1

    searchTxt = ReadLeftCell(GetConfigValue("InputCell_DescripSearch"))
    tagTxt = ReadLeftCell(GetConfigValue("InputCell_ValveNumSearch"))

    wsDiag.Cells(NextRow, 1).value = "Search Text": wsDiag.Cells(NextRow, 2).value = searchTxt: NextRow = NextRow + 1
    wsDiag.Cells(NextRow, 1).value = "Tag Text": wsDiag.Cells(NextRow, 2).value = tagTxt: NextRow = NextRow + 1

    Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Then
        wsDiag.Cells(NextRow, 1).value = "Error"
        wsDiag.Cells(NextRow, 2).value = "DataTable not found"
        Exit Sub
    End If

    searchColIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)
    wsDiag.Cells(NextRow, 1).value = "Description Column Index": wsDiag.Cells(NextRow, 2).value = searchColIdx: NextRow = NextRow + 1

    Set mapLo = lo(MAPPING_TABLE_NAME)
    Set synIndex = BuildSynonymIndex(mapLo)
    If Len(Trim$(searchTxt)) > 0 Then
        rxArr = BuildSearchRegexes(searchTxt, synIndex)
    Else
        rxArr = Array()
    End If

    wsDiag.Cells(NextRow, 1).value = "Regex Array Count"
    If IsArray(rxArr) Then
        wsDiag.Cells(NextRow, 2).value = UBound(rxArr) - LBound(rxArr) + 1
    Else
        wsDiag.Cells(NextRow, 2).value = "Not an array"
    End If
    NextRow = NextRow + 1

    ' Reflect gating: if inputs are empty/inactive, we would OutputNoResults
    Dim tagActive As Boolean: tagActive = (Len(Trim$(tagTxt)) >= TAG_SEARCH_MIN_LEN)
    If Not IsArray(rxArr) Or (IsArray(rxArr) And (UBound(rxArr) < LBound(rxArr))) Then
        If Not tagActive Then
            wsDiag.Cells(NextRow, 1).value = "Gating"
            wsDiag.Cells(NextRow, 2).value = "Empty inputs; would OutputNoResults"
            Exit Sub
        End If
    End If

    idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        wsDiag.Cells(NextRow, 1).value = "Visible Rows"
        wsDiag.Cells(NextRow, 2).value = "None"
        Exit Sub
    Else
        wsDiag.Cells(NextRow, 1).value = "Visible Rows Count"
        wsDiag.Cells(NextRow, 2).value = UBound(idxs)
        NextRow = NextRow + 1
    End If

    kept = 0
    For i = 1 To UBound(idxs)
        ri = idxs(i)
        keep = True
        If IsArray(rxArr) And (UBound(rxArr) - LBound(rxArr) + 1) > 0 Then
            descText = SafeCellText(dataLo.DataBodyRange.Cells(ri, searchColIdx).value)
            For p = LBound(rxArr) To UBound(rxArr)
                If Not rxArr(p).Test(descText) Then keep = False: Exit For
            Next p
        End If
        If keep Then kept = kept + 1
    Next i

    wsDiag.Cells(NextRow, 1).value = "Matched Rows": wsDiag.Cells(NextRow, 2).value = kept
End Sub










' =====================================================================================
' Handler: Search_SootblowerLocation
' Purpose: Filters equipment records by search term and sootblower location
' Triggered by: ModeDrivenSearch when SearchMode = "Sootblower Location"
' =====================================================================================

Public Sub Search_SootblowerLocation()
    Dim wsData As Worksheet, wsResults As Worksheet
    Dim tblData As ListObject
    Dim searchTerm As String, locationFilter As String
    Dim r As ListRow, matchFound As Boolean
    Dim resultRow As Long

    ' === Setup ===
    Set wsData = ThisWorkbook.Sheets("EquipmentData")
    Set wsResults = ThisWorkbook.Sheets("SearchResults")
    Set tblData = wsData.ListObjects("tbl_Equipment")

    searchTerm = Trim(wsResults.Range("SearchInput").value)
    locationFilter = Trim(wsResults.Range("LocationFilter").value)

    wsResults.Range("ResultsTable").ClearContents
    resultRow = wsResults.Range("ResultsTable").row

    ' === Loop through data table ===
    For Each r In tblData.ListRows
        matchFound = False

        ' Match search term in Tag or Description
        If InStr(1, r.Range(tblData.ListColumns("Tag").Index).value, searchTerm, vbTextCompare) > 0 _
        Or InStr(1, r.Range(tblData.ListColumns("Description").Index).value, searchTerm, vbTextCompare) > 0 Then
            matchFound = True
        End If

        ' Match location filter
        If matchFound Then
            If locationFilter = "" Or _
               StrComp(r.Range(tblData.ListColumns("Location").Index).value, locationFilter, vbTextCompare) = 0 Then
                ' Copy matching row to results
                r.Range.Copy Destination:=wsResults.Cells(resultRow, 1)
                resultRow = resultRow + 1
            End If
        End If
    Next r

    ' === Finalize ===
    If resultRow = wsResults.Range("ResultsTable").row Then
        MsgBox "No matching records found for Sootblower Location mode.", vbInformation
    End If
End Sub




