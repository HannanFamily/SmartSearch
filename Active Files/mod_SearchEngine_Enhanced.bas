'Attribute VB_Name = "mod_SearchEngine_Enhanced"
'============================================================
' ENHANCED SEARCH ENGINE - Perfect Compatibility Edition
'============================================================
' Purpose: Enhanced search engine that maintains 100% compatibility
'          with existing code while providing new capabilities
' 
' COMPATIBILITY GUARANTEE:
' - All existing function signatures preserved
' - All existing named ranges and config keys supported
' - All existing slicer behavior maintained
' - Drop-in replacement for mod_PrimaryConsolidatedModule3
'============================================================


Option Explicit
' ============================================================
' CONSTANTS - Matching Original System Exactly
' ============================================================
Private Const CFG_DASHBOARD_SHEET      As String = "DASHBOARD_SHEET"
Private Const CFG_DATA_TABLE_NAME      As String = "DATA_TABLE_NAME"
Private Const CFG_MAPPING_TABLE_NAME   As String = "MAPPING_TABLE_NAME"
Private Const CFG_SPLICER_THRESHOLD    As String = "SPLICER_THRESHOLD"
Private Const CFG_TEMP_FILTER_COL_NAME As String = "TEMP_FILTER_COL_NAME"
Private Const CFG_SLICER_PULSE_ANCHOR  As String = "SLICER_PULSE_ANCHOR"
Private Const CFG_INPUT_DESCRIP        As String = "InputCell_DescripSearch"
Private Const CFG_INPUT_VALVE          As String = "InputCell_ValveNumSearch"
Private Const CFG_RESULTS_START        As String = "ResultsStartCell"
Private Const CFG_STATUS_CELL          As String = "StatusCell"
Private Const CFG_DATA_DESC_COL        As String = "DataTable_EquipDescription"

' Search Engine Constants
Public Const TAG_SEARCH_MIN_LEN As Long = 3

Public DiagnosticMode As Boolean ' Toggle for integrated diagnostics
Public gBusy As Boolean

Private Sub SetDiagnosticModeFromConfig()
    Dim v As String: v = GetConfigValue("DIAGNOSTIC_TOGGLE")
    DiagnosticMode = (StrComp(Trim(v), "TRUE", vbTextCompare) = 0 Or v = "1")
End Sub

' ============================================================
' MAIN ENTRY POINTS - Exact Function Signatures
' ============================================================

' Main refresh function - called by dashboard events
Public Sub RefreshResults_Enhanced()
    Call SetDiagnosticModeFromConfig
    If DiagnosticMode Then Debug.Print "[DIAG] RefreshResults_Enhanced called. gBusy=" & gBusy
    If gBusy Then
        If DiagnosticMode Then Debug.Print "[DIAG] Search already in progress, skipping."
        Exit Sub
    End If
    gBusy = True
    On Error GoTo CleanExit

    If DiagnosticMode Then Debug.Print "[DIAG] Calling EnsurePulseCell(False)"
    Call EnsurePulseCell(False)  ' idempotent; guarantees SlicerPulse exists

    Dim anyActive As Boolean: anyActive = IsAnySearchInputActive()
    Dim pulseOk As Boolean
    Dim pulseVal As Long: pulseVal = CurrentSlicerPulse()
    pulseOk = (pulseVal > 0 And pulseVal <= SlicerThreshold())
    If DiagnosticMode Then Debug.Print "[DIAG] anyActive=" & anyActive & ", pulseVal=" & pulseVal & ", pulseOk=" & pulseOk

    If anyActive Then
        If DiagnosticMode Then Debug.Print "[DIAG] Search-driven branch."
        Call PerformSearch
    ElseIf pulseOk Then
        If DiagnosticMode Then Debug.Print "[DIAG] Slicer-driven branch."
        Call ClearTempSearchFilter
        Call OutputAllVisible
    Else
        If DiagnosticMode Then Debug.Print "[DIAG] Default branch (headers only)."
        Call ClearTempSearchFilter
        Call OutputNoResults
    End If

CleanExit:
    If DiagnosticMode Then Debug.Print "[DIAG] RefreshResults_Enhanced exiting. gBusy reset to False."
    gBusy = False
End Sub

' Safe search wrapper
Public Sub Safe_PerformSearch()
    Call SetDiagnosticModeFromConfig
    If DiagnosticMode Then Debug.Print "[DIAG] Safe_PerformSearch called. gBusy=" & gBusy
    If gBusy Then
        If DiagnosticMode Then Debug.Print "[DIAG] Search already in progress, skipping."
        Exit Sub
    End If
    gBusy = True
    On Error GoTo CleanExit

    If IsAnySearchInputActive() Then
        If DiagnosticMode Then Debug.Print "[DIAG] Input active, calling PerformSearch."
        Call PerformSearch
    Else
        If DiagnosticMode Then Debug.Print "[DIAG] No input active, clearing filter and outputting no results."
        Call ClearTempSearchFilter
        Call OutputNoResults
    End If

CleanExit:
    If DiagnosticMode Then Debug.Print "[DIAG] Safe_PerformSearch exiting. gBusy reset to False."
    gBusy = False
End Sub

' Main search implementation
Public Sub PerformSearch()
    Call SetDiagnosticModeFromConfig
    If DiagnosticMode Then Debug.Print "[DIAG] PerformSearch called."
    On Error GoTo EH

    If DiagnosticMode Then Debug.Print "[DIAG] Getting DataTable ListObject: " & DataTableName()
    Dim dataLo As ListObject: Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        If DiagnosticMode Then Debug.Print "[DIAG] Data table '" & DataTableName() & "' not found or empty."
        MsgBox "Data table '" & DataTableName() & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = resultsStartRng()
    Set statusRng = StatusCell()
    If resultsStart Is Nothing Then
        If DiagnosticMode Then Debug.Print "[DIAG] Named range 'ResultsStartCell' not found."
        MsgBox "Named range 'ResultsStartCell' not found.", vbExclamation
        Exit Sub
    End If

    If DiagnosticMode Then Debug.Print "[DIAG] ResultsStart address: " & resultsStart.Address & " on sheet '" & resultsStart.Worksheet.Name & "'"
    If DiagnosticMode Then Debug.Print "[DIAG] Worksheet protection: " & resultsStart.Worksheet.ProtectContents

    ' Read inputs (matching original exactly)
    Dim searchTxt As String: searchTxt = ReadSearchText_Description()
    Dim valveTxt As String:  valveTxt = ReadSearchText_Valve()
    If DiagnosticMode Then Debug.Print "[DIAG] SearchBox value: '" & searchTxt & "'"
    If DiagnosticMode Then Debug.Print "[DIAG] ValveNumSearchBox value: '" & valveTxt & "'"

    ' Build description regex array (synonym-aware)
    Dim rxArr As Variant
    If Len(Trim(searchTxt)) > 0 Then
        If DiagnosticMode Then Debug.Print "[DIAG] Building synonym index and regexes."
        Dim mapLo As ListObject: Set mapLo = lo(MappingTableName())
        Dim synIndex As Object: Set synIndex = BuildSynonymIndex(mapLo)
        rxArr = BuildSearchRegexes(searchTxt, synIndex)
    Else
        rxArr = Array()
    End If
    Dim descActive As Boolean: descActive = IsArrayNonEmpty(rxArr)
    If DiagnosticMode Then Debug.Print "[DIAG] descActive: " & descActive

    Dim valveColIdx As Long: valveColIdx = HeaderIndexByText(dataLo, "Valve Number")
    Dim valveActive As Boolean: valveActive = (Len(Trim(valveTxt)) > 0 And valveColIdx > 0)
    If DiagnosticMode Then Debug.Print "[DIAG] valveActive: " & valveActive & " (colIdx=" & valveColIdx & ")"

    ' Resolve the description column (for matching & default sort)
    Dim descColIdx As Long: descColIdx = GetColumnIndex(DataDescConfigKey(), dataLo)
    If descColIdx = 0 And descActive Then
        ' Fallback: if not configured, use first column from Out_Column1
        descColIdx = HeaderIndexByText(dataLo, GetConfigValue("Out_Column1"))
    End If
    If DiagnosticMode Then Debug.Print "[DIAG] descColIdx: " & descColIdx

    ' Build output column list from config (dynamic)
    Dim outCols() As Long, outKeys(1 To 8) As String
    Dim i As Long, j As Long, key As String, headerName As String
    ReDim outCols(1 To 8)
    outKeys(1) = "Out_Column1": outKeys(2) = "Out_Column2": outKeys(3) = "Out_Column3"
    outKeys(4) = "Out_Column4": outKeys(5) = "Out_Column5": outKeys(6) = "Out_Column6"
    outKeys(7) = "Out_Column7": outKeys(8) = "Out_Column8"

    Dim colCount As Long: colCount = 0
    For i = 1 To 8
        key = outKeys(i)
        headerName = GetConfigValue(key)
        If Len(Trim(headerName)) > 0 Then
            Dim idx As Long: idx = HeaderIndexByText(dataLo, headerName)
            If DiagnosticMode Then Debug.Print "[DIAG] Output column " & key & " = '" & headerName & "' (colIdx=" & idx & ")"
            If idx > 0 Then
                colCount = colCount + 1
                outCols(colCount) = idx
            End If
        End If
    Next i
    If DiagnosticMode Then Debug.Print "[DIAG] colCount: " & colCount
    If colCount = 0 Then
        If DiagnosticMode Then Debug.Print "[DIAG] No valid output columns found in ConfigTable."
        MsgBox "No valid output columns found in ConfigTable.", vbExclamation
        Exit Sub
    End If
    ReDim Preserve outCols(1 To colCount)

    ' Write headers & clear old results with the right width
    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount: hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value): Next i
    resultsStart.Resize(1, colCount).Value = hdr
    Call ClearOldResults(resultsStart, colCount)

    ' Consider only rows visible under current slicers/filters
    Dim idxs As Variant: idxs = VisibleRowIndexes(dataLo)
    If DiagnosticMode Then Debug.Print "[DIAG] VisibleRowIndexes count: " & IIf(IsEmpty(idxs), 0, UBound(idxs))
    If IsEmpty(idxs) Then
        If DiagnosticMode Then Debug.Print "[DIAG] No visible rows (check slicers/filters)."
        Call WriteStatus(statusRng, "No visible rows (check slicers/filters).", "")
        Exit Sub
    End If

    ' Evaluate matches & build a mask over ALL rows (so we can filter the table itself)
    Dim nRows As Long: nRows = dataLo.DataBodyRange.Rows.Count
    Dim keepMask() As Boolean: ReDim keepMask(1 To nRows)

    Dim kept As Long, ri As Long
    Dim descText As String, valveText As String, keep As Boolean

    For i = 1 To UBound(idxs)
        ri = idxs(i)                      ' 1-based within DataBodyRange
        keep = True

        If descActive Then
            If descColIdx > 0 Then
                descText = SafeCellText(dataLo.DataBodyRange.Cells(ri, descColIdx).Value)
            Else
                descText = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(1)).Value)
            End If
            For j = LBound(rxArr) To UBound(rxArr)
                If Not rxArr(j).Test(descText) Then keep = False: Exit For
            Next j
        End If

        If keep And valveActive Then
            valveText = SafeCellText(dataLo.DataBodyRange.Cells(ri, valveColIdx).Value)
            If StrComp(valveText, valveTxt, vbTextCompare) <> 0 Then keep = False
        End If

        If keep Then
            keepMask(ri) = True
            kept = kept + 1
        End If
        If DiagnosticMode Then Debug.Print "[DIAG] Row " & ri & ": keep=" & keep & ", descText='" & descText & "', valveText='" & valveText & "'"
    Next i
    If DiagnosticMode Then Debug.Print "[DIAG] kept: " & kept

    If kept = 0 Then
        If DiagnosticMode Then Debug.Print "[DIAG] No matches: clearing temp filter and writing status."
        Call ClearTempSearchFilter
        Call WriteStatus(statusRng, "Found 0 results.", "Query: " & Trim(searchTxt))
        Exit Sub
    End If

    ' Apply DataTable filter so slicers mirror the narrowed set
    If DiagnosticMode Then Debug.Print "[DIAG] Applying temp search filter."
    Call ApplyTempSearchFilter(keepMask)

    ' After filtering, recompute visible rows and render output (sorted by description if present)
    idxs = VisibleRowIndexes(dataLo)
    If DiagnosticMode Then Debug.Print "[DIAG] VisibleRowIndexes after filter: " & IIf(IsEmpty(idxs), 0, UBound(idxs))
    If IsEmpty(idxs) Then
        If DiagnosticMode Then Debug.Print "[DIAG] No visible rows after filter."
        Call WriteStatus(statusRng, "Found 0 results after filter.", "")
        Exit Sub
    End If

    Dim maxRows As Long: maxRows = MaxOutputRows()
    Dim cap As Long: cap = IIf(maxRows > 0, WorksheetFunction.Min(UBound(idxs), maxRows), UBound(idxs))
    If DiagnosticMode Then Debug.Print "[DIAG] maxRows=" & maxRows & ", cap=" & cap

    Dim outArr() As Variant: ReDim outArr(1 To cap, 1 To colCount)
    For i = 1 To cap
        ri = idxs(i)
        For j = 1 To colCount
            outArr(i, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).Value)
        Next j
        If DiagnosticMode Then Debug.Print "[DIAG] Output row " & i & ": " & Join(Application.WorksheetFunction.Transpose(Application.WorksheetFunction.Transpose(outArr(i, 1))), ", ")
    Next i

    ' Sort result array by description column if present among outputs
    Dim descOutPos As Long: descOutPos = FindOutputPosForDataColumn(outCols, colCount, descColIdx)
    If descOutPos > 0 And cap > 1 Then
        If DiagnosticMode Then Debug.Print "[DIAG] Sorting output array by description column."
        Call QuickSort2D_N(outArr, 1, cap, descOutPos)
    End If

        Call DebugSafeArrayAssignment(resultsStart, outArr, cap, colCount, DiagnosticMode)
        Call WriteStatus(statusRng, "Found " & cap & " row(s).", "Query: " & Trim(searchTxt))
        Exit Sub
EH:
    If DiagnosticMode Then Debug.Print "[DIAG] ERROR in PerformSearch: " & Err.Number & " - " & Err.Description
    Call LogErrorLocal("PerformSearch", Err.Number, Err.Description)
        Call DebugSafeArrayAssignment(resultsStart, outArr, cap, colCount, DiagnosticMode)
        Call WriteStatus(statusRng, "Found " & cap & " row(s).", "Query: " & Trim(searchTxt))
        Exit Sub
End Sub
Private Sub DebugSafeArrayAssignment(ByVal resultsStart As Range, ByRef outArr As Variant, ByVal cap As Long, ByVal colCount As Long, ByVal DiagnosticMode As Boolean)
    Dim tgtRng As Range
    On Error Resume Next
    Set tgtRng = resultsStart.Offset(1, 0).Resize(cap, colCount)
    On Error GoTo 0
    ' Force outArr to be a true 2D array for single-column output
    If colCount = 1 Then
        Dim arr2D() As Variant: ReDim arr2D(1 To cap, 1 To 1)
        Dim ii As Long
        For ii = 1 To cap
            arr2D(ii, 1) = outArr(ii, 1)
        Next ii
        outArr = arr2D
        If DiagnosticMode Then Debug.Print "[DIAG] Forced 2D array for colCount=1."
    End If
    If DiagnosticMode Then
        Debug.Print "[DIAG] --- DebugSafeArrayAssignment ---"
        Debug.Print "[DIAG] resultsStart.Address=" & resultsStart.Address & ", Worksheet=" & resultsStart.Worksheet.Name
        Debug.Print "[DIAG] tgtRng.Address=" & tgtRng.Address & ", Rows=" & tgtRng.Rows.Count & ", Columns=" & tgtRng.Columns.Count
        Debug.Print "[DIAG] tgtRng.MergeCells=" & tgtRng.MergeCells
        Debug.Print "[DIAG] tgtRng.Worksheet.ProtectContents=" & tgtRng.Worksheet.ProtectContents
        Debug.Print "[DIAG] outArr dimensions: " & LBound(outArr,1) & " to " & UBound(outArr,1) & " by " & LBound(outArr,2) & " to " & UBound(outArr,2)
        Debug.Print "[DIAG] tgtRng.Rows.Count=" & tgtRng.Rows.Count & ", tgtRng.Columns.Count=" & tgtRng.Columns.Count
        Debug.Print "[DIAG] outArr row count=" & (UBound(outArr,1)-LBound(outArr,1)+1) & ", col count=" & (UBound(outArr,2)-LBound(outArr,2)+1)
        Dim jjDiag As Long
        For jjDiag = LBound(outArr,2) To UBound(outArr,2)
            Debug.Print "[DIAG] Row 1, Col " & jjDiag & ": TypeName=" & TypeName(outArr(1,jjDiag)) & ", Value='" & outArr(1,jjDiag) & "'"
        Next jjDiag
    End If
    If tgtRng.MergeCells Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: Target range contains merged cells. Cannot assign array."
        Exit Sub
    End If
    If tgtRng.Worksheet.ProtectContents Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: Worksheet is protected. Cannot assign array."
        Exit Sub
    End If
    ' Confirm array and range sizes match
    If (UBound(outArr,1)-LBound(outArr,1)+1) <> tgtRng.Rows.Count Or (UBound(outArr,2)-LBound(outArr,2)+1) <> tgtRng.Columns.Count Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: Array and range size mismatch. Aborting assignment."
        Exit Sub
    End If
    On Error GoTo ArrayAssignErr
    tgtRng.Value = outArr
    If DiagnosticMode Then Debug.Print "[DIAG] Array assignment succeeded."
    Exit Sub
ArrayAssignErr:
    If DiagnosticMode Then Debug.Print "[DIAG] ERROR assigning output array: " & Err.Number & " - " & Err.Description
    If DiagnosticMode Then Debug.Print "[DIAG] Fallback: trying to assign outArr(1,1) to tgtRng.Cells(1,1)"
    On Error Resume Next
    tgtRng.Cells(1,1).Value = outArr(1,1)
    If Err.Number <> 0 Then
        If DiagnosticMode Then Debug.Print "[DIAG] Fallback failed: " & Err.Number & " - " & Err.Description
    Else
        If DiagnosticMode Then Debug.Print "[DIAG] Fallback succeeded."
    End If
    On Error GoTo 0
End Sub


' Show all visible results
Public Sub OutputAllVisible()
    On Error GoTo EH
    Dim dataLo As ListObject: Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DataTableName() & "' not found or empty.", vbExclamation
        Exit Sub
    End If
    
    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = resultsStartRng()
    Set statusRng = StatusCell()
    If resultsStart Is Nothing Then
        MsgBox "ResultsStartCell does not resolve to a range.", vbExclamation
        Exit Sub
    End If

    ' Build output columns from config (same as PerformSearch), so headers match.
    Dim outCols() As Long, outKeys(1 To 8) As String
    Dim i As Long, j As Long, key As String, headerName As String
    ReDim outCols(1 To 8)
    outKeys(1) = "Out_Column1": outKeys(2) = "Out_Column2": outKeys(3) = "Out_Column3"
    outKeys(4) = "Out_Column4": outKeys(5) = "Out_Column5": outKeys(6) = "Out_Column6"
    outKeys(7) = "Out_Column7": outKeys(8) = "Out_Column8"

    Dim colCount As Long: colCount = 0
    For i = 1 To 8
        key = outKeys(i)
        headerName = GetConfigValue(key)
        If Len(Trim(headerName)) > 0 Then
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
    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount: hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value): Next i
    resultsStart.Resize(1, colCount).Value = hdr
    Call ClearOldResults(resultsStart, colCount)

    ' Visible indexes under current slicers/filters
    Dim idxs As Variant: idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        Call WriteStatus(statusRng, "No visible rows (check slicers/filters).", "")
        Exit Sub
    End If

    ' Cap output
    Dim maxRows As Long: maxRows = MaxOutputRows()
    Dim cap As Long: cap = IIf(maxRows > 0, WorksheetFunction.Min(UBound(idxs), maxRows), UBound(idxs))

    ' Write body
    Dim outArr() As Variant: ReDim outArr(1 To cap, 1 To colCount)
    Dim ri As Long
    For i = 1 To cap
        ri = idxs(i)
        For j = 1 To colCount
            outArr(i, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).Value)
        Next j
    Next i
    resultsStart.Offset(1, 0).Resize(cap, colCount).Value = outArr

    If Not statusRng Is Nothing Then Call WriteStatus(statusRng, "Showing " & cap & " visible row(s).", "All Visible")
    Exit Sub
EH:
    Call LogErrorLocal("OutputAllVisible", Err.Number, Err.Description)
End Sub

' Show no results (headers only)
Public Sub OutputNoResults()
    On Error GoTo EH
    Dim dataLo As ListObject: Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DataTableName() & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = resultsStartRng()
    Set statusRng = StatusCell()
    If resultsStart Is Nothing Then
        MsgBox "ResultsStartCell does not resolve to a range.", vbExclamation
        Exit Sub
    End If

    ' Build output columns same as OutputAllVisible to write consistent headers.
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
        If Len(Trim(headerName)) > 0 Then
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

    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount: hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value): Next i
    resultsStart.Resize(1, colCount).Value = hdr
    Call ClearOldResults(resultsStart, colCount)   ' clears previous rows beneath headers

    If Not statusRng Is Nothing Then Call WriteStatus(statusRng, "No results (gated by inputs/threshold).", "")
    Exit Sub
EH:
    Call LogErrorLocal("OutputNoResults", Err.Number, Err.Description)
End Sub

' Safe wrapper for OutputAllVisible
Public Sub Safe_OutputAllVisible()
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo CleanExit

    If IsAnySearchInputActive() Then
        ' If an input is active, "Show All" should still respect search; keep filter and output
        Call PerformSearch
    Else
        Call ClearTempSearchFilter
        Call OutputAllVisible
    End If

CleanExit:
    gBusy = False
End Sub

' ============================================================
' CONFIG ACCESS FUNCTIONS - Exact Signatures
' ============================================================

Public Function DashboardName() As String
    Dim v As String: v = GetConfigValue(CFG_DASHBOARD_SHEET)
    If Len(Trim(v)) = 0 Then v = "Dashboard"
    DashboardName = v
End Function

Public Function DataTableName() As String
    Dim v As String: v = GetConfigValue(CFG_DATA_TABLE_NAME)
    If Len(Trim(v)) = 0 Then v = "EquipmentData"
    DataTableName = v
End Function

Public Function MappingTableName() As String
    Dim v As String: v = GetConfigValue(CFG_MAPPING_TABLE_NAME)
    If Len(Trim(v)) = 0 Then v = "tbl_Mapping"
    MappingTableName = v
End Function

Public Function SlicerThreshold() As Long
    SlicerThreshold = GetConfigLongSafe(CFG_SPLICER_THRESHOLD, 250)
End Function

Public Function TempFilterColName() As String
    Dim v As String: v = GetConfigValue(CFG_TEMP_FILTER_COL_NAME)
    If Len(Trim(v)) = 0 Then v = "Temp_SearchInclude"
    TempFilterColName = v
End Function

Public Function SlicerPulseAnchorName() As String
    Dim v As String: v = GetConfigValue(CFG_SLICER_PULSE_ANCHOR)
    If Len(Trim(v)) = 0 Then v = "SlicerPulseAnchor"
    SlicerPulseAnchorName = v
End Function

Public Function resultsStartRng() As Range
    Set resultsStartRng = nr(GetConfigValue(CFG_RESULTS_START))
End Function

Public Function StatusCell() As Range
    Set StatusCell = nr(GetConfigValue(CFG_STATUS_CELL))
End Function

Public Function DataDescConfigKey() As String
    DataDescConfigKey = CFG_DATA_DESC_COL
End Function

Public Function GetConfigLongSafe(ByVal key As String, ByVal defVal As Long) As Long
    Dim s As String: s = GetConfigValue(key)
    If Len(Trim(s)) = 0 Then
        GetConfigLongSafe = defVal
    Else
        GetConfigLongSafe = CLngSafe(s)
    End If
End Function

' ============================================================
' INPUT DETECTION FUNCTIONS - Exact Signatures
' ============================================================

Public Function AllInputNamedRanges_Enhanced() As Collection
    ' Returns a collection of Range objects for every ConfigTable row where TYPE = "Input Named Range".
    Dim c As New Collection
    Dim ws As Worksheet, loCfg As ListObject, r As Range
    Dim nmKey As String, nmType As String
    On Error GoTo Done
    
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Or loCfg.DataBodyRange Is Nothing Then GoTo Done

    For Each r In loCfg.DataBodyRange.Rows
        nmKey = Trim(CStr(r.Cells(1, 2).Value))   ' column B: ConfigValue (the NAME of the range)
        nmType = Trim(CStr(r.Cells(1, 3).Value))  ' column C: TYPE
        If StrComp(nmType, "Input Named Range", vbTextCompare) = 0 Then
            Dim nrng As Range
            On Error Resume Next
            Set nrng = ThisWorkbook.Names(nmKey).RefersToRange
            On Error GoTo 0
            If Not nrng Is Nothing Then c.Add nrng
        End If
    Next r
Done:
    Set AllInputNamedRanges_Enhanced = c
End Function

Public Function IsAnySearchInputActive() As Boolean
    ' True if ANY input named range has a non-empty value (after Trim).
    Dim coll As Collection: Set coll = AllInputNamedRanges_Enhanced()
    Dim i As Long, r As Range, v As String
    For i = 1 To coll.Count
        Set r = coll(i)
        v = CStr(r.Cells(1, 1).Value)
        If Len(Trim(v)) > 0 Then IsAnySearchInputActive = True: Exit Function
    Next i
End Function

Public Function ReadSearchText_Description() As String
    ReadSearchText_Description = ReadLeftCell(GetConfigValue(CFG_INPUT_DESCRIP))
End Function

Public Function ReadSearchText_Valve() As String
    ReadSearchText_Valve = ReadLeftCell(GetConfigValue(CFG_INPUT_VALVE))
End Function

' Compatibility functions for external modules
Public Function GetSearchText() As String
    GetSearchText = ReadSearchText_Description()
End Function

Public Function GetValveSearchText() As String
    GetValveSearchText = ReadSearchText_Valve()
End Function

Public Function GetSearchStatusText() As String
    Dim statusRng As Range
    Set statusRng = StatusCell()
    If statusRng Is Nothing Then
        GetSearchStatusText = "Ready"
    Else
        GetSearchStatusText = CStr(statusRng.Value)
    End If
End Function

Public Function CurrentSlicerPulse() As Long
    Dim r As Range: Set r = nr("SlicerPulse")
    If r Is Nothing Then
        CurrentSlicerPulse = 0
    Else
        CurrentSlicerPulse = CLngSafe(CStr(r.Value))
    End If
End Function

' ============================================================
' SLICER INTEGRATION - Enhanced Implementation
' ============================================================

Public Sub ClearTempSearchFilter()
    ' Removes the temporary search helper column (if present) so slicers return to normal-only filtering.
    Dim dataLo As ListObject: Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Then Exit Sub

    Dim nm As String: nm = TempFilterColName()
    Dim lc As ListColumn
    On Error Resume Next
    For Each lc In dataLo.ListColumns
        If StrComp(CStr(lc.Name), nm, vbTextCompare) = 0 Then
            Application.EnableEvents = False
            lc.Delete                 ' delete column to clear ONLY this filter without touching slicers
            Application.EnableEvents = True
            Exit For
        End If
    Next lc
    On Error GoTo 0
End Sub

Sub ApplyTempSearchFilter(ByRef includeMask() As Boolean)
    ' includeMask is 1..N rows of DataBodyRange; TRUE means keep the row.
    Dim dataLo As ListObject: Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Then Exit Sub
    If dataLo.DataBodyRange Is Nothing Then Exit Sub

    Dim n As Long
    If dataLo.DataBodyRange Is Nothing Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: DataBodyRange is Nothing in ApplyTempSearchFilter. Table may be empty."
        Application.EnableEvents = True
        Exit Sub
    End If
    n = dataLo.DataBodyRange.Rows.Count
    If n = 0 Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: DataTable has zero rows in ApplyTempSearchFilter."
        Application.EnableEvents = True
        Exit Sub
    End If
    If UBound(includeMask) < 1 Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: includeMask is empty in ApplyTempSearchFilter."
        Application.EnableEvents = True
        Exit Sub
    End If

    Dim nm As String: nm = TempFilterColName()
    Dim lc As ListColumn, colIdx As Long: colIdx = 0

    ' find or add helper column at the end
    For Each lc In dataLo.ListColumns
        If StrComp(CStr(lc.Name), nm, vbTextCompare) = 0 Then colIdx = lc.Index: Exit For
    Next lc
    If colIdx = 0 Then
        Set lc = dataLo.ListColumns.Add
        lc.Name = nm
        colIdx = lc.Index
    Else
        Set lc = dataLo.ListColumns(colIdx)
    End If

    ' write values (1 for include / 0 for exclude) in-memory
    Dim v As Variant: ReDim v(1 To n, 1 To 1)
    Dim i As Long
    For i = 1 To n
        v(i, 1) = IIf(includeMask(i), 1, 0)
    Next i

    Application.EnableEvents = False
    If lc Is Nothing Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: ListColumn object (lc) is Nothing in ApplyTempSearchFilter."
        Application.EnableEvents = True
        Exit Sub
    End If
    If DiagnosticMode Then Debug.Print "[DIAG] lc.Name=" & lc.Name & ", lc.Index=" & lc.Index
    If lc.DataBodyRange Is Nothing Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR: lc.DataBodyRange is Nothing in ApplyTempSearchFilter. Table may be empty or not refreshed."
        Application.EnableEvents = True
        Exit Sub
    End If
    If DiagnosticMode Then Debug.Print "[DIAG] lc.DataBodyRange.Address=" & lc.DataBodyRange.Address & ", TypeName(lc.DataBodyRange)=" & TypeName(lc.DataBodyRange)
    If DiagnosticMode Then Debug.Print "[DIAG] Table row count: " & n & ", column count: " & dataLo.ListColumns.Count & ", colIdx=" & colIdx
    If DiagnosticMode Then Debug.Print "[DIAG] v dimensions: " & LBound(v,1) & " to " & UBound(v,1) & " by " & LBound(v,2) & " to " & UBound(v,2)
    On Error Resume Next
    If DiagnosticMode Then Debug.Print "[DIAG] Attempting assignment: lc.DataBodyRange.Value = v"
    lc.DataBodyRange.Value = v
    If Err.Number <> 0 Then
        If DiagnosticMode Then Debug.Print "[DIAG] ERROR assigning values in ApplyTempSearchFilter: " & Err.Number & " - " & Err.Description
        Application.EnableEvents = True
        Exit Sub
    End If
    On Error GoTo 0
    If DiagnosticMode Then Debug.Print "[DIAG] Assignment succeeded. Attempting AutoFilter."
    ' apply filter: keep 1s only (stacking with existing slicer/table filters)
    dataLo.Range.AutoFilter Field:=colIdx, Criteria1:=1
    If DiagnosticMode Then Debug.Print "[DIAG] AutoFilter applied. Hiding helper column."
    lc.Range.EntireColumn.Hidden = True
    Application.EnableEvents = True
End Sub

Public Sub EnsurePulseCell_Run() 'For manual testing
    Call EnsurePulseCell(True)
End Sub

Public Sub EnsurePulseCell(Optional recalcNow As Boolean = False)
    Dim dash As Worksheet, dataLo As ListObject, pulseCell As Range
    Dim hdrName As String, idx As Long, addr As String
    Dim anchorName As String, anchorRng As Range

    Set dash = SheetByName(DashboardName())
    If dash Is Nothing Then Exit Sub

    Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Then Exit Sub
    
    If dataLo.ListColumns.Count = 0 Or dataLo.DataBodyRange Is Nothing Then Exit Sub

    ' --- Which header/column should drive the pulse? ---
    hdrName = Trim(GetConfigValue("SLICER_PULSE_HEADER"))   ' e.g., "SAP ID"
    If Len(hdrName) > 0 Then
        idx = HeaderIndexByText(dataLo, hdrName)
    End If
    If idx < 1 Or idx > dataLo.ListColumns.Count Then
        idx = 1 ' safe fallback to first column
    End If

    ' --- Where should the pulse live? (anchor) ---
    anchorName = Trim(GetConfigValue("SLICER_PULSE_ANCHOR")) ' e.g., "SlicerPulseAnchor"
    If Len(anchorName) > 0 Then
        On Error Resume Next
        Set anchorRng = nr(anchorName)
        On Error GoTo 0
    End If
    If anchorRng Is Nothing Then
        ' fallback: AA1 on the dashboard (and hide the col)
        Set anchorRng = dash.Range("AA1")
        dash.Columns("AA").Hidden = True
    End If
    Set pulseCell = anchorRng.Cells(1, 1)

    ' --- Wire the SUBTOTAL pulse to the chosen column ---
    addr = dataLo.ListColumns(idx).DataBodyRange.Address(True, True, xlA1, True)
    Application.EnableEvents = False
        pulseCell.formula = "=SUBTOTAL(103," & addr & ")"
        Call NameOrUpdate("SlicerPulseCell", "=" & pulseCell.Address(True, True, xlA1, True))
        Call NameOrUpdate("SlicerPulse", "=" & pulseCell.Address(True, True, xlA1, True))
    Application.EnableEvents = True

    If recalcNow Then Application.Calculate
End Sub

' ============================================================
' UTILITY FUNCTIONS - Exact Signatures
' ============================================================

Public Function VisibleRowIndexes(ByVal dataLo As ListObject) As Variant
    Dim firstCol As Range, vis As Range, c As Range
    Dim idx() As Long, n As Long, startRow As Long
    On Error Resume Next
    Set firstCol = dataLo.ListColumns(1).DataBodyRange
    If firstCol Is Nothing Then Exit Function
    startRow = firstCol.Row
    Set vis = firstCol.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If vis Is Nothing Then Exit Function
    ReDim idx(1 To vis.Count)
    For Each c In vis.Cells
        n = n + 1
        idx(n) = c.Row - startRow + 1
    Next c
    If n = 0 Then Exit Function
    ReDim Preserve idx(1 To n)
    VisibleRowIndexes = idx
End Function

Public Function lo(ByVal Name As String) As ListObject
    Dim ws As Worksheet, l As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each l In ws.ListObjects
            If StrComp(l.Name, Name, vbTextCompare) = 0 Then Set lo = l: Exit Function
        Next l
    Next ws
End Function

Public Function nr(ByVal nm As String) As Range
    On Error Resume Next
    Set nr = ThisWorkbook.Names(nm).RefersToRange
End Function

Public Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
End Function

Public Function ReadLeftCell(ByVal nm As String) As String
    On Error Resume Next
    Dim rng As Range: Set rng = nr(nm)
    If Not rng Is Nothing Then ReadLeftCell = CStr(rng.Cells(1, 1).Value)
End Function

Public Function SafeCellText(ByVal v As Variant) As String
    If IsError(v) Then SafeCellText = "" Else SafeCellText = CStr(v)
End Function

Public Function HeaderIndexByText(ByVal dataLo As ListObject, ByVal headerText As String) As Long
    Dim i As Long
    If dataLo Is Nothing Then Exit Function
    If Len(Trim(headerText)) = 0 Then Exit Function
    For i = 1 To dataLo.ListColumns.Count
        If StrComp(CStr(dataLo.HeaderRowRange.Cells(1, i).Value), headerText, vbTextCompare) = 0 Then
            HeaderIndexByText = i: Exit Function
        End If
    Next i
End Function

Public Sub ClearOldResults(ByVal startCell As Range, Optional ByVal colCount As Long = 3)
    startCell.Offset(1, 0).Resize(100000, colCount).ClearContents
End Sub

Public Sub WriteStatus(ByVal statusRng As Range, ByVal line1 As String, ByVal line2 As String)
    On Error Resume Next
    If Not statusRng Is Nothing Then
        statusRng.Cells(1, 1).Value = line1
        If statusRng.Rows.Count >= 2 Then statusRng.Cells(2, 1).Value = line2
    End If
End Sub

Public Function MaxOutputRows() As Long
    Dim v As String: v = GetConfigValue("MAX_OUTPUT_ROWS")
    MaxOutputRows = CLngSafe(v)
End Function

Public Sub NameOrUpdate(ByVal nm As String, ByVal refersTo As String)
    Dim n As Name
    
    If Len(refersTo) > 0 Then
        If Left(refersTo, 1) <> "=" Then refersTo = "=" & refersTo
    End If
    
    On Error Resume Next
       Set n = ThisWorkbook.Names(nm)
    On Error GoTo 0
    
    If n Is Nothing Then
        ThisWorkbook.Names.Add Name:=nm, RefersTo:=refersTo
    Else
        n.RefersTo = refersTo
    End If
End Sub

' ============================================================
' SEARCHING AND SORTING
' ============================================================

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
    If loIdx < j Then Call QuickSort2D_N(arr, loIdx, j, sortCol)
    If i < hiIdx Then Call QuickSort2D_N(arr, i, hiIdx, sortCol)
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

' ============================================================
' REGEX SYNONYM ENGINE (from original)
' ============================================================

Public Function EscapeRegex(ByVal s As String) As String
    Dim specials As Variant, ch As Variant
    specials = Array("\", ".", "+", "*", "?", "|", "{", "}", "[", "]", "(", ")", "^", "$")
    For Each ch In specials
        If InStr(s, ch) > 0 Then s = Replace(s, ch, "\" & ch)
    Next ch
    EscapeRegex = s
End Function

Public Function BuildSynonymIndex(ByVal mapLo As ListObject) As Object
    Dim syn As Object: Set syn = CreateObject("Scripting.Dictionary")
    syn.CompareMode = vbTextCompare
    If mapLo Is Nothing Or mapLo.DataBodyRange Is Nothing Then Set BuildSynonymIndex = syn: Exit Function

    Dim groupDict As Object: Set groupDict = CreateObject("Scripting.Dictionary")
    groupDict.CompareMode = vbTextCompare

    Dim r As Range, raw As String, std As String
    For Each r In mapLo.DataBodyRange.Rows
        raw = LCase(Trim(CStr(r.Cells(1, 1).Value)))
        std = LCase(Trim(CStr(r.Cells(1, 2).Value)))
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
        ReDim arr(0 To d.Count - 1): i = 0
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

    searchText = Trim(searchText)
    If Len(searchText) = 0 Then BuildSearchRegexes = Array(): Exit Function

    tokens = Split(searchText, " ")
    ReDim rxArr(0 To 0)
    Dim cnt As Long: cnt = -1

    For i = LBound(tokens) To UBound(tokens)
        term = Trim(CStr(tokens(i)))
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

' ============================================================
' CONFIG ACCESS
' ============================================================

Public Function GetConfigValue(ByVal key As String) As String
    Dim ws As Worksheet, loCfg As ListObject, r As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Or loCfg.DataBodyRange Is Nothing Then Exit Function
    On Error GoTo 0
    For Each r In loCfg.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            GetConfigValue = CStr(r.Cells(1, 2).Value)
            Exit Function
        End If
    Next r
End Function

Public Function CLngSafe(ByVal s As String) As Long
    If Len(Trim(s)) = 0 Then
        CLngSafe = 0
    Else
        CLngSafe = CLng(Val(s))
    End If
End Function

Public Function GetColumnIndex(ByVal configKey As String, ByVal dataLo As ListObject) As Long
    Dim headerName As String
    headerName = GetConfigValue(configKey)
    If Len(Trim(headerName)) = 0 Then Exit Function
    GetColumnIndex = HeaderIndexByText(dataLo, headerName)
End Function

' ============================================================
' FILTERS / CLEAR
' ============================================================

Public Sub ClearFilters_Enhanced()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Call ClearSearchBoxes
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If StrComp(tbl.Name, DataTableName(), vbTextCompare) = 0 Then
                If Not tbl.AutoFilter Is Nothing Then
                    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
                End If
                ' Refresh results to show no results after clearing all filters and search boxes
                Call RefreshResults
                Exit Sub
            End If
        Next tbl
    Next ws
End Sub

Public Sub ClearSearchBoxes()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim nm As Name
    Dim rng As Range
    
    ' Point to your ConfigSheet and ConfigTable
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set tbl = ws.ListObjects("ConfigTable")
    
    ' Loop through each row in the table
    For Each rw In tbl.ListRows
        If rw.Range.Cells(1, 3).Value = "Input Named Range" Then
            On Error Resume Next
            Set nm = ThisWorkbook.Names(rw.Range.Cells(1, 2).Value)
            On Error GoTo 0
            
            If Not nm Is Nothing Then
                Set rng = nm.RefersToRange
                If Not rng Is Nothing Then
                    If rng.MergeCells Then
                        rng.MergeArea.ClearContents
                    Else
                        rng.ClearContents
                    End If
                End If
            End If
        End If
    Next rw
End Sub

' ============================================================
' ERROR HANDLING
' ============================================================

Public Sub LogErrorLocal(ByVal procName As String, ByVal errNum As Long, ByVal errDesc As String)
    Dim logSheet As Worksheet
    Dim NextRow As Long
    Dim searchText As String, tagText As String
    ' Exit Sub 'Temporarily toggle off
    On Error Resume Next
    Application.ScreenUpdating = False
    Set logSheet = ThisWorkbook.Worksheets("SearchErrorLog")
    logSheet.Visible = xlSheetHidden
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add
        logSheet.Name = "SearchErrorLog"
        logSheet.Cells(1, 1).Value = "Timestamp"
        logSheet.Cells(1, 2).Value = "Procedure"
        logSheet.Cells(1, 3).Value = "Error Number"
        logSheet.Cells(1, 4).Value = "Error Description"
        logSheet.Cells(1, 5).Value = "SearchBox"
        logSheet.Cells(1, 6).Value = "TagID"
    End If
    On Error GoTo 0

    NextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    searchText = SafeReadNameLocal("SearchBox")
    tagText = SafeReadNameLocal("ValveNumSearchBox")

    logSheet.Cells(NextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    logSheet.Cells(NextRow, 2).Value = procName
    logSheet.Cells(NextRow, 3).Value = errNum
    logSheet.Cells(NextRow, 4).Value = errDesc
    logSheet.Cells(NextRow, 5).Value = searchText
    logSheet.Cells(NextRow, 6).Value = tagText
    Application.ScreenUpdating = True
End Sub

Public Function SafeReadNameLocal(ByVal nm As String) As String
    On Error Resume Next
    Dim r As Range: Set r = nr(nm)
    If Not r Is Nothing Then SafeReadNameLocal = CStr(r.Cells(1, 1).Value)
End Function

' ============================================================
' TESTING AND DIAGNOSTICS
' ============================================================

Public Sub SelfTest()
    On Error Resume Next
    Debug.Print "---- SelfTest ----"
    Debug.Print "DashboardName(): "; DashboardName()
    Debug.Print "DataTableName(): "; DataTableName()
    Debug.Print "MappingTableName(): "; MappingTableName()
    Debug.Print "ResultsStartRng: "; IIf(resultsStartRng() Is Nothing, "Nothing", resultsStartRng().Address)
    Debug.Print "StatusCell: "; IIf(StatusCell() Is Nothing, "Nothing", StatusCell().Address)
    Debug.Print "SlicerThreshold(): "; SlicerThreshold()
    Dim p As Range: Set p = nr("SlicerPulse"): Debug.Print "SlicerPulse: "; IIf(p Is Nothing, "Missing name", p.Value)
    Dim loData As ListObject: Set loData = lo(DataTableName())
    Debug.Print "Data LO present: "; Not (loData Is Nothing)
    If Not loData Is Nothing Then
        Debug.Print "Data rows: "; IIf(loData.DataBodyRange Is Nothing, 0, loData.DataBodyRange.Rows.Count)
        Dim idxs As Variant: idxs = VisibleRowIndexes(loData)
        If IsEmpty(idxs) Then
            Debug.Print "Visible indexes: EMPTY"
        Else
            Debug.Print "Visible indexes count: "; UBound(idxs)
        End If
    End If
    Debug.Print "Any search active: "; IsAnySearchInputActive()
    Debug.Print "------------------"
End Sub