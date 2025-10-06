'Attribute VB_Name = "mod_PrimaryConsolidatedModule3"  ' commented for copy/paste
Option Explicit

'''Attribute VB_Name = "mod_PrimaryConsolidatedModule"
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


'
'============================================================
' EQUIPMENT SEARCH ENGINE (Dev)   Consolidated Single-Module Layout
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
' ------------------------------------------------
'  Constant key strings (compile-time)
' ------------------------------------------------
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
'==============================
' SMART SEARCH  PRODUCTION (Description + Tag ID)
'==============================

' Tag search behavior
Public Const TAG_SEARCH_MIN_LEN As Long = 3

' Policy: when inputs are empty, do NOT auto-show all results.
' This is enforced in RefreshResults, Safe_PerformSearch, and Safe_OutputAllVisible.
' The Show All button remains available to intentionally display everything.

' Re-entrancy guard
Public gBusy As Boolean





'=============================================
' Smart Search vNext � slicer-aware gating & dynamic inputs
' Date: 2025-10-05
'=============================================
'
' WHAT'S NEW (high-level)
' -----------------------
' 1) Results gating now considers BOTH search-input activity and slicer pulse:
'    - If ANY search input (ConfigTable TYPE = "Input Named Range") is non-empty ? show results.
'    - Else, if SlicerPulse < SPLICER_THRESHOLD ? show results.
'    - Else ? default to headers-only (no results shown).
'
' 2) Slicers now automatically "match" search results:
'    - When a search is active, the DataTable is filtered to the matched rows via a temporary
'      helper column (name from Config: TEMP_FILTER_COL_NAME, default "Temp_SearchInclude").
'      Because slicers are bound to DataTable, they immediately reflect the narrowed set.
'    - When no search inputs are active, the temporary filter is removed so slicers reflect
'      only their own state.
'
' 3) Dynamic recognition of *all* search inputs:
'    - We no longer hardcode input named ranges in gating or Dashboard change events.
'      Instead, we enumerate every ConfigTable row whose TYPE is "Input Named Range".
'
' 4) No hardcoded cell addresses for the slicer pulse anchor:
'    - EnsurePulseCell now looks for a named anchor from Config (SLICER_PULSE_ANCHOR ?
'      named range, default "SlicerPulseAnchor"). If missing, it gracefully falls back
'      (with a comment) without breaking.
'
' 5) Comments & config-first philosophy:
'    - Any variable likely to change lives in ConfigTable. All sheet/range access uses names.
'
' ---------------------------------------------
' CONFIGTABLE KEYS (add/confirm)
' ---------------------------------------------
'  - DASHBOARD_SHEET                 e.g., "Dashboard"
'  - SPLICER_THRESHOLD               e.g., 250
'  - TEMP_FILTER_COL_NAME            e.g., "Temp_SearchInclude"
'  - SLICER_PULSE_ANCHOR             e.g., "SlicerPulseAnchor"   ' a single hidden cell on Dashboard
'  - InputCell_DescripSearch         (existing)
'  - InputCell_ValveNumSearch        (existing)
'  - Out_Column1..Out_Column8        (existing)
'  - DataTable_EquipDescription      (existing)
'
' NOTE: Create a named range on the Dashboard sheet called SlicerPulseAnchor (or whatever
'       you set SLICER_PULSE_ANCHOR to). It can point to any hidden cell; code will populate it.
'
' ========================================================================================
'  MODULE: mod_PrimaryConsolidatedModule � NEW / UPDATED MEMBERS
' ========================================================================================


' -------------------------------------------------------
' UPDATED: PerformSearch � now also filters the DataTable
' -------------------------------------------------------





'------------------------------------------------------
'Config-backed getters (no literals used)
' -----------------------------------------------------
Public Function DashboardName() As String
    Dim v As String: v = GetConfigValue(CFG_DASHBOARD_SHEET)
    If Len(Trim$(v)) = 0 Then v = "Dashboard"
    DashboardName = v
End Function

Private Function DataTableName() As String ' Demoted to Private; canonical public version in mod_SearchEngine_Enhanced
    Dim v As String: v = GetConfigValue(CFG_DATA_TABLE_NAME)
    If Len(Trim$(v)) = 0 Then v = "EquipmentData"
    DataTableName = v
End Function

Private Function MappingTableName() As String ' Demoted to Private; canonical public version in mod_SearchEngine_Enhanced
    Dim v As String: v = GetConfigValue(CFG_MAPPING_TABLE_NAME)
    If Len(Trim$(v)) = 0 Then v = "tbl_Mapping"
    MappingTableName = v
End Function

Public Function SlicerThreshold() As Long
    SlicerThreshold = GetConfigLongSafe(CFG_SPLICER_THRESHOLD, 250)
End Function

Public Function TempFilterColName() As String
    Dim v As String: v = GetConfigValue(CFG_TEMP_FILTER_COL_NAME)
    If Len(Trim$(v)) = 0 Then v = "Temp_SearchInclude"
    TempFilterColName = v
End Function

Public Function SlicerPulseAnchorName() As String
    Dim v As String: v = GetConfigValue(CFG_SLICER_PULSE_ANCHOR)
    If Len(Trim$(v)) = 0 Then v = "SlicerPulseAnchor"
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
    If Len(Trim$(s)) = 0 Then
        GetConfigLongSafe = defVal
    Else
        GetConfigLongSafe = CLngSafe(s)
    End If
End Function

' ----------------------------------------------------------------
' Dynamic discovery of search input named ranges
' ----------------------------------------------------------------
Public Function AllInputNamedRanges() As Collection
    ' Returns a collection of Range objects for every ConfigTable row where TYPE = "Input Named Range".
    Dim c As New Collection
    Dim ws As Worksheet, loCfg As ListObject, r As Range
    Dim nmKey As String, nmType As String
    On Error GoTo Done
    
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Or loCfg.DataBodyRange Is Nothing Then GoTo Done

    For Each r In loCfg.DataBodyRange.Rows
        nmKey = Trim$(CStr(r.Cells(1, 2).Value))   ' column B: ConfigValue (the NAME of the range)''??? Will this still work if I add a config column?
        nmType = Trim$(CStr(r.Cells(1, 3).Value))  ' column C: TYPE                               ''???
        If StrComp(nmType, "Input Named Range", vbTextCompare) = 0 Then
            Dim nrng As Range
            On Error Resume Next
            Set nrng = ThisWorkbook.Names(nmKey).RefersToRange
            On Error GoTo 0
            If Not nrng Is Nothing Then c.Add nrng
        End If
    Next r
Done:
    Set AllInputNamedRanges = c
End Function

Public Function IsAnySearchInputActive() As Boolean
    ' True if ANY input named range has a non-empty value (after Trim).
    Dim coll As Collection: Set coll = AllInputNamedRanges()
    Dim i As Long, r As Range, v As String
    For i = 1 To coll.Count
        Set r = coll(i)
        v = CStr(r.Cells(1, 1).Value)
        If Len(Trim$(v)) > 0 Then IsAnySearchInputActive = True: Exit Function
    Next i
End Function

Public Function ReadSearchText_Description() As String
    ReadSearchText_Description = ReadLeftCell(GetConfigValue(CFG_INPUT_DESCRIP))
End Function

Public Function ReadSearchText_Valve() As String
    ReadSearchText_Valve = ReadLeftCell(GetConfigValue(CFG_INPUT_VALVE))
End Function

Private Function CurrentSlicerPulse() As Long
    Dim r As Range: Set r = nr("SlicerPulse")
    If r Is Nothing Then
        CurrentSlicerPulse = 0
    Else
        CurrentSlicerPulse = CLngSafe(CStr(r.Value))
    End If
End Function


' ---------------------------------------------------------------
' SECTION 5 � Temp search filter (helper column in DataTable)
' ---------------------------------------------------------------
Public Sub ClearTempSearchFilter()
    ' Removes the temporary search helper column (if present) so slicers return to normal-only filtering.
    Dim dataLo As ListObject: Set dataLo = lo(DataTableName())
    If dataLo Is Nothing Then Exit Sub

    Dim nm As String: nm = TempFilterColName()
    Dim lc As ListColumn
    On Error Resume Next
    For Each lc In dataLo.ListColumns
        If StrComp(CStr(lc.name), nm, vbTextCompare) = 0 Then
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

    Dim n As Long: n = dataLo.DataBodyRange.Rows.Count
    If n = 0 Then Exit Sub
    If UBound(includeMask) < 1 Then Exit Sub

    Dim nm As String: nm = TempFilterColName()
    Dim lc As ListColumn, colIdx As Long: colIdx = 0

    ' find or add helper column at the end
    For Each lc In dataLo.ListColumns
        If StrComp(CStr(lc.name), nm, vbTextCompare) = 0 Then colIdx = lc.Index: Exit For
    Next lc
    If colIdx = 0 Then
        Set lc = dataLo.ListColumns.Add
        lc.name = nm
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
    lc.Range.DataBodyRange.Value = v
    ' apply filter: keep 1s only (stacking with existing slicer/table filters)
    dataLo.Range.AutoFilter Field:=colIdx, Criteria1:=1
    lc.Range.EntireColumn.Hidden = True
    Application.EnableEvents = True
End Sub



' -------------------------------------------------------------
' SECTION 6 � Slicer pulse (no hardcoded addresses anymore)
' -------------------------------------------------------------
Public Sub EnsurePulseCell_Run() 'For manual testing
    EnsurePulseCell True
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
    hdrName = Trim$(GetConfigValue("SLICER_PULSE_HEADER"))   ' e.g., "SAP ID"
    If Len(hdrName) > 0 Then
        idx = HeaderIndexByText(dataLo, hdrName)
    End If
    If idx < 1 Or idx > dataLo.ListColumns.Count Then
        idx = 1 ' safe fallback to first column
    End If

    ' --- Where should the pulse live? (anchor) ---
    anchorName = Trim$(GetConfigValue("SLICER_PULSE_ANCHOR")) ' e.g., "SlicerPulseAnchor"
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
        NameOrUpdate "SlicerPulseCell", "=" & pulseCell.Address(True, True, xlA1, True)
        NameOrUpdate "SlicerPulse", "=" & pulseCell.Address(True, True, xlA1, True)
    Application.EnableEvents = True

    If recalcNow Then Application.Calculate

    
End Sub


' -------------------------------------------------------------
' SECTION 7 � Gating: when to show vs hide results
' -------------------------------------------------------------
Public Sub RefreshResults()
Dim TempDiagToggle As Boolean
TempDiagToggle = False
If TempDiagToggle = False Then
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo CleanExit
    
    EnsurePulseCell False 'idempotent; guarantees SlicerPulse exists
    
    Dim anyActive As Boolean: anyActive = IsAnySearchInputActive()
    Dim pulseOk As Boolean
    Dim pulseVal As Long: pulseVal = CurrentSlicerPulse()
    pulseOk = (pulseVal > 0 And pulseVal <= SlicerThreshold)

    If anyActive Then
        ' Search-driven: filter DataTable to matches (so slicers reflect results), then output
        PerformSearch
    ElseIf pulseOk Then
        ' Slicer-driven (no search inputs): make sure temp search filter is cleared, then show visible
        ClearTempSearchFilter
        OutputAllVisible
    Else
        ' Default: headers only
        ClearTempSearchFilter
        OutputNoResults
    End If



Else
'=== TEMP ONLY: force-show visible rows to confirm output works ===

    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo CleanExit

    ' COMMENT OUT normal gating to isolate output path:
    'Dim anyActive As Boolean: anyActive = IsAnySearchInputActive()
    'Dim pulseOk As Boolean
    'Dim pulseVal As Long: pulseVal = CurrentSlicerPulse()
    'pulseOk = (pulseVal > 0 And pulseVal <= SlicerThreshold)

    ' TEMP: ALWAYS SHOW VISIBLE
    ClearTempSearchFilter
    OutputAllVisible
End If
CleanExit:
    gBusy = False

End Sub


Public Sub Safe_PerformSearch()
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo CleanExit

    If IsAnySearchInputActive() Then
        PerformSearch
    Else
        ClearTempSearchFilter
        OutputNoResults
    End If

CleanExit:
    gBusy = False
End Sub

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
    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount: hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value): Next i
    resultsStart.Resize(1, colCount).Value = hdr
    ClearOldResults resultsStart, colCount

    ' Visible indexes under current slicers/filters
    Dim idxs As Variant: idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        WriteStatus statusRng, "No visible rows (check slicers/filters).", ""
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

    If Not statusRng Is Nothing Then WriteStatus statusRng, "Showing " & cap & " visible row(s).", "All Visible"
    Exit Sub
EH:
    LogErrorLocal "OutputAllVisible", Err.Number, Err.Description
End Sub



Public Sub Safe_OutputAllVisible()
    If gBusy Then Exit Sub
    gBusy = True
    On Error GoTo CleanExit

    If IsAnySearchInputActive() Then
        ' If an input is active, "Show All" should still respect search; keep filter and output
        PerformSearch
    Else
        ClearTempSearchFilter
        OutputAllVisible
    End If

CleanExit:
    gBusy = False
End Sub

' ----------------------------------------------------------------
' SECTION 8   PerformSearch (filters table so slicers match)
' ----------------------------------------------------------------
Public Sub PerformSearch()
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
        MsgBox "Named range 'ResultsStartCell' not found.", vbExclamation
        Exit Sub
    End If

    ' --- Read inputs (still using current two core inputs for filtering logic) ---
    Dim searchTxt As String: searchTxt = ReadSearchText_Description()
    Dim valveTxt As String:  valveTxt = ReadSearchText_Valve()

    ' Build description regex array (synonym-aware)
    Dim rxArr As Variant
    If Len(Trim$(searchTxt)) > 0 Then
        Dim mapLo As ListObject: Set mapLo = lo(MappingTableName())
        Dim synIndex As Object: Set synIndex = BuildSynonymIndex(mapLo)
        rxArr = BuildSearchRegexes(searchTxt, synIndex)
    Else
        rxArr = Array()
    End If
    Dim descActive As Boolean: descActive = IsArrayNonEmpty(rxArr)

    Dim valveColIdx As Long: valveColIdx = HeaderIndexByText(dataLo, "Valve Number")
    Dim valveActive As Boolean: valveActive = (Len(Trim$(valveTxt)) > 0 And valveColIdx > 0)

    ' Resolve the description column (for matching & default sort)
    Dim descColIdx As Long: descColIdx = GetColumnIndex(DataDescConfigKey(), dataLo)
    If descColIdx = 0 And descActive Then
        ' Fallback: if not configured, use first column from Out_Column1
        descColIdx = HeaderIndexByText(dataLo, GetConfigValue("Out_Column1"))
    End If

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

    ' Write headers & clear old results with the right width
    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount: hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value): Next i
    resultsStart.Resize(1, colCount).Value = hdr
    ClearOldResults resultsStart, colCount

    ' Consider only rows visible under current slicers/filters
    Dim idxs As Variant: idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        WriteStatus statusRng, "No visible rows (check slicers/filters).", ""
        Exit Sub
    End If

    ' Evaluate matches & build a mask over ALL rows (so we can filter the table itself)
    Dim nRows As Long: nRows = dataLo.DataBodyRange.Rows.Count
    Dim keepMask() As Boolean: ReDim keepMask(1 To nRows)

    Dim kept As Long, ri As Long
    Dim descText As String, valveText As String, keep As Boolean

    For i = 1 To 1000 'UBound(idxs)
        ri = idxs(i)                      ' 1-based within DataBodyRange
        keep = True


'----------------------------------------------------------------------------------------------------------


    ' --- TEST 1: is ri relative to DataBodyRange? (must be 1..Rows.Count) ---
    Debug.Print "Probe1 rowIndex:", ri, "rowsInData:", lo(DataTableName()).DataBodyRange.Rows.Count
    Debug.Print lo(DataTableName()).DataBodyRange.Cells(ri, descColIdx).Address   ' <-- should NOT error
    Debug.Print lo(DataTableName()).DataBodyRange.Cells(ri, descColIdx).Value

    ' --- TEST 2: does the current regex actually match this row's description? ---
    Debug.Print "SearchTxt:", "[" & searchTxt & "]"
    Debug.Print "Pattern0:", rxArr(0).pattern, "IgnoreCase:", rxArr(0).IgnoreCase
    Debug.Print "Test0:", rxArr(0).Test( _
        CStr(lo(DataTableName()).DataBodyRange.Cells(ri, descColIdx).Value) _
    )

    'Exit For ' run just one row for clarity
'Next i


'----------------------------------------------------------------------------------------------------------



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
    Next i

    If kept = 0 Then
        ' No matches: clear temp filter so slicers show their state, then message
        ClearTempSearchFilter
        WriteStatus statusRng, "Found 0 results.", "Query: " & Trim$(searchTxt)
        Exit Sub
    End If

    ' Apply DataTable filter so slicers mirror the narrowed set
    ApplyTempSearchFilter keepMask

    ' After filtering, recompute visible rows and render output (sorted by description if present)
    idxs = VisibleRowIndexes(dataLo)
    If IsEmpty(idxs) Then
        WriteStatus statusRng, "Found 0 results after filter.", ""
        Exit Sub
    End If

    Dim maxRows As Long: maxRows = MaxOutputRows()
    Dim cap As Long: cap = IIf(maxRows > 0, WorksheetFunction.Min(UBound(idxs), maxRows), UBound(idxs))

    Dim outArr() As Variant: ReDim outArr(1 To cap, 1 To colCount)
    For i = 1 To cap
        ri = idxs(i)
        For j = 1 To colCount
            outArr(i, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).Value)
        Next j
    Next i

    ' Sort result array by description column if present among outputs
    Dim descOutPos As Long: descOutPos = FindOutputPosForDataColumn(outCols, colCount, descColIdx)
    If descOutPos > 0 And cap > 1 Then QuickSort2D_N outArr, 1, cap, descOutPos

    resultsStart.Offset(1, 0).Resize(cap, colCount).Value = outArr
    WriteStatus statusRng, "Found " & cap & " row(s).", "Query: " & Trim$(searchTxt)
    Exit Sub
EH:
    LogErrorLocal "PerformSearch", Err.Number, Err.Description
End Sub


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

    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount: hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value): Next i
    resultsStart.Resize(1, colCount).Value = hdr
    ClearOldResults resultsStart, colCount   ' clears previous rows beneath headers

    If Not statusRng Is Nothing Then WriteStatus statusRng, "No results (gated by inputs/threshold).", ""
    Exit Sub
EH:
    LogErrorLocal "OutputNoResults", Err.Number, Err.Description
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

' Validate indices exist
Public Function IndicesValid(ByVal dataLo As ListObject, _
ByVal idx1 As Long, ByVal idx2 As Long, ByVal idx3 As Long) As Boolean
    Dim n As Long: n = dataLo.ListColumns.Count
    If idx1 < 1 Or idx1 > n Or idx2 < 1 Or idx2 > n Or idx3 < 1 Or idx3 > n Then
        MsgBox "Config column index/indices exceed the DataTable column count.", vbExclamation
        IndicesValid = False
    Else
        IndicesValid = True
    End If
End Function

Private Function lo(ByVal name As String) As ListObject
    Dim ws As Worksheet, l As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each l In ws.ListObjects
            If StrComp(l.name, name, vbTextCompare) = 0 Then Set lo = l: Exit Function
        Next l
    Next ws
End Function

Private Function nr(ByVal nm As String) As Range
    On Error Resume Next
    Set nr = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
End Function





Public Function ReadLeftCell(ByVal nm As String) As String
    On Error Resume Next
    Dim rng As Range: Set rng = nr(nm)
    If Not rng Is Nothing Then ReadLeftCell = CStr(rng.Cells(1, 1).Value)
End Function

Private Function SafeCellText(ByVal v As Variant) As String
    If IsError(v) Then SafeCellText = "" Else SafeCellText = CStr(v)
End Function

Private Function HeaderIndexByText(ByVal dataLo As ListObject, ByVal headerText As String) As Long
    Dim i As Long
    If dataLo Is Nothing Then Exit Function
    If Len(Trim$(headerText)) = 0 Then Exit Function
    For i = 1 To dataLo.ListColumns.Count
        If StrComp(CStr(dataLo.HeaderRowRange.Cells(1, i).Value), headerText, vbTextCompare) = 0 Then
            HeaderIndexByText = i: Exit Function
        End If
    Next i
End Function

Public Sub WriteHeaders(ByVal startCell As Range, ByVal dataLo As ListObject, _
                         ByVal codeIdx As Long, ByVal sortIdx As Long, ByVal searchIdx As Long)
    ' Legacy 3-col header writer kept for compatibility (unused in new dynamic flow)
    Dim hdr(1 To 1, 1 To 3) As Variant
    hdr(1, 1) = CStr(dataLo.HeaderRowRange.Cells(1, codeIdx).Value)
    hdr(1, 2) = CStr(dataLo.HeaderRowRange.Cells(1, sortIdx).Value)
    hdr(1, 3) = CStr(dataLo.HeaderRowRange.Cells(1, searchIdx).Value)
    startCell.Resize(1, 3).Value = hdr
End Sub

Private Sub ClearOldResults(ByVal startCell As Range, Optional ByVal colCount As Long = 3)
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
    Dim n As name
    
    If Len(refersTo) > 0 Then
        If Left$(refersTo, 1) <> "=" Then refersTo = "=" & refersTo
    End If
    
    On Error Resume Next
       Set n = ThisWorkbook.Names(nm)
    On Error GoTo 0
    
    If n Is Nothing Then
        ThisWorkbook.Names.Add name:=nm, refersTo:=refersTo
    Else
        n.refersTo = refersTo
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
        raw = LCase(Trim$(CStr(r.Cells(1, 1).Value)))
        std = LCase(Trim$(CStr(r.Cells(1, 2).Value)))
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
        If StrComp(Trim$(CStr(c.Value)), headerName, vbTextCompare) = 0 Then col = c.Column: Exit For
    Next c
    If col = 0 Then GoTo Clean
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Dim r As Long, v As String
    For r = 2 To lastRow
        v = Trim$(UCase$(CStr(ws.Cells(r, col).Value)))
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
    If ms Is Nothing Or ms.Count = 0 Then Exit Function
    For Each m In ms
        v = m.Value
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
Exit Sub 'Temperary toggle off
    On Error Resume Next
    Application.ScreenUpdating = False
    Set logSheet = ThisWorkbook.Worksheets("SearchErrorLog")
    logSheet.Visible = xlSheetHidden
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add
        logSheet.name = "SearchErrorLog"
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
 
   ' MsgBox "Unexpected error in " & procName & "." & vbCrLf & "Logged to worksheet: SearchErrorLog", vbExclamation
End Sub


Public Function SafeReadNameLocal(ByVal nm As String) As String
On Error Resume Next
Dim r As Range: Set r = nr(nm)
If Not r Is Nothing Then SafeReadNameLocal = CStr(r.Cells(1, 1).Value)
End Function


'============================================================
' CONFIG ACCESS (from ConfigSheet/ConfigTable)
'============================================================
Private Function GetConfigValue(ByVal key As String) As String ' Demoted to Private; canonical public version in mod_SearchEngine_Enhanced
    Dim ws As Worksheet, loCfg As ListObject, r As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Or loCfg.DataBodyRange Is Nothing Then Exit Function
    On Error GoTo 0
    For Each r In loCfg.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), key, vbTextCompare) = 0 Then   ''???
            GetConfigValue = CStr(r.Cells(1, 2).Value) ''???
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
            If StrComp(tbl.name, DataTableName(), vbTextCompare) = 0 Then
                If Not tbl.AutoFilter Is Nothing Then
                    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
                End If
                ' Refresh results to show no results after clearing all filters and search boxes
                RefreshResults
                
                Exit Sub
            End If
        Next tbl
    Next ws

    
    'MsgBox "Table '" & DATA_TABLE_NAME & "' not found.", vbExclamation
End Sub

Public Sub ClearSearchBoxes()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim nm As name
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




Public Sub CloseWorkbookWithoutPrompt()
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub



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



