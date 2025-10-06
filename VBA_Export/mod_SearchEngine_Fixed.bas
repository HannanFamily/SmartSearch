'Attribute VB_Name = "mod_SearchEngine_Fixed"  ' commented for copy/paste
'============================================================
' SMART SEARCH ENGINE (Consolidated & Fixed)
'============================================================
' Purpose: Unified search engine with extensible mode support
' Features: 
'   - Mode-driven architecture for easy extension
'   - Robust error handling and logging
'   - Config-driven approach with sensible defaults
'   - Support for multiple search types (text, exact, regex)
'   - Integrated slicer and filter management
'============================================================

Option Explicit

' Search Engine Constants
Private Const DEFAULT_DATA_TABLE As String = "EquipmentData"
Private Const DEFAULT_CONFIG_SHEET As String = "ConfigSheet"
Private Const DEFAULT_CONFIG_TABLE As String = "ConfigTable"
Private Const DEFAULT_MODE_TABLE As String = "ModeConfigTable"
Private Const DEFAULT_RESULTS_RANGE As String = "A1"
Private Const MAX_RESULTS_DEFAULT As Long = 1000

' ConfigTable Keys (matching original system)
Private Const CFG_DASHBOARD_SHEET As String = "DASHBOARD_SHEET"
Private Const CFG_DATA_TABLE_NAME As String = "DATA_TABLE_NAME"
Private Const CFG_MAPPING_TABLE_NAME As String = "MAPPING_TABLE_NAME"
Private Const CFG_SPLICER_THRESHOLD As String = "SPLICER_THRESHOLD"
Private Const CFG_TEMP_FILTER_COL_NAME As String = "TEMP_FILTER_COL_NAME"
Private Const CFG_SLICER_PULSE_ANCHOR As String = "SLICER_PULSE_ANCHOR"
Private Const CFG_INPUT_DESCRIP As String = "InputCell_DescripSearch"
Private Const CFG_INPUT_VALVE As String = "InputCell_ValveNumSearch"
Private Const CFG_RESULTS_START As String = "ResultsStartCell"
Private Const CFG_STATUS_CELL As String = "StatusCell"
Private Const CFG_DATA_DESC_COL As String = "DataTable_EquipDescription"
Private Const CFG_OUT_COLUMN1 As String = "Out_Column1"
Private Const CFG_OUT_COLUMN2 As String = "Out_Column2"
Private Const CFG_OUT_COLUMN3 As String = "Out_Column3"
Private Const CFG_OUT_COLUMN4 As String = "Out_Column4"
Private Const CFG_OUT_COLUMN5 As String = "Out_Column5"
Private Const CFG_OUT_COLUMN6 As String = "Out_Column6"
Private Const CFG_OUT_COLUMN7 As String = "Out_Column7"
Private Const CFG_OUT_COLUMN8 As String = "Out_Column8"

' Search Mode Structure
Public Type SearchMode
    ModeName As String
    SourceTable As String
    FilterFormula As String
    OutputColumns As String
    SortColumn As String
    MaxResults As Long
    IsActive As Boolean
End Type

' Search Criteria Structure
Public Type SearchCriteria
    SearchText As String
    ExactMatch As Boolean
    CaseSensitive As Boolean
    SearchColumns As String
    CustomFilters As String
End Type

' Global state management
Private g_LastError As String
Private g_IsInitialized As Boolean

'============================================================
' PUBLIC API - MAIN SEARCH FUNCTIONS
'============================================================

' Main search entry point - handles all search types
Public Function ExecuteSearch(Optional modeName As String = "", Optional criteria As SearchCriteria) As Long
    On Error GoTo ErrorHandler
    
    ClearLastError
    
    ' Initialize if needed
    If Not g_IsInitialized Then
        If Not InitializeSearchEngine() Then
            ExecuteSearch = -1
            Exit Function
        End If
    End If
    
    ' Determine search mode
    If Len(Trim(modeName)) = 0 Then
        modeName = GetActiveMode()
    End If
    
    ' Load mode configuration
    Dim mode As SearchMode
    If Not LoadSearchMode(modeName, mode) Then
        SetLastError "Search mode '" & modeName & "' not found or invalid"
        ExecuteSearch = -1
        Exit Function
    End If
    
    ' Execute the search
    Dim resultCount As Long
    resultCount = PerformSearch(mode, criteria)
    
    ExecuteSearch = resultCount
    LogMessage "Search completed: " & resultCount & " results for mode '" & modeName & "'"
    Exit Function
    
ErrorHandler:
    SetLastError "ExecuteSearch: " & Err.Number & " - " & Err.Description
    ExecuteSearch = -1
End Function

' Quick text search with default settings
Public Function QuickSearch(searchText As String, Optional exactMatch As Boolean = False) As Long
    Dim criteria As SearchCriteria
    criteria.SearchText = searchText
    criteria.ExactMatch = exactMatch
    criteria.CaseSensitive = False
    QuickSearch = ExecuteSearch("", criteria)
End Function

' Refresh all visible results (for slicer changes)
Public Sub RefreshVisible()
    On Error GoTo ErrorHandler
    
    Dim mode As SearchMode
    If LoadSearchMode(GetActiveMode(), mode) Then
        Dim criteria As SearchCriteria
        ' Empty criteria means show all visible
        Call PerformSearch(mode, criteria)
    End If
    Exit Sub
    
ErrorHandler:
    LogMessage "RefreshVisible Error: " & Err.Number & " - " & Err.Description
End Sub

' Clear all search results and filters
Public Sub ClearSearch()
    On Error Resume Next
    
    ' Get data table
    Dim dataTable As ListObject
    Set dataTable = GetListObject(GetDataTableName())
    
    ' Clear search inputs
    ClearSearchInputs
    
    ' Clear temporary search filter (preserves slicer state)
    If Not dataTable Is Nothing Then
        Call ClearTempSearchFilter(dataTable)
    End If
    
    ' Clear data table filters (but preserve slicers)
    ClearTableFilters GetDataTableName()
    
    ' Clear results display
    ClearResultsDisplay
    
    ' Option: Reset slicers (comment out if you want to preserve slicer state)
    ' ResetSlicers GetDataTableName()
    
    LogMessage "Search cleared (slicers preserved)"
End Sub

'============================================================
' CORE SEARCH LOGIC
'============================================================

Private Function PerformSearch(mode As SearchMode, criteria As SearchCriteria) As Long
    On Error GoTo ErrorHandler
    
    ' Get source data
    Dim dataTable As ListObject
    Set dataTable = GetListObject(mode.SourceTable)
    If dataTable Is Nothing Then
        SetLastError "Source table '" & mode.SourceTable & "' not found"
        Exit Function
    End If
    
    ' Ensure slicer pulse is set up
    Call EnsureSlicerPulse(dataTable)
    
    ' If no search criteria, show all visible (slicer-filtered) rows
    If Len(Trim(criteria.SearchText)) = 0 And Len(Trim(mode.FilterFormula)) = 0 Then
        ' Clear any existing search filter to show slicer state only
        Call ClearTempSearchFilter(dataTable)
        Dim visibleRows As Collection
        Set visibleRows = GetVisibleRows(dataTable)
        Dim outputCount As Long
        outputCount = OutputResults(mode, dataTable, visibleRows)
        PerformSearch = outputCount
        Exit Function
    End If
    
    ' Get currently visible rows (respects slicer state)
    Dim slicerVisibleRows As Collection
    Set slicerVisibleRows = GetVisibleRows(dataTable)
    
    ' Apply search criteria to slicer-visible rows only
    Dim searchMatches As Collection
    Set searchMatches = FilterSearchResults(dataTable, slicerVisibleRows, criteria, mode)
    
    ' Create mask for temporary filter column
    Dim includeMask() As Boolean
    ReDim includeMask(1 To dataTable.ListRows.Count)
    
    Dim i As Long, rowIndex As Long
    For i = 1 To searchMatches.Count
        rowIndex = searchMatches(i)
        If rowIndex >= 1 And rowIndex <= UBound(includeMask) Then
            includeMask(rowIndex) = True
        End If
    Next i
    
    ' Apply temporary search filter (preserves slicer filtering)
    Call ApplyTempSearchFilter(dataTable, includeMask)
    
    ' Get final visible rows after both slicer and search filtering
    Dim finalRows As Collection
    Set finalRows = GetVisibleRows(dataTable)
    
    ' Output results
    outputCount = OutputResults(mode, dataTable, finalRows)
    PerformSearch = outputCount
    Exit Function
    
ErrorHandler:
    SetLastError "PerformSearch: " & Err.Number & " - " & Err.Description
    PerformSearch = -1
End Function

Private Function FilterSearchResults(dataTable As ListObject, visibleRows As Collection, criteria As SearchCriteria, mode As SearchMode) As Collection
    On Error GoTo ErrorHandler
    
    Dim matches As New Collection
    Dim i As Long, rowIndex As Long
    
    ' Filter visible rows based on search criteria
    For i = 1 To visibleRows.Count
        rowIndex = visibleRows(i)
        
        If EvaluateRowMatch(dataTable.ListRows(rowIndex).Range, mode.FilterFormula, criteria, dataTable) Then
            matches.Add rowIndex
        End If
    Next i
    
    Set FilterSearchResults = matches
    Exit Function
    
ErrorHandler:
    SetLastError "FilterSearchResults: " & Err.Number & " - " & Err.Description
    Set FilterSearchResults = New Collection
End Function

Private Function FilterDataTable(dataTable As ListObject, filterCondition As String, criteria As SearchCriteria) As Collection
    On Error GoTo ErrorHandler
    
    Dim matches As New Collection
    Dim rowData As Range
    Dim i As Long
    
    ' If no specific search criteria, return visible rows
    If Len(Trim(criteria.SearchText)) = 0 And Len(Trim(filterCondition)) = 0 Then
        Set matches = GetVisibleRows(dataTable)
        Set FilterDataTable = matches
        Exit Function
    End If
    
    ' Process each data row
    For i = 1 To dataTable.ListRows.Count
        Set rowData = dataTable.ListRows(i).Range
        
        If EvaluateRowMatch(rowData, filterCondition, criteria, dataTable) Then
            matches.Add i
        End If
    Next i
    
    Set FilterDataTable = matches
    Exit Function
    
ErrorHandler:
    SetLastError "FilterDataTable: " & Err.Number & " - " & Err.Description
    Set FilterDataTable = New Collection
End Function

Private Function EvaluateRowMatch(rowData As Range, filterCondition As String, criteria As SearchCriteria, dataTable As ListObject) As Boolean
    On Error GoTo ErrorHandler
    
    ' Start with true, apply filters as AND conditions
    Dim isMatch As Boolean
    isMatch = True
    
    ' Apply text search if specified
    If Len(Trim(criteria.SearchText)) > 0 Then
        isMatch = isMatch And EvaluateTextSearch(rowData, criteria, dataTable)
    End If
    
    ' Apply formula filter if specified
    If Len(Trim(filterCondition)) > 0 Then
        isMatch = isMatch And EvaluateFormulaFilter(rowData, filterCondition, dataTable)
    End If
    
    EvaluateRowMatch = isMatch
    Exit Function
    
ErrorHandler:
    EvaluateRowMatch = False
End Function

Private Function EvaluateTextSearch(rowData As Range, criteria As SearchCriteria, dataTable As ListObject) As Boolean
    On Error GoTo ErrorHandler
    
    Dim searchText As String
    searchText = criteria.SearchText
    
    Dim searchColumns As Variant
    If Len(Trim(criteria.SearchColumns)) > 0 Then
        searchColumns = Split(criteria.SearchColumns, ",")
    Else
        ' Default to searching all text columns
        searchColumns = GetTextColumns(dataTable)
    End If
    
    Dim i As Long, colIndex As Long
    Dim cellValue As String
    Dim compareMode As VbCompareMethod
    compareMode = IIf(criteria.CaseSensitive, vbBinaryCompare, vbTextCompare)
    
    ' Search in specified columns
    For i = 0 To UBound(searchColumns)
        colIndex = GetColumnIndex(dataTable, Trim(searchColumns(i)))
        If colIndex > 0 Then
            cellValue = CStr(rowData.Cells(1, colIndex).Value)
            
            If criteria.ExactMatch Then
                If StrComp(cellValue, searchText, compareMode) = 0 Then
                    EvaluateTextSearch = True
                    Exit Function
                End If
            Else
                If InStr(1, cellValue, searchText, compareMode) > 0 Then
                    EvaluateTextSearch = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    EvaluateTextSearch = False
    Exit Function
    
ErrorHandler:
    EvaluateTextSearch = False
End Function

Private Function EvaluateFormulaFilter(rowData As Range, filterFormula As String, dataTable As ListObject) As Boolean
    On Error GoTo ErrorHandler
    
    ' Replace column references in formula with actual values
    Dim evaluableFormula As String
    evaluableFormula = ReplaceColumnReferences(filterFormula, rowData, dataTable)
    
    ' Evaluate the formula
    Dim result As Variant
    result = Application.Evaluate(evaluableFormula)
    
    EvaluateFormulaFilter = CBool(result)
    Exit Function
    
ErrorHandler:
    EvaluateFormulaFilter = False
End Function

'============================================================
' OUTPUT AND DISPLAY
'============================================================

Private Function OutputResults(mode As SearchMode, dataTable As ListObject, matchedRows As Collection) As Long
    On Error GoTo ErrorHandler
    
    ' Get output columns from ConfigTable or mode definition
    Dim outputCols As Variant
    If Len(Trim(mode.OutputColumns)) > 0 Then
        outputCols = Split(mode.OutputColumns, ",")
    Else
        outputCols = GetOutputColumns()  ' From ConfigTable Out_Column1..8
    End If
    
    ' Get output range from ConfigTable
    Dim outputRange As Range
    Set outputRange = GetOutputRange()
    If outputRange Is Nothing Then
        SetLastError "Output range not found"
        Exit Function
    End If
    
    ' Clear previous results
    Call ClearResultsDisplay()
    
    ' Write headers using config-driven column names
    Call WriteConfigHeaders(outputRange, outputCols, dataTable)
    
    ' Determine how many rows to output
    Dim maxRows As Long
    maxRows = IIf(mode.MaxResults > 0, mode.MaxResults, MAX_RESULTS_DEFAULT)
    Dim outputRows As Long
    outputRows = Application.WorksheetFunction.Min(matchedRows.Count, maxRows)
    
    ' Write data
    If outputRows > 0 Then
        Call WriteConfigDataRows(outputRange.Offset(1, 0), outputCols, dataTable, matchedRows, outputRows)
    End If
    
    ' Update status using ConfigTable status cell
    Call UpdateConfigStatus(outputRows, matchedRows.Count, mode.ModeName)
    
    OutputResults = outputRows
    Exit Function
    
ErrorHandler:
    SetLastError "OutputResults: " & Err.Number & " - " & Err.Description
    OutputResults = -1
End Function

Private Sub WriteConfigHeaders(outputRange As Range, outputCols As Variant, dataTable As ListObject)
    On Error Resume Next
    
    Dim i As Long
    For i = 0 To UBound(outputCols)
        Dim colName As String
        colName = Trim(outputCols(i))
        outputRange.Offset(0, i).Value = colName
    Next i
End Sub

Private Sub WriteConfigDataRows(outputRange As Range, outputCols As Variant, dataTable As ListObject, matchedRows As Collection, rowCount As Long)
    On Error Resume Next
    
    Dim i As Long, j As Long
    Dim rowIndex As Long, colIndex As Long
    
    For i = 1 To Application.WorksheetFunction.Min(rowCount, matchedRows.Count)
        rowIndex = matchedRows(i)
        
        For j = 0 To UBound(outputCols)
            colIndex = GetColumnIndex(dataTable, Trim(outputCols(j)))
            If colIndex > 0 Then
                outputRange.Offset(i - 1, j).Value = dataTable.ListRows(rowIndex).Range.Cells(1, colIndex).Value
            End If
        Next j
    Next i
End Sub

Private Sub UpdateConfigStatus(displayedRows As Long, totalMatches As Long, modeName As String)
    On Error Resume Next
    
    ' Get status cell from ConfigTable
    Dim statusRange As Range
    Set statusRange = GetNamedRange(GetConfigValue(CFG_STATUS_CELL))
    If statusRange Is Nothing Then
        Set statusRange = GetNamedRange("SearchStatus")
    End If
    
    If Not statusRange Is Nothing Then
        Dim statusText As String
        statusText = "Found " & totalMatches & " matches, showing " & displayedRows & " rows (Mode: " & modeName & ")"
        statusRange.Value = statusText
    End If
End Sub

Private Sub WriteHeaders(outputRange As Range, outputCols As Variant, dataTable As ListObject)
    On Error Resume Next
    
    Dim i As Long
    For i = 0 To UBound(outputCols)
        Dim colName As String
        colName = Trim(outputCols(i))
        outputRange.Offset(0, i).Value = colName
    Next i
End Sub

Private Sub WriteDataRows(outputRange As Range, outputCols As Variant, dataTable As ListObject, matchedRows As Collection, rowCount As Long)
    On Error Resume Next
    
    Dim i As Long, j As Long
    Dim rowIndex As Long, colIndex As Long
    
    For i = 1 To Application.WorksheetFunction.Min(rowCount, matchedRows.Count)
        rowIndex = matchedRows(i)
        
        For j = 0 To UBound(outputCols)
            colIndex = GetColumnIndex(dataTable, Trim(outputCols(j)))
            If colIndex > 0 Then
                outputRange.Offset(i - 1, j).Value = dataTable.ListRows(rowIndex).Range.Cells(1, colIndex).Value
            End If
        Next j
    Next i
End Sub

'============================================================
' MODE MANAGEMENT
'============================================================

Public Function RegisterSearchMode(modeName As String, sourceTable As String, outputColumns As String, Optional filterFormula As String = "", Optional sortColumn As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim modeTable As ListObject
    Set modeTable = GetModeConfigTable()
    If modeTable Is Nothing Then
        ' Create mode table if it doesn't exist
        Set modeTable = CreateModeConfigTable()
        If modeTable Is Nothing Then
            SetLastError "Could not create mode configuration table"
            Exit Function
        End If
    End If
    
    ' Check if mode already exists
    If ModeExists(modeName) Then
        ' Update existing mode
        UpdateSearchMode modeName, sourceTable, outputColumns, filterFormula, sortColumn
    Else
        ' Add new mode
        AddNewSearchMode modeTable, modeName, sourceTable, outputColumns, filterFormula, sortColumn
    End If
    
    RegisterSearchMode = True
    LogMessage "Registered search mode: " & modeName
    Exit Function
    
ErrorHandler:
    SetLastError "RegisterSearchMode: " & Err.Number & " - " & Err.Description
    RegisterSearchMode = False
End Function

Public Function GetAvailableModes() As Collection
    On Error GoTo ErrorHandler
    
    Dim modes As New Collection
    Dim modeTable As ListObject
    Set modeTable = GetModeConfigTable()
    
    If Not modeTable Is Nothing Then
        Dim i As Long
        For i = 1 To modeTable.ListRows.Count
            Dim modeName As String
            modeName = Trim(CStr(modeTable.ListRows(i).Range.Cells(1, 1).Value))
            If Len(modeName) > 0 Then
                modes.Add modeName
            End If
        Next i
    End If
    
    Set GetAvailableModes = modes
    Exit Function
    
ErrorHandler:
    Set GetAvailableModes = New Collection
End Function

'============================================================
' UTILITY FUNCTIONS
'============================================================

Private Function InitializeSearchEngine() As Boolean
    On Error GoTo ErrorHandler
    
    ' Verify required components exist
    If Not VerifyRequiredComponents() Then
        Exit Function
    End If
    
    ' Set up default modes if none exist
    SetupDefaultModes
    
    g_IsInitialized = True
    InitializeSearchEngine = True
    LogMessage "Search engine initialized"
    Exit Function
    
ErrorHandler:
    SetLastError "InitializeSearchEngine: " & Err.Number & " - " & Err.Description
    InitializeSearchEngine = False
End Function

Private Function VerifyRequiredComponents() As Boolean
    ' Check for data table
    If GetListObject(GetDataTableName()) Is Nothing Then
        SetLastError "Data table not found: " & GetDataTableName()
        Exit Function
    End If
    
    ' Other verifications can be added here
    VerifyRequiredComponents = True
End Function

Private Sub SetupDefaultModes()
    On Error Resume Next
    
    ' Register basic search mode if no modes exist
    Dim modes As Collection
    Set modes = GetAvailableModes()
    
    If modes.Count = 0 Then
        Call RegisterSearchMode("Default", GetDataTableName(), GetDefaultOutputColumns(GetListObject(GetDataTableName())))
    End If
End Sub

Private Function GetDataTableName() As String
    ' Try to get from config, fall back to default
    Dim tableName As String
    tableName = GetConfigValue(CFG_DATA_TABLE_NAME)
    If Len(Trim(tableName)) = 0 Then
        tableName = DEFAULT_DATA_TABLE
    End If
    GetDataTableName = tableName
End Function

Private Function GetDashboardSheet() As Worksheet
    On Error Resume Next
    Dim sheetName As String
    sheetName = GetConfigValue(CFG_DASHBOARD_SHEET)
    If Len(Trim(sheetName)) = 0 Then sheetName = "Dashboard"
    Set GetDashboardSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function GetSlicerThreshold() As Long
    Dim threshold As String
    threshold = GetConfigValue(CFG_SPLICER_THRESHOLD)
    GetSlicerThreshold = IIf(Len(Trim(threshold)) > 0, CLng(Val(threshold)), 250)
End Function

Private Function GetOutputColumns() As Variant
    ' Get output columns from ConfigTable (Out_Column1..8)
    Dim cols As Collection
    Set cols = New Collection
    
    Dim outKeys(1 To 8) As String
    outKeys(1) = CFG_OUT_COLUMN1: outKeys(2) = CFG_OUT_COLUMN2
    outKeys(3) = CFG_OUT_COLUMN3: outKeys(4) = CFG_OUT_COLUMN4
    outKeys(5) = CFG_OUT_COLUMN5: outKeys(6) = CFG_OUT_COLUMN6
    outKeys(7) = CFG_OUT_COLUMN7: outKeys(8) = CFG_OUT_COLUMN8
    
    Dim i As Long, colName As String
    For i = 1 To 8
        colName = GetConfigValue(outKeys(i))
        If Len(Trim(colName)) > 0 Then
            cols.Add colName
        End If
    Next i
    
    ' Convert to array
    If cols.Count > 0 Then
        Dim arr() As String
        ReDim arr(0 To cols.Count - 1)
        For i = 1 To cols.Count
            arr(i - 1) = cols(i)
        Next i
        GetOutputColumns = arr
    Else
        GetOutputColumns = Array("Column1", "Column2", "Column3")
    End If
End Function

Private Function CreateNamedRange(rangeName As String, refersTo As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim n As Name
    Set n = Nothing
    
    ' Try to find existing name
    On Error Resume Next
    Set n = ThisWorkbook.Names(rangeName)
    On Error GoTo ErrorHandler
    
    If Not refersTo Like "=*" Then refersTo = "=" & refersTo
    
    If n Is Nothing Then
        ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=refersTo
    Else
        n.RefersTo = refersTo
    End If
    
    CreateNamedRange = True
    Exit Function
    
ErrorHandler:
    CreateNamedRange = False
End Function

Private Function GetConfigValue(key As String) As String
    On Error Resume Next
    
    Dim configSheet As Worksheet
    Set configSheet = ThisWorkbook.Worksheets(DEFAULT_CONFIG_SHEET)
    If configSheet Is Nothing Then Exit Function
    
    Dim configTable As ListObject
    Set configTable = configSheet.ListObjects(DEFAULT_CONFIG_TABLE)
    If configTable Is Nothing Then Exit Function
    
    Dim i As Long
    For i = 1 To configTable.ListRows.Count
        If StrComp(CStr(configTable.ListRows(i).Range.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            GetConfigValue = CStr(configTable.ListRows(i).Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next i
End Function

Private Function GetListObject(tableName As String) As ListObject
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set GetListObject = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function GetColumnIndex(dataTable As ListObject, columnName As String) As Long
    On Error Resume Next
    
    Dim i As Long
    For i = 1 To dataTable.ListColumns.Count
        If StrComp(CStr(dataTable.HeaderRowRange.Cells(1, i).Value), columnName, vbTextCompare) = 0 Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function GetVisibleRows(dataTable As ListObject) As Collection
    On Error GoTo ErrorHandler
    
    Dim visibleRows As New Collection
    Dim i As Long
    
    ' Check if table has filters applied (including slicers)
    If dataTable.ShowAutoFilter And dataTable.AutoFilter.FilterMode Then
        ' Get visible rows only using SpecialCells
        Dim visibleRange As Range
        Set visibleRange = dataTable.ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible)
        
        Dim cell As Range
        For Each cell In visibleRange.Cells
            visibleRows.Add cell.Row - dataTable.DataBodyRange.Row + 1
        Next cell
    Else
        ' Check if slicers are filtering (even without AutoFilter)
        If IsSlicerFiltered(dataTable.Name) Then
            ' Use first column to detect visible rows under slicer filtering
            Dim firstCol As Range
            Set firstCol = dataTable.ListColumns(1).DataBodyRange
            
            On Error Resume Next
            Set visibleRange = firstCol.SpecialCells(xlCellTypeVisible)
            On Error GoTo ErrorHandler
            
            If Not visibleRange Is Nothing Then
                For Each cell In visibleRange.Cells
                    visibleRows.Add cell.Row - dataTable.DataBodyRange.Row + 1
                Next cell
            End If
        Else
            ' All rows are visible
            For i = 1 To dataTable.ListRows.Count
                visibleRows.Add i
            Next i
        End If
    End If
    
    Set GetVisibleRows = visibleRows
    Exit Function
    
ErrorHandler:
    Set GetVisibleRows = New Collection
End Function

'============================================================
' HELPER FUNCTIONS
'============================================================

Private Function LoadSearchMode(modeName As String, ByRef mode As SearchMode) As Boolean
    On Error GoTo ErrorHandler
    
    Dim modeTable As ListObject
    Set modeTable = GetModeConfigTable()
    If modeTable Is Nothing Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To modeTable.ListRows.Count
        If StrComp(CStr(modeTable.ListRows(i).Range.Cells(1, 1).Value), modeName, vbTextCompare) = 0 Then
            With mode
                .ModeName = CStr(modeTable.ListRows(i).Range.Cells(1, 1).Value)
                .SourceTable = CStr(modeTable.ListRows(i).Range.Cells(1, 2).Value)
                .OutputColumns = CStr(modeTable.ListRows(i).Range.Cells(1, 3).Value)
                .FilterFormula = CStr(modeTable.ListRows(i).Range.Cells(1, 4).Value)
                .SortColumn = CStr(modeTable.ListRows(i).Range.Cells(1, 5).Value)
                .MaxResults = CLng(Val(CStr(modeTable.ListRows(i).Range.Cells(1, 6).Value)))
                .IsActive = True
            End With
            LoadSearchMode = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    LoadSearchMode = False
End Function

Private Function GetActiveMode() As String
    ' Try to get from named range, default to first available mode
    On Error Resume Next
    
    Dim modeSelector As Range
    Set modeSelector = ThisWorkbook.Names("ModeSelector").RefersToRange
    If Not modeSelector Is Nothing Then
        GetActiveMode = CStr(modeSelector.Value)
        If Len(Trim(GetActiveMode)) > 0 Then Exit Function
    End If
    
    ' Fall back to first available mode
    Dim modes As Collection
    Set modes = GetAvailableModes()
    If modes.Count > 0 Then
        GetActiveMode = modes(1)
    Else
        GetActiveMode = "Default"
    End If
End Function

Private Function GetModeConfigTable() As ListObject
    Set GetModeConfigTable = GetListObject(DEFAULT_MODE_TABLE)
End Function

Private Function ModeExists(modeName As String) As Boolean
    Dim modes As Collection
    Set modes = GetAvailableModes()
    
    Dim i As Long
    For i = 1 To modes.Count
        If StrComp(CStr(modes(i)), modeName, vbTextCompare) = 0 Then
            ModeExists = True
            Exit Function
        End If
    Next i
End Function

Private Function GetDefaultOutputColumns(dataTable As ListObject) As String
    On Error Resume Next
    
    If dataTable Is Nothing Then
        GetDefaultOutputColumns = "Column1,Column2,Column3"
        Exit Function
    End If
    
    Dim cols As String
    Dim i As Long
    Dim maxCols As Long
    maxCols = Application.WorksheetFunction.Min(5, dataTable.ListColumns.Count)
    
    For i = 1 To maxCols
        If i > 1 Then cols = cols & ","
        cols = cols & CStr(dataTable.HeaderRowRange.Cells(1, i).Value)
    Next i
    
    GetDefaultOutputColumns = cols
End Function

Private Function GetTextColumns(dataTable As ListObject) As Variant
    On Error Resume Next
    
    ' For now, return all columns - could be enhanced to detect text columns
    Dim cols() As String
    ReDim cols(0 To dataTable.ListColumns.Count - 1)
    
    Dim i As Long
    For i = 1 To dataTable.ListColumns.Count
        cols(i - 1) = CStr(dataTable.HeaderRowRange.Cells(1, i).Value)
    Next i
    
    GetTextColumns = cols
End Function

Private Function BuildFilterCondition(mode As SearchMode, criteria As SearchCriteria, dataTable As ListObject) As String
    ' Build filter from mode and criteria
    Dim condition As String
    
    ' Start with mode filter formula
    If Len(Trim(mode.FilterFormula)) > 0 Then
        condition = mode.FilterFormula
    End If
    
    ' Add custom filters from criteria
    If Len(Trim(criteria.CustomFilters)) > 0 Then
        If Len(condition) > 0 Then
            condition = "(" & condition & ") AND (" & criteria.CustomFilters & ")"
        Else
            condition = criteria.CustomFilters
        End If
    End If
    
    BuildFilterCondition = condition
End Function

Private Function ReplaceColumnReferences(formula As String, rowData As Range, dataTable As ListObject) As String
    ' Replace [@ColumnName] references with actual cell values
    ' This is a simplified implementation
    ReplaceColumnReferences = formula
    
    ' TODO: Implement proper column reference replacement
    ' For now, return the formula as-is
End Function

Private Function GetOutputRange() As Range
    On Error Resume Next
    
    ' Try to get from ConfigTable first
    Dim rangeName As String
    rangeName = GetConfigValue(CFG_RESULTS_START)
    
    If Len(Trim(rangeName)) > 0 Then
        Set GetOutputRange = GetNamedRange(rangeName)
        If Not GetOutputRange Is Nothing Then Exit Function
    End If
    
    ' Fall back to common named ranges
    Dim fallbackNames As Variant
    fallbackNames = Array("ResultsStartCell", "ResultsStart", "SearchResults")
    
    Dim i As Long
    For i = 0 To UBound(fallbackNames)
        Set GetOutputRange = GetNamedRange(CStr(fallbackNames(i)))
        If Not GetOutputRange Is Nothing Then Exit Function
    Next i
    
    ' Final fallback to Dashboard A1
    Dim dashSheet As Worksheet
    Set dashSheet = GetDashboardSheet()
    If Not dashSheet Is Nothing Then
        Set GetOutputRange = dashSheet.Range("A1")
    End If
End Function

Private Function GetNamedRange(rangeName As String) As Range
    On Error Resume Next
    If Len(Trim(rangeName)) > 0 Then
        Set GetNamedRange = ThisWorkbook.Names(rangeName).RefersToRange
    End If
End Function

'============================================================
' STATE MANAGEMENT
'============================================================

Private Sub SetLastError(errorMsg As String)
    g_LastError = errorMsg
    LogMessage "ERROR: " & errorMsg
End Sub

Private Sub ClearLastError()
    g_LastError = ""
End Sub

Public Function GetLastError() As String
    GetLastError = g_LastError
End Function

Private Sub LogMessage(msg As String)
    ' Simple logging to debug window
    Debug.Print "[SearchEngine] " & Format(Now, "hh:mm:ss") & " " & msg
End Sub

'============================================================
' CLEANUP AND MAINTENANCE
'============================================================

Private Sub ClearSearchInputs()
    On Error Resume Next
    
    ' Clear inputs based on ConfigTable entries with TYPE = "Input Named Range"
    Dim configSheet As Worksheet
    Set configSheet = ThisWorkbook.Worksheets(DEFAULT_CONFIG_SHEET)
    If configSheet Is Nothing Then GoTo FallbackClear
    
    Dim configTable As ListObject
    Set configTable = configSheet.ListObjects(DEFAULT_CONFIG_TABLE)
    If configTable Is Nothing Then GoTo FallbackClear
    
    Dim r As Range, nmKey As String, nmType As String
    For Each r In configTable.DataBodyRange.Rows
        nmKey = Trim(CStr(r.Cells(1, 2).Value))   ' ConfigValue (range name)
        nmType = Trim(CStr(r.Cells(1, 3).Value))  ' TYPE
        
        If StrComp(nmType, "Input Named Range", vbTextCompare) = 0 Then
            Dim inputRange As Range
            Set inputRange = GetNamedRange(nmKey)
            If Not inputRange Is Nothing Then
                inputRange.ClearContents
            End If
        End If
    Next r
    
    Exit Sub
    
FallbackClear:
    ' Fallback: Clear common search input ranges
    Dim inputRanges As Variant
    inputRanges = Array("SearchBox", "ValveNumSearchBox", "DescriptionSearch", "TagSearch", _
                       GetConfigValue(CFG_INPUT_DESCRIP), GetConfigValue(CFG_INPUT_VALVE))
    
    Dim i As Long
    Dim rng As Range
    For i = 0 To UBound(inputRanges)
        If Len(Trim(CStr(inputRanges(i)))) > 0 Then
            Set rng = GetNamedRange(CStr(inputRanges(i)))
            If Not rng Is Nothing Then
                rng.ClearContents
            End If
        End If
    Next i
End Sub

Private Sub ClearTableFilters(tableName As String)
    On Error Resume Next
    
    Dim dataTable As ListObject
    Set dataTable = GetListObject(tableName)
    If Not dataTable Is Nothing Then
        If dataTable.ShowAutoFilter Then
            dataTable.AutoFilter.ShowAllData
        End If
    End If
End Sub

Private Sub ClearResultsDisplay()
    On Error Resume Next
    
    Dim outputRange As Range
    Set outputRange = GetOutputRange()
    If Not outputRange Is Nothing Then
        ' Clear a reasonable area (100 rows x 20 columns)
        outputRange.Resize(100, 20).ClearContents
    End If
End Sub

Private Sub ResetSlicers(tableName As String)
    On Error Resume Next
    
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If StrComp(sc.SourceName, tableName, vbTextCompare) = 0 Then
            sc.ClearManualFilter
        End If
    Next sc
End Sub

Private Function IsSlicerFiltered(tableName As String) As Boolean
    On Error Resume Next
    
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If StrComp(sc.SourceName, tableName, vbTextCompare) = 0 Then
            If sc.FilterCleared = False Then
                IsSlicerFiltered = True
                Exit Function
            End If
        End If
    Next sc
End Function

' Apply temporary search filter using helper column (preserves slicer state)
Private Sub ApplyTempSearchFilter(dataTable As ListObject, includeMask() As Boolean)
    On Error GoTo ErrorHandler
    
    If dataTable Is Nothing Or dataTable.DataBodyRange Is Nothing Then Exit Sub
    
    Dim n As Long
    n = dataTable.DataBodyRange.Rows.Count
    If n = 0 Or UBound(includeMask) < 1 Then Exit Sub
    
    Dim tempColName As String
    tempColName = GetConfigValue(CFG_TEMP_FILTER_COL_NAME)
    If Len(Trim(tempColName)) = 0 Then tempColName = "Temp_SearchInclude"
    
    Dim lc As ListColumn, colIdx As Long
    colIdx = 0
    
    ' Find or add helper column at the end
    For Each lc In dataTable.ListColumns
        If StrComp(CStr(lc.Name), tempColName, vbTextCompare) = 0 Then
            colIdx = lc.Index
            Exit For
        End If
    Next lc
    
    If colIdx = 0 Then
        Set lc = dataTable.ListColumns.Add
        lc.Name = tempColName
        colIdx = lc.Index
    Else
        Set lc = dataTable.ListColumns(colIdx)
    End If
    
    ' Write filter values
    Dim v As Variant
    ReDim v(1 To n, 1 To 1)
    Dim i As Long
    For i = 1 To n
        v(i, 1) = IIf(i <= UBound(includeMask) And includeMask(i), 1, 0)
    Next i
    
    Application.EnableEvents = False
    lc.Range.DataBodyRange.Value = v
    
    ' Apply filter: keep 1s only (stacks with existing slicer/table filters)
    dataTable.Range.AutoFilter Field:=colIdx, Criteria1:=1
    lc.Range.EntireColumn.Hidden = True
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    LogMessage "ApplyTempSearchFilter Error: " & Err.Number & " - " & Err.Description
End Sub

' Clear temporary search filter (restores slicer-only state)
Private Sub ClearTempSearchFilter(dataTable As ListObject)
    On Error Resume Next
    
    If dataTable Is Nothing Then Exit Sub
    
    Dim tempColName As String
    tempColName = GetConfigValue(CFG_TEMP_FILTER_COL_NAME)
    If Len(Trim(tempColName)) = 0 Then tempColName = "Temp_SearchInclude"
    
    Dim lc As ListColumn
    For Each lc In dataTable.ListColumns
        If StrComp(CStr(lc.Name), tempColName, vbTextCompare) = 0 Then
            Application.EnableEvents = False
            lc.Delete
            Application.EnableEvents = True
            Exit For
        End If
    Next lc
End Sub

' Ensure slicer pulse cell exists and is properly configured
Private Sub EnsureSlicerPulse(dataTable As ListObject)
    On Error Resume Next
    
    If dataTable Is Nothing Then Exit Sub
    
    Dim dashSheet As Worksheet
    Set dashSheet = GetDashboardSheet()
    If dashSheet Is Nothing Then Exit Sub
    
    ' Get pulse anchor from config
    Dim anchorName As String
    anchorName = GetConfigValue(CFG_SLICER_PULSE_ANCHOR)
    If Len(Trim(anchorName)) = 0 Then anchorName = "SlicerPulseAnchor"
    
    Dim anchorRange As Range
    Set anchorRange = GetNamedRange(anchorName)
    If anchorRange Is Nothing Then
        ' Create default pulse cell
        Set anchorRange = dashSheet.Range("AA1")
        dashSheet.Columns("AA").Hidden = True
        CreateNamedRange "SlicerPulseAnchor", anchorRange.Address(True, True, xlA1, True)
    End If
    
    ' Set up SUBTOTAL formula to detect slicer changes
    Dim pulseFormula As String
    pulseFormula = "=SUBTOTAL(103," & dataTable.ListColumns(1).DataBodyRange.Address(True, True, xlA1, True) & ")"
    
    Application.EnableEvents = False
    anchorRange.Formula = pulseFormula
    CreateNamedRange "SlicerPulse", anchorRange.Address(True, True, xlA1, True)
    CreateNamedRange "SlicerPulseCell", anchorRange.Address(True, True, xlA1, True)
    Application.EnableEvents = True
End Sub

Private Sub UpdateSearchStatus(displayedRows As Long, totalMatches As Long, modeName As String)
    On Error Resume Next
    
    Dim statusRange As Range
    Set statusRange = ThisWorkbook.Names("SearchStatus").RefersToRange
    If Not statusRange Is Nothing Then
        Dim statusText As String
        statusText = "Found " & totalMatches & " matches, showing " & displayedRows & " rows (Mode: " & modeName & ")"
        statusRange.Value = statusText
    End If
End Sub

'============================================================
' ADVANCED FEATURES (Stubs for future enhancement)
'============================================================

Private Function CreateModeConfigTable() As ListObject
    ' Stub for creating mode config table if it doesn't exist
    ' TODO: Implement table creation
    Set CreateModeConfigTable = Nothing
End Function

Private Sub AddNewSearchMode(modeTable As ListObject, modeName As String, sourceTable As String, outputColumns As String, filterFormula As String, sortColumn As String)
    ' Stub for adding new search mode
    ' TODO: Implement mode addition
End Sub

Private Sub UpdateSearchMode(modeName As String, sourceTable As String, outputColumns As String, filterFormula As String, sortColumn As String)
    ' Stub for updating existing search mode
    ' TODO: Implement mode update
End Sub