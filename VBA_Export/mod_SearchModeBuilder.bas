'Attribute VB_Name = "mod_SearchModeBuilder"  ' commented for copy/paste
'============================================================
' SEARCH MODE BUILDER - Easy Mode Management
'============================================================
' Purpose: Simplified interface for creating and managing search modes
' Usage: Provides easy-to-use functions for adding new search capabilities
'============================================================

Option Explicit

'============================================================
' QUICK MODE SETUP FUNCTIONS
'============================================================

' Creates a simple text search mode
Public Function CreateTextSearchMode(modeName As String, tableName As String, searchColumns As String) As Boolean
    Dim outputCols As String
    outputCols = GetFirstNColumns(tableName, 5) ' Default to first 5 columns
    
    CreateTextSearchMode = RegisterSearchMode(modeName, tableName, outputCols)
    If CreateTextSearchMode Then
        LogModeBuilder "Created text search mode: " & modeName
    End If
End Function

' Creates an exact match search mode (for IDs, codes, etc.)
Public Function CreateExactMatchMode(modeName As String, tableName As String, targetColumn As String) As Boolean
    Dim filterFormula As String
    filterFormula = "[@" & targetColumn & "]=" & Chr(34) & "{SEARCH_VALUE}" & Chr(34)
    
    Dim outputCols As String
    outputCols = GetFirstNColumns(tableName, 5)
    
    CreateExactMatchMode = RegisterSearchMode(modeName, tableName, outputCols, filterFormula)
    If CreateExactMatchMode Then
        LogModeBuilder "Created exact match mode: " & modeName & " for column: " & targetColumn
    End If
End Function

' Creates a filtered view mode (shows subset based on criteria)
Public Function CreateFilteredViewMode(modeName As String, tableName As String, filterCriteria As String, displayColumns As String) As Boolean
    CreateFilteredViewMode = RegisterSearchMode(modeName, tableName, displayColumns, filterCriteria)
    If CreateFilteredViewMode Then
        LogModeBuilder "Created filtered view mode: " & modeName
    End If
End Function

' Creates a location-based search mode
Public Function CreateLocationSearchMode(modeName As String, tableName As String, locationColumn As String) As Boolean
    Dim filterFormula As String
    filterFormula = "[@" & locationColumn & "]=" & Chr(34) & "{LOCATION}" & Chr(34)
    
    Dim outputCols As String
    outputCols = GetLocationOutputColumns(tableName, locationColumn)
    
    CreateLocationSearchMode = RegisterSearchMode(modeName, tableName, outputCols, filterFormula)
    If CreateLocationSearchMode Then
        LogModeBuilder "Created location search mode: " & modeName
    End If
End Function

'============================================================
' MODE CUSTOMIZATION HELPERS
'============================================================

' Adds a custom column filter to an existing mode
Public Function AddColumnFilter(modeName As String, columnName As String, filterType As String, filterValue As String) As Boolean
    ' Implementation would modify existing mode configuration
    ' For now, log the intent
    LogModeBuilder "Would add filter to mode " & modeName & ": " & columnName & " " & filterType & " " & filterValue
    AddColumnFilter = True
End Function

' Sets custom output columns for a mode
Public Function SetModeOutputColumns(modeName As String, columnList As String) As Boolean
    ' Implementation would update mode configuration
    LogModeBuilder "Would set output columns for mode " & modeName & ": " & columnList
    SetModeOutputColumns = True
End Function

' Sets sorting for a mode
Public Function SetModeSorting(modeName As String, sortColumn As String, Optional ascending As Boolean = True) As Boolean
    LogModeBuilder "Would set sorting for mode " & modeName & ": " & sortColumn & " (" & IIf(ascending, "ASC", "DESC") & ")"
    SetModeSorting = True
End Function

'============================================================
' PREDEFINED MODE TEMPLATES
'============================================================

' Sets up common equipment search modes
Public Sub SetupEquipmentModes()
    LogModeBuilder "Setting up equipment search modes..."
    
    ' Equipment by description
    Call CreateTextSearchMode("Equipment Search", "EquipmentData", "Description,Tag,Type")
    
    ' Equipment by ID/Tag
    Call CreateExactMatchMode("Tag Lookup", "EquipmentData", "Tag")
    
    ' Equipment by type
    Call CreateFilteredViewMode("By Type", "EquipmentData", "[@Type]={EQUIPMENT_TYPE}", "Tag,Description,Type,Location")
    
    ' Equipment by location  
    Call CreateLocationSearchMode("By Location", "EquipmentData", "Location")
    
    LogModeBuilder "Equipment modes setup complete"
End Sub

' Sets up sootblower-specific modes
Public Sub SetupSootblowerModes()
    LogModeBuilder "Setting up sootblower search modes..."
    
    ' Sootblower by number
    Call CreateExactMatchMode("Sootblower by Number", "SootblowerData", "Number")
    
    ' Sootblower by location
    Call CreateLocationSearchMode("Sootblower by Location", "SootblowerData", "Floor")
    
    ' Sootblower by cabinet
    Call CreateFilteredViewMode("By Cabinet", "SootblowerData", "[@SB Cabinet]={CABINET}", "Type,Number,Floor,Side,SB Cabinet")
    
    LogModeBuilder "Sootblower modes setup complete"
End Sub

' Sets up valve-specific modes
Public Sub SetupValveModes()
    LogModeBuilder "Setting up valve search modes..."
    
    ' Valve by number
    Call CreateExactMatchMode("Valve by Number", "ValveData", "Valve Number")
    
    ' Valve by system
    Call CreateFilteredViewMode("By System", "ValveData", "[@System]={SYSTEM}", "Valve Number,Description,System,Location")
    
    LogModeBuilder "Valve modes setup complete"
End Sub

'============================================================
' UTILITY FUNCTIONS
'============================================================

Private Function GetFirstNColumns(tableName As String, n As Long) As String
    On Error GoTo ErrorHandler
    
    Dim dataTable As ListObject
    Set dataTable = GetListObject(tableName)
    If dataTable Is Nothing Then
        GetFirstNColumns = "Column1,Column2,Column3"
        Exit Function
    End If
    
    Dim cols As String
    Dim i As Long
    Dim maxCols As Long
    maxCols = Application.WorksheetFunction.Min(n, dataTable.ListColumns.Count)
    
    For i = 1 To maxCols
        If i > 1 Then cols = cols & ","
        cols = cols & CStr(dataTable.HeaderRowRange.Cells(1, i).Value)
    Next i
    
    GetFirstNColumns = cols
    Exit Function
    
ErrorHandler:
    GetFirstNColumns = "Column1,Column2,Column3"
End Function

Private Function GetLocationOutputColumns(tableName As String, locationColumn As String) As String
    ' Get columns relevant for location-based searches
    Dim standardCols As String
    standardCols = GetFirstNColumns(tableName, 3)
    
    ' Add location column if not already included
    If InStr(1, standardCols, locationColumn, vbTextCompare) = 0 Then
        standardCols = standardCols & "," & locationColumn
    End If
    
    GetLocationOutputColumns = standardCols
End Function

Private Sub LogModeBuilder(msg As String)
    Debug.Print "[ModeBuilder] " & Format(Now, "hh:mm:ss") & " " & msg
End Sub

'============================================================
' VALIDATION AND TESTING
'============================================================

' Validates that all registered modes work correctly
Public Sub ValidateAllModes()
    LogModeBuilder "Validating all search modes..."
    
    Dim modes As Collection
    Set modes = GetAvailableModes()
    
    Dim i As Long
    For i = 1 To modes.Count
        Dim modeName As String
        modeName = CStr(modes(i))
        
        If ValidateMode(modeName) Then
            LogModeBuilder "✓ Mode '" & modeName & "' is valid"
        Else
            LogModeBuilder "✗ Mode '" & modeName & "' has issues: " & GetLastError()
        End If
    Next i
    
    LogModeBuilder "Mode validation complete"
End Sub

Private Function ValidateMode(modeName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim mode As SearchMode
    If Not LoadSearchMode(modeName, mode) Then
        Exit Function
    End If
    
    ' Check if source table exists
    If GetListObject(mode.SourceTable) Is Nothing Then
        SetLastError "Source table '" & mode.SourceTable & "' not found"
        Exit Function
    End If
    
    ' Check if output columns exist in source table
    If Not ValidateOutputColumns(mode.SourceTable, mode.OutputColumns) Then
        Exit Function
    End If
    
    ValidateMode = True
    Exit Function
    
ErrorHandler:
    SetLastError "ValidateMode error: " & Err.Number & " - " & Err.Description
    ValidateMode = False
End Function

Private Function ValidateOutputColumns(tableName As String, columnList As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim dataTable As ListObject
    Set dataTable = GetListObject(tableName)
    If dataTable Is Nothing Then
        SetLastError "Table '" & tableName & "' not found"
        Exit Function
    End If
    
    Dim cols As Variant
    cols = Split(columnList, ",")
    
    Dim i As Long
    For i = 0 To UBound(cols)
        Dim colName As String
        colName = Trim(cols(i))
        
        If GetColumnIndex(dataTable, colName) = 0 Then
            SetLastError "Column '" & colName & "' not found in table '" & tableName & "'"
            Exit Function
        End If
    Next i
    
    ValidateOutputColumns = True
    Exit Function
    
ErrorHandler:
    SetLastError "ValidateOutputColumns error: " & Err.Number & " - " & Err.Description
    ValidateOutputColumns = False
End Function

'============================================================
' DEMO AND TESTING FUNCTIONS
'============================================================

' Demonstrates creating and using different search modes
Public Sub DemoSearchModes()
    LogModeBuilder "=== SEARCH MODE DEMO ==="
    
    ' Setup demo modes
    Call SetupEquipmentModes
    
    ' Test text search
    LogModeBuilder "Testing text search..."
    Dim resultCount As Long
    resultCount = QuickSearch("pump")
    LogModeBuilder "Text search returned " & resultCount & " results"
    
    ' Test exact match
    LogModeBuilder "Testing exact match..."
    ' resultCount = ExecuteSearch("Tag Lookup", criteria with specific tag)
    
    LogModeBuilder "=== DEMO COMPLETE ==="
End Sub

' Creates a comprehensive set of modes for testing
Public Sub SetupTestModes()
    LogModeBuilder "Setting up test modes..."
    
    ' Create various mode types for testing
    Call CreateTextSearchMode("Test_TextSearch", "EquipmentData", "Description")
    Call CreateExactMatchMode("Test_ExactMatch", "EquipmentData", "Tag")
    Call CreateFilteredViewMode("Test_FilteredView", "EquipmentData", "[@Type]='Pump'", "Tag,Description,Type")
    
    LogModeBuilder "Test modes created"
End Sub