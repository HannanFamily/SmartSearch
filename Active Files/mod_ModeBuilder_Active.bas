Attribute VB_Name = "mod_ModeBuilder_Active"
'============================================================
' MODE BUILDER - Active Version for Production Use
'============================================================
' Purpose: Simplified mode builder that works with the enhanced search engine
' This version focuses on ease of use and reliability
'============================================================

Option Explicit

'============================================================
' QUICK MODE SETUP - Easy Functions
'============================================================

' Create a simple text search mode
Public Function CreateTextSearchMode(modeName As String, tableName As String, Optional searchColumns As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    If Len(Trim(searchColumns)) = 0 Then
        searchColumns = GetFirstColumns(tableName, 3)
    End If
    
    ' For now, just log the mode creation since we're using the original table-driven system
    Debug.Print "[ModeBuilder] Created text search mode: " & modeName & " for table: " & tableName & " columns: " & searchColumns
    CreateTextSearchMode = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[ModeBuilder] Error creating text search mode: " & Err.Description
    CreateTextSearchMode = False
End Function

' Create equipment search modes using the existing ConfigTable system
Public Sub SetupEquipmentSearchModes()
    Debug.Print "[ModeBuilder] Setting up equipment search modes..."
    
    ' Since the enhanced engine uses the existing ConfigTable structure,
    ' we just need to ensure the config entries exist
    Call EnsureConfigEntry("Out_Column1", "Tag", "Output Column")
    Call EnsureConfigEntry("Out_Column2", "Description", "Output Column")
    Call EnsureConfigEntry("Out_Column3", "Type", "Output Column")
    Call EnsureConfigEntry("Out_Column4", "Location", "Output Column")
    Call EnsureConfigEntry("Out_Column5", "System", "Output Column")
    
    ' Ensure search inputs are configured
    Call EnsureConfigEntry("InputCell_DescripSearch", "SearchBox", "Input Named Range")
    Call EnsureConfigEntry("InputCell_ValveNumSearch", "ValveNumSearchBox", "Input Named Range")
    
    Debug.Print "[ModeBuilder] Equipment search modes configured"
End Sub

' Create sootblower search modes
Public Sub SetupSootblowerSearchModes()
    Debug.Print "[ModeBuilder] Setting up sootblower search modes..."
    
    ' Configure output columns for sootblower data
    Call EnsureConfigEntry("Out_Column1", "Type", "Output Column")
    Call EnsureConfigEntry("Out_Column2", "Number", "Output Column")
    Call EnsureConfigEntry("Out_Column3", "Floor", "Output Column")
    Call EnsureConfigEntry("Out_Column4", "Side", "Output Column")
    Call EnsureConfigEntry("Out_Column5", "SB Cabinet", "Output Column")
    
    Debug.Print "[ModeBuilder] Sootblower search modes configured"
End Sub

'============================================================
' CONFIG TABLE MANAGEMENT
'============================================================

Private Sub EnsureConfigEntry(configKey As String, configValue As String, configType As String)
    On Error Resume Next
    
    ' Check if ConfigTable exists and has this entry
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    If ws Is Nothing Then
        Debug.Print "[ModeBuilder] Warning: ConfigSheet not found"
        Exit Sub
    End If
    
    Dim configTable As ListObject
    Set configTable = ws.ListObjects("ConfigTable")
    If configTable Is Nothing Then
        Debug.Print "[ModeBuilder] Warning: ConfigTable not found"
        Exit Sub
    End If
    
    ' Check if entry already exists
    Dim r As Range
    For Each r In configTable.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), configKey, vbTextCompare) = 0 Then
            ' Entry exists, update if needed
            If StrComp(CStr(r.Cells(1, 2).Value), configValue, vbTextCompare) <> 0 Then
                r.Cells(1, 2).Value = configValue
                Debug.Print "[ModeBuilder] Updated config: " & configKey & " = " & configValue
            End If
            Exit Sub
        End If
    Next r
    
    ' Entry doesn't exist, add it
    Dim newRow As ListRow
    Set newRow = configTable.ListRows.Add
    newRow.Range.Cells(1, 1).Value = configKey
    newRow.Range.Cells(1, 2).Value = configValue
    If configTable.ListColumns.Count >= 3 Then
        newRow.Range.Cells(1, 3).Value = configType
    End If
    
    Debug.Print "[ModeBuilder] Added config: " & configKey & " = " & configValue
End Sub

'============================================================
' UTILITY FUNCTIONS
'============================================================

Private Function GetFirstColumns(tableName As String, count As Long) As String
    On Error Resume Next
    
    Dim dataTable As ListObject
    Set dataTable = FindTable(tableName)
    If dataTable Is Nothing Then
        GetFirstColumns = "Column1,Column2,Column3"
        Exit Function
    End If
    
    Dim cols As String
    Dim i As Long
    Dim maxCols As Long
    maxCols = Application.WorksheetFunction.Min(count, dataTable.ListColumns.Count)
    
    For i = 1 To maxCols
        If i > 1 Then cols = cols & ","
        cols = cols & CStr(dataTable.HeaderRowRange.Cells(1, i).Value)
    Next i
    
    GetFirstColumns = cols
End Function

Private Function FindTable(tableName As String) As ListObject
    On Error Resume Next
    
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

'============================================================
' VALIDATION AND TESTING
'============================================================

' Test the mode builder functions
Public Sub TestModeBuilder()
    Debug.Print "=== MODE BUILDER TEST ==="
    
    ' Test equipment modes
    Call SetupEquipmentSearchModes
    
    ' Test sootblower modes
    Call SetupSootblowerSearchModes
    
    ' Test configuration
    Call ValidateConfiguration
    
    Debug.Print "=== MODE BUILDER TEST COMPLETE ==="
End Sub

Public Sub ValidateConfiguration()
    Debug.Print "[ModeBuilder] Validating configuration..."
    
    ' Check essential config entries
    Dim essentialConfigs As Variant
    essentialConfigs = Array("DATA_TABLE_NAME", "DASHBOARD_SHEET", "ResultsStartCell", "StatusCell")
    
    Dim i As Long
    For i = 0 To UBound(essentialConfigs)
        Dim value As String
        value = GetConfigValueSafe(CStr(essentialConfigs(i)))
        If Len(Trim(value)) > 0 Then
            Debug.Print "[ModeBuilder] ✓ " & essentialConfigs(i) & " = " & value
        Else
            Debug.Print "[ModeBuilder] ✗ Missing: " & essentialConfigs(i)
        End If
    Next i
    
    Debug.Print "[ModeBuilder] Configuration validation complete"
End Sub

Private Function GetConfigValueSafe(key As String) As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    If ws Is Nothing Then Exit Function
    
    Dim configTable As ListObject
    Set configTable = ws.ListObjects("ConfigTable")
    If configTable Is Nothing Then Exit Function
    
    Dim r As Range
    For Each r In configTable.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            GetConfigValueSafe = CStr(r.Cells(1, 2).Value)
            Exit Function
        End If
    Next r
End Function

'============================================================
' SETUP HELPERS FOR COMMON SCENARIOS
'============================================================

' Complete setup for a new dashboard
Public Sub SetupNewDashboard()
    Debug.Print "[ModeBuilder] Setting up new dashboard..."
    
    ' Essential config entries
    Call EnsureConfigEntry("DASHBOARD_SHEET", "Dashboard", "Sheet Name")
    Call EnsureConfigEntry("DATA_TABLE_NAME", "EquipmentData", "Table Name")
    Call EnsureConfigEntry("ResultsStartCell", "ResultsStart", "Named Range")
    Call EnsureConfigEntry("StatusCell", "SearchStatus", "Named Range")
    Call EnsureConfigEntry("SPLICER_THRESHOLD", "250", "Number")
    Call EnsureConfigEntry("TEMP_FILTER_COL_NAME", "Temp_SearchInclude", "Column Name")
    
    ' Setup search modes
    Call SetupEquipmentSearchModes
    
    Debug.Print "[ModeBuilder] New dashboard setup complete"
End Sub

' Quick configuration for sootblower dashboard
Public Sub ConfigureForSootblowers()
    Debug.Print "[ModeBuilder] Configuring for sootblower data..."
    
    Call EnsureConfigEntry("DATA_TABLE_NAME", "SootblowerData", "Table Name")
    Call EnsureConfigEntry("DataTable_EquipDescription", "Type", "Column Name")
    Call SetupSootblowerSearchModes
    
    Debug.Print "[ModeBuilder] Sootblower configuration complete"
End Sub