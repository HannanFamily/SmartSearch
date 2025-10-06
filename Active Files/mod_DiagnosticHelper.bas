Attribute VB_Name = "mod_DiagnosticHelper"
'============================================================
' DASHBOARD DIAGNOSTIC HELPER
'============================================================
' Purpose: Diagnose why search function isn't working
' Run DiagnoseSearchIssues() to get detailed analysis
'============================================================

Option Explicit

'============================================================
' MAIN DIAGNOSTIC FUNCTION
'============================================================

Public Sub DiagnoseSearchIssues()
    Debug.Print "========================================="
    Debug.Print "SEARCH FUNCTION DIAGNOSTIC"
    Debug.Print "========================================="
    
    Call CheckNamedRanges
    Call CheckConfiguration
    Call CheckDataTables
    Call CheckDashboardEvents
    Call TestSearchFunctions
    
    Debug.Print "========================================="
    Debug.Print "DIAGNOSTIC COMPLETE"
    Debug.Print "========================================="
End Sub

'============================================================
' DETAILED CHECKS
'============================================================

Private Sub CheckNamedRanges()
    Debug.Print "--- NAMED RANGES CHECK ---"
    
    Dim requiredRanges As Variant
    requiredRanges = Array("SearchBox", "ValveNumSearchBox", "ResultsStart", "SearchStatus")
    
    Dim i As Long
    For i = 0 To UBound(requiredRanges)
        Dim rangeName As String
        rangeName = CStr(requiredRanges(i))
        
        Dim testRange As Range
        On Error Resume Next
        Set testRange = ThisWorkbook.Names(rangeName).RefersToRange
        On Error GoTo 0
        
        If testRange Is Nothing Then
            Debug.Print "❌ MISSING: " & rangeName
        Else
            Debug.Print "✓ Found: " & rangeName & " → " & testRange.Address(True, True, xlA1, True)
        End If
    Next i
End Sub

Private Sub CheckConfiguration()
    Debug.Print "--- CONFIGURATION CHECK ---"
    
    ' Check if ConfigSheet exists
    Dim configSheet As Worksheet
    On Error Resume Next
    Set configSheet = ThisWorkbook.Worksheets("ConfigSheet")
    On Error GoTo 0
    
    If configSheet Is Nothing Then
        Debug.Print "❌ ConfigSheet not found"
        Exit Sub
    Else
        Debug.Print "✓ ConfigSheet found"
    End If
    
    ' Check if ConfigTable exists
    Dim configTable As ListObject
    On Error Resume Next
    Set configTable = configSheet.ListObjects("ConfigTable")
    On Error GoTo 0
    
    If configTable Is Nothing Then
        Debug.Print "❌ ConfigTable not found"
    Else
        Debug.Print "✓ ConfigTable found (" & configTable.ListRows.Count & " entries)"
    End If
    
    ' Check key configuration values
    Debug.Print "Key Config Values:"
    Call CheckConfigValue("DATA_TABLE_NAME")
    Call CheckConfigValue("DASHBOARD_SHEET")
    Call CheckConfigValue("ResultsStartCell")
    Call CheckConfigValue("StatusCell")
End Sub

Private Sub CheckConfigValue(configKey As String)
    On Error Resume Next
    Dim value As String
    value = GetConfigValue(configKey)
    If Len(value) > 0 Then
        Debug.Print "  ✓ " & configKey & " = " & value
    Else
        Debug.Print "  ❌ " & configKey & " = [MISSING]"
    End If
    On Error GoTo 0
End Sub

Private Sub CheckDataTables()
    Debug.Print "--- DATA TABLES CHECK ---"
    
    On Error Resume Next
    Dim tableName As String
    tableName = DataTableName()
    
    If Len(tableName) = 0 Then
        Debug.Print "❌ No data table name configured"
        Exit Sub
    End If
    
    Debug.Print "Looking for table: " & tableName
    
    Dim dataTable As ListObject
    Set dataTable = lo(tableName)
    
    If dataTable Is Nothing Then
        Debug.Print "❌ Data table '" & tableName & "' not found"
    Else
        Debug.Print "✓ Data table found: " & tableName & " (" & dataTable.ListRows.Count & " rows)"
    End If
    On Error GoTo 0
End Sub

Private Sub CheckDashboardEvents()
    Debug.Print "--- DASHBOARD EVENTS CHECK ---"
    
    ' Check if Dashboard worksheet exists
    Dim dashSheet As Worksheet
    On Error Resume Next
    Set dashSheet = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    
    If dashSheet Is Nothing Then
        Debug.Print "❌ Dashboard worksheet not found"
    Else
        Debug.Print "✓ Dashboard worksheet found"
        
        ' Check if it has the right class name
        If TypeName(dashSheet) = "Dashboard" Then
            Debug.Print "✓ Dashboard class properly connected"
        Else
            Debug.Print "⚠️  Dashboard class may not be connected"
        End If
    End If
End Sub

Private Sub TestSearchFunctions()
    Debug.Print "--- SEARCH FUNCTIONS TEST ---"
    
    On Error Resume Next
    
    ' Test input detection
    Dim hasActiveInput As Boolean
    hasActiveInput = IsAnySearchInputActive()
    Debug.Print "Has active search input: " & hasActiveInput
    
    ' Test input reading
    Dim searchText As String
    Dim valveText As String
    searchText = ReadSearchText_Description()
    valveText = ReadSearchText_Valve()
    Debug.Print "Current search text: '" & searchText & "'"
    Debug.Print "Current valve text: '" & valveText & "'"
    
    ' Test pulse system
    Dim pulseValue As Long
    pulseValue = CurrentSlicerPulse()
    Debug.Print "Current slicer pulse: " & pulseValue
    
    On Error GoTo 0
End Sub

'============================================================
' QUICK FIX FUNCTIONS
'============================================================

Public Sub CreateMissingNamedRanges()
    Debug.Print "Creating missing named ranges..."
    
    On Error Resume Next
    
    ' Create named ranges with default locations
    If ThisWorkbook.Names("SearchBox").RefersToRange Is Nothing Then
        ThisWorkbook.Names.Add "SearchBox", "=Dashboard!$B$5"
        Debug.Print "Created: SearchBox → Dashboard!$B$5"
    End If
    
    If ThisWorkbook.Names("ValveNumSearchBox").RefersToRange Is Nothing Then
        ThisWorkbook.Names.Add "ValveNumSearchBox", "=Dashboard!$B$7"
        Debug.Print "Created: ValveNumSearchBox → Dashboard!$B$7"
    End If
    
    If ThisWorkbook.Names("ResultsStart").RefersToRange Is Nothing Then
        ThisWorkbook.Names.Add "ResultsStart", "=Dashboard!$A$10"
        Debug.Print "Created: ResultsStart → Dashboard!$A$10"
    End If
    
    If ThisWorkbook.Names("SearchStatus").RefersToRange Is Nothing Then
        ThisWorkbook.Names.Add "SearchStatus", "=Dashboard!$A$8"
        Debug.Print "Created: SearchStatus → Dashboard!$A$8"
    End If
    
    On Error GoTo 0
    Debug.Print "Named ranges creation complete"
End Sub

Public Sub TestSearchNow()
    Debug.Print "Testing search function..."
    On Error Resume Next
    Call RefreshResults
    Debug.Print "Search test complete - check Dashboard for results"
    On Error GoTo 0
End Sub

Public Sub ForceSlicerPulseSetup()
    Debug.Print "Setting up slicer pulse detection..."
    Call EnsurePulseCell(True)
    Debug.Print "Slicer pulse setup complete"
End Sub