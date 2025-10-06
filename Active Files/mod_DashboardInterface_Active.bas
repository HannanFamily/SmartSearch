Attribute VB_Name = "mod_DashboardInterface_Active"
'============================================================
' DASHBOARD INTERFACE - Active Version for Production Use
'============================================================
' Purpose: Simple, reliable dashboard interface functions
' Connects to the enhanced search engine with minimal complexity
'============================================================

Option Explicit

'============================================================
' MAIN DASHBOARD FUNCTIONS
'============================================================

' Main search button handler - call this from your search button
Public Sub PerformDashboardSearch()
    On Error GoTo ErrorHandler
    
    ' Use the enhanced search engine's RefreshResults function
    Call RefreshResults
    Exit Sub
    
ErrorHandler:
    MsgBox "Search error: " & Err.Description, vbExclamation, "Search Engine"
End Sub

' Clear search button handler - call this from your clear button
Public Sub ClearDashboardSearch()
    On Error Resume Next
    
    ' Use the enhanced search engine's clear functions
        Call ClearFilters_Enhanced
End Sub

' Show all data button handler
Public Sub ShowAllDashboardData()
    On Error Resume Next
    
    ' Clear search inputs but preserve slicer state
    Call ClearSearchBoxes
    Call ClearTempSearchFilter
    Call OutputAllVisible
End Sub

' Refresh current view (for slicer changes)
Public Sub RefreshDashboardView()
    On Error Resume Next
    Call RefreshResults
End Sub

'============================================================
' INITIALIZATION AND SETUP
'============================================================

' Initialize the dashboard system - call this when workbook opens
Public Sub InitializeDashboard()
    On Error GoTo ErrorHandler
    
    ' Setup pulse cell for slicer detection
    Call EnsurePulseCell(True)
    
    ' Ensure basic configuration exists
    Call EnsureBasicConfig
    
    ' Clear any existing search state
        Call ClearFilters_Enhanced
    
    ' Show initial state
    Call RefreshResults
    
    Debug.Print "[Dashboard] Initialization complete"
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Dashboard] Initialization error: " & Err.Description
End Sub

' Setup basic configuration if missing
Private Sub EnsureBasicConfig()
    On Error Resume Next
    
    ' Use the mode builder to ensure basic config
    Call SetupNewDashboard
End Sub

'============================================================
' BUTTON EVENT HANDLERS
'============================================================
' These functions can be directly assigned to button Click events

Public Sub btnSearch_Click()
    Call PerformDashboardSearch
End Sub

Public Sub btnClear_Click()
    Call ClearDashboardSearch
End Sub

Public Sub btnShowAll_Click()
    Call ShowAllDashboardData
End Sub

Public Sub btnRefresh_Click()
    Call RefreshDashboardView
End Sub

'============================================================
' STATUS AND FEEDBACK
'============================================================

' Get current search status
Public Function GetSearchStatus() As String
    On Error Resume Next
    
    Dim statusCell As Range
    Set statusCell = StatusCell()
    If Not statusCell Is Nothing Then
        GetSearchStatus = CStr(statusCell.Value)
    Else
        GetSearchStatus = "Ready"
    End If
End Function

' Update search status message
Public Sub UpdateSearchStatus(message As String)
    On Error Resume Next
    
    Dim statusCell As Range
    Set statusCell = StatusCell()
    If Not statusCell Is Nothing Then
        statusCell.Value = message
    End If
End Sub

'============================================================
' INPUT VALIDATION AND HELPERS
'============================================================

' Check if any search inputs have values
Public Function HasActiveSearch() As Boolean
    HasActiveSearch = IsAnySearchInputActive()
End Function

' Get the current search text
Public Function GetCurrentSearchText() As String
    On Error Resume Next
    GetCurrentSearchText = ReadSearchText_Description()
End Function

' Get the current valve search text
Public Function GetCurrentValveSearch() As String
    On Error Resume Next
    GetCurrentValveSearch = ReadSearchText_Valve()
End Function

'============================================================
' DIAGNOSTICS AND TESTING
'============================================================

' Test all dashboard functions
Public Sub TestDashboard()
    Debug.Print "=== DASHBOARD TEST ==="
    
    ' Test initialization
    Call InitializeDashboard
    Debug.Print "[Dashboard] ✓ Initialization test"
    
    ' Test status functions
    Debug.Print "[Dashboard] Search status: " & GetSearchStatus()
    Debug.Print "[Dashboard] Has active search: " & HasActiveSearch()
    Debug.Print "[Dashboard] Current search text: '" & GetCurrentSearchText() & "'"
    
    ' Test configuration with error handling
    On Error Resume Next
    Debug.Print "[Dashboard] Data table: " & DataTableName()
    Debug.Print "[Dashboard] Dashboard sheet: " & DashboardName()
    On Error GoTo 0
    
    Debug.Print "=== DASHBOARD TEST COMPLETE ==="
End Sub

' Show current dashboard state for debugging
Public Sub ShowDashboardState()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== DASHBOARD STATE ==="
    Debug.Print "[Dashboard] Search Text: '" & GetCurrentSearchText() & "'"
    Debug.Print "[Dashboard] Valve Search: '" & GetCurrentValveSearch() & "'"
    Debug.Print "[Dashboard] Has Active Search: " & HasActiveSearch()
    Debug.Print "[Dashboard] Status: " & GetSearchStatus()
    
    On Error Resume Next
    Debug.Print "[Dashboard] Data Table: " & DataTableName()
    
    Dim resultsRng As Range
    Set resultsRng = resultsStartRng()
    If resultsRng Is Nothing Then
        Debug.Print "[Dashboard] Results Range: Not Set"
    Else
        Debug.Print "[Dashboard] Results Range: " & resultsRng.Address
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "========================"
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Dashboard] Error in ShowDashboardState: " & Err.Description
    Debug.Print "========================"
End Sub

'============================================================
' ADVANCED FEATURES
'============================================================

' Force refresh of slicer pulse (if slicers seem stuck)
Public Sub RefreshSlicerPulse()
    On Error Resume Next
    Call EnsurePulseCell(True)
    Debug.Print "[Dashboard] Slicer pulse refreshed"
End Sub

' Check if slicers are currently filtering data
Public Function AreSslicersActive() As Boolean
    On Error Resume Next
    
    Dim pulseVal As Long
    pulseVal = CurrentSlicerPulse()
    Dim threshold As Long
    threshold = SlicerThreshold()
    
    AreSslicersActive = (pulseVal > 0 And pulseVal <= threshold)
End Function

' Get count of visible rows (after slicer filtering)
Public Function GetVisibleRowCount() As Long
    On Error Resume Next
    
    Dim dataTable As ListObject
    Set dataTable = lo(DataTableName())
    If dataTable Is Nothing Then Exit Function
    
    Dim visibleIndexes As Variant
    visibleIndexes = VisibleRowIndexes(dataTable)
    If Not IsEmpty(visibleIndexes) Then
        GetVisibleRowCount = UBound(visibleIndexes)
    End If
End Function

' Private function to get current slicer pulse value
Private Function CurrentSlicerPulse() As Long
    Dim r As Range: Set r = nr("SlicerPulse")
    If r Is Nothing Then
        CurrentSlicerPulse = 0
    Else
        CurrentSlicerPulse = CLngSafe(CStr(r.Value))
    End If
End Function

'============================================================
' ERROR HANDLING HELPERS
'============================================================

' Safe execution wrapper for search operations
Public Sub SafeExecuteSearch()
    If gBusy Then
        Debug.Print "[Dashboard] Search already in progress, skipping"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    Call PerformDashboardSearch
    Exit Sub
    
ErrorHandler:
    Call UpdateSearchStatus("Search error: " & Err.Description)
    Debug.Print "[Dashboard] Search error: " & Err.Description
End Sub

' Safe execution wrapper for clear operations
Public Sub SafeClearSearch()
    On Error GoTo ErrorHandler
    Call ClearDashboardSearch
    Call UpdateSearchStatus("Search cleared")
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Dashboard] Clear error: " & Err.Description
End Sub