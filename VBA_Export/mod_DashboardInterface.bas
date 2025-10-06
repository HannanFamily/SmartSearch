'Attribute VB_Name = "mod_DashboardInterface"  ' commented for copy/paste
'============================================================
' DASHBOARD INTERFACE - Search Integration
'============================================================
' Purpose: Connects the search engine to dashboard controls
' Provides easy-to-use public functions for dashboard events
'============================================================

Option Explicit

'============================================================
' PUBLIC DASHBOARD FUNCTIONS
'============================================================

' Main search function - called from dashboard buttons/events
Public Sub PerformDashboardSearch()
    On Error GoTo ErrorHandler
    
    ' Get search criteria from dashboard inputs
    Dim criteria As SearchCriteria
    criteria = GetDashboardSearchCriteria()
    
    ' Get active mode
    Dim activeMode As String
    activeMode = GetActiveDashboardMode()
    
    ' Execute search
    Dim resultCount As Long
    resultCount = ExecuteSearch(activeMode, criteria)
    
    ' Update dashboard status
    UpdateDashboardStatus resultCount, activeMode
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Search error: " & GetLastError(), vbExclamation, "Search Engine"
End Sub

' Quick search from main search box
Public Sub QuickDashboardSearch()
    On Error GoTo ErrorHandler
    
    Dim searchText As String
    searchText = GetDashboardSearchText()
    
    If Len(Trim(searchText)) = 0 Then
        Call ClearDashboardResults
        Exit Sub
    End If
    
    Dim resultCount As Long
    resultCount = QuickSearch(searchText)
    
    UpdateDashboardStatus resultCount, "Quick Search"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Quick search error: " & GetLastError(), vbExclamation, "Search Engine"
End Sub

' Refresh current view (for slicer changes)
Public Sub RefreshDashboardView()
    On Error Resume Next
    Call RefreshVisible
    UpdateDashboardStatus -1, "Refreshed"
End Sub

' Clear all search and filters
Public Sub ClearDashboardSearch()
    On Error Resume Next
    Call ClearSearch
    Call ClearDashboardInputs
    UpdateDashboardStatus 0, "Cleared"
End Sub

' Show all data (remove filters)
Public Sub ShowAllData()
    On Error Resume Next
    Call ClearSearch
    
    ' Show all data in default mode
    Dim criteria As SearchCriteria
    ' Empty criteria = show all
    
    Dim resultCount As Long
    resultCount = ExecuteSearch(GetActiveDashboardMode(), criteria)
    
    UpdateDashboardStatus resultCount, "Show All"
End Sub

'============================================================
' MODE SWITCHING
'============================================================

' Switch to a specific search mode
Public Sub SwitchToMode(modeName As String)
    On Error GoTo ErrorHandler
    
    ' Set the active mode
    Call SetActiveDashboardMode(modeName)
    
    ' Clear current results
    Call ClearDashboardResults
    
    ' Update mode-specific UI elements
    Call UpdateModeSpecificUI(modeName)
    
    ' Refresh with new mode
    Call RefreshDashboardView
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error switching to mode '" & modeName & "': " & GetLastError(), vbExclamation
End Sub

' Get list of available modes for dropdown
Public Function GetModeList() As Variant
    On Error GoTo ErrorHandler
    
    Dim modes As Collection
    Set modes = GetAvailableModes()
    
    If modes.Count = 0 Then
        GetModeList = Array("Default")
        Exit Function
    End If
    
    Dim modeArray() As String
    ReDim modeArray(0 To modes.Count - 1)
    
    Dim i As Long
    For i = 1 To modes.Count
        modeArray(i - 1) = CStr(modes(i))
    Next i
    
    GetModeList = modeArray
    Exit Function
    
ErrorHandler:
    GetModeList = Array("Default")
End Function

'============================================================
' DASHBOARD INPUT/OUTPUT HELPERS
'============================================================

Private Function GetDashboardSearchCriteria() As SearchCriteria
    Dim criteria As SearchCriteria
    
    ' Get search text
    criteria.SearchText = GetDashboardSearchText()
    
    ' Get search options
    criteria.ExactMatch = GetDashboardExactMatch()
    criteria.CaseSensitive = GetDashboardCaseSensitive()
    
    ' Get custom filters
    criteria.CustomFilters = GetDashboardCustomFilters()
    
    GetDashboardSearchCriteria = criteria
End Function

Private Function GetDashboardSearchText() As String
    On Error Resume Next
    
    ' Try multiple possible search input names
    Dim searchRanges As Variant
    searchRanges = Array("SearchBox", "DescriptionSearch", "MainSearch", "SearchInput")
    
    Dim i As Long
    Dim rng As Range
    For i = 0 To UBound(searchRanges)
        Set rng = GetNamedRange(CStr(searchRanges(i)))
        If Not rng Is Nothing Then
            GetDashboardSearchText = Trim(CStr(rng.Value))
            If Len(GetDashboardSearchText) > 0 Then Exit Function
        End If
    Next i
End Function

Private Function GetDashboardExactMatch() As Boolean
    On Error Resume Next
    
    Dim rng As Range
    Set rng = GetNamedRange("ExactMatch")
    If Not rng Is Nothing Then
        GetDashboardExactMatch = CBool(rng.Value)
    End If
End Function

Private Function GetDashboardCaseSensitive() As Boolean
    On Error Resume Next
    
    Dim rng As Range
    Set rng = GetNamedRange("CaseSensitive")
    If Not rng Is Nothing Then
        GetDashboardCaseSensitive = CBool(rng.Value)
    End If
End Function

Private Function GetDashboardCustomFilters() As String
    On Error Resume Next
    
    Dim rng As Range
    Set rng = GetNamedRange("CustomFilters")
    If Not rng Is Nothing Then
        GetDashboardCustomFilters = Trim(CStr(rng.Value))
    End If
End Function

Private Function GetActiveDashboardMode() As String
    On Error Resume Next
    
    Dim rng As Range
    Set rng = GetNamedRange("ModeSelector")
    If Not rng Is Nothing Then
        GetActiveDashboardMode = Trim(CStr(rng.Value))
    End If
    
    ' Default fallback
    If Len(GetActiveDashboardMode) = 0 Then
        GetActiveDashboardMode = "Default"
    End If
End Function

Private Sub SetActiveDashboardMode(modeName As String)
    On Error Resume Next
    
    Dim rng As Range
    Set rng = GetNamedRange("ModeSelector")
    If Not rng Is Nothing Then
        rng.Value = modeName
    End If
End Sub

Private Sub UpdateDashboardStatus(resultCount As Long, modeName As String)
    On Error Resume Next
    
    Dim statusText As String
    If resultCount >= 0 Then
        statusText = "Found " & resultCount & " results (" & modeName & ")"
    Else
        statusText = modeName & " - Ready"
    End If
    
    Dim rng As Range
    Set rng = GetNamedRange("SearchStatus")
    If Not rng Is Nothing Then
        rng.Value = statusText
    End If
    
    ' Also update timestamp
    Set rng = GetNamedRange("LastUpdated")
    If Not rng Is Nothing Then
        rng.Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
    End If
End Sub

Private Sub ClearDashboardInputs()
    On Error Resume Next
    
    ' Clear search inputs
    Dim inputRanges As Variant
    inputRanges = Array("SearchBox", "DescriptionSearch", "MainSearch", "SearchInput", "ValveNumSearchBox", "TagSearch")
    
    Dim i As Long
    Dim rng As Range
    For i = 0 To UBound(inputRanges)
        Set rng = GetNamedRange(CStr(inputRanges(i)))
        If Not rng Is Nothing Then
            rng.ClearContents
        End If
    Next i
End Sub

Private Sub ClearDashboardResults()
    On Error Resume Next
    Call ClearResultsDisplay
End Sub

Private Sub UpdateModeSpecificUI(modeName As String)
    ' Placeholder for mode-specific UI updates
    ' Could show/hide different input controls based on mode
    On Error Resume Next
    
    ' Example: Show location filter for location-based modes
    If InStr(1, modeName, "Location", vbTextCompare) > 0 Then
        ' Show location controls
    ElseIf InStr(1, modeName, "Tag", vbTextCompare) > 0 Then
        ' Show tag-specific controls
    End If
End Sub

'============================================================
' HELPER FUNCTIONS
'============================================================

Private Function GetNamedRange(rangeName As String) As Range
    On Error Resume Next
    Set GetNamedRange = ThisWorkbook.Names(rangeName).RefersToRange
End Function

'============================================================
' SETUP AND INITIALIZATION
'============================================================

' Initialize dashboard search system
Public Sub InitializeDashboard()
    On Error GoTo ErrorHandler
    
    ' Initialize search engine
    Call ExecuteSearch("Default")  ' This will initialize if needed
    
    ' Setup default modes if none exist
    Call SetupDefaultDashboardModes
    
    ' Set initial mode
    Call SwitchToMode("Default")
    
    ' Clear any existing search
    Call ClearDashboardSearch
    
    UpdateDashboardStatus -1, "Ready"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Dashboard initialization error: " & GetLastError(), vbExclamation
End Sub

Private Sub SetupDefaultDashboardModes()
    On Error Resume Next
    
    ' Check if modes already exist
    Dim modes As Collection
    Set modes = GetAvailableModes()
    
    If modes.Count = 0 Then
        ' Create basic modes based on available tables
        Call CreateDefaultModeForTable("EquipmentData")
        Call CreateDefaultModeForTable("SootblowerData")
        Call CreateDefaultModeForTable("ValveData")
    End If
End Sub

Private Sub CreateDefaultModeForTable(tableName As String)
    On Error Resume Next
    
    ' Check if table exists
    If Not GetListObject(tableName) Is Nothing Then
        Call CreateTextSearchMode(tableName & " Search", tableName, GetFirstNColumns(tableName, 5))
    End If
End Sub

'============================================================
' BUTTON EVENT HANDLERS
'============================================================

' These can be assigned to dashboard buttons

Public Sub btnSearch_Click()
    Call PerformDashboardSearch
End Sub

Public Sub btnQuickSearch_Click()
    Call QuickDashboardSearch
End Sub

Public Sub btnClear_Click()
    Call ClearDashboardSearch
End Sub

Public Sub btnShowAll_Click()
    Call ShowAllData
End Sub

Public Sub btnRefresh_Click()
    Call RefreshDashboardView
End Sub

'============================================================
' DIAGNOSTIC AND TESTING
'============================================================

' Test dashboard functions
Public Sub TestDashboard()
    Debug.Print "=== DASHBOARD TEST ==="
    
    ' Initialize
    Call InitializeDashboard
    Debug.Print "Dashboard initialized"
    
    ' Test mode switching
    Call SwitchToMode("Default")
    Debug.Print "Switched to Default mode"
    
    ' Test search
    ' (Would need actual data to test meaningfully)
    
    Debug.Print "=== DASHBOARD TEST COMPLETE ==="
End Sub

' Show current dashboard state
Public Sub ShowDashboardState()
    Debug.Print "=== DASHBOARD STATE ==="
    Debug.Print "Active Mode: " & GetActiveDashboardMode()
    Debug.Print "Search Text: '" & GetDashboardSearchText() & "'"
    Debug.Print "Exact Match: " & GetDashboardExactMatch()
    Debug.Print "Available Modes: " & Join(GetModeList(), ", ")
    Debug.Print "========================"
End Sub