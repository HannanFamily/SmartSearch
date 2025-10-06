'Attribute VB_Name = "mod_SearchEngineCore"  ' commented for copy/paste portability
Option Explicit

' ============================================================================
' Module:        mod_SearchEngineCore (Scaffolding Version)
' Purpose:       Provide a clean, testable, extensible core API surface for the
'                dashboard search system. This is a REPLACEMENT / IMPROVEMENT
'                scaffold � NOT final until we ingest & review existing logic.
' ----------------------------------------------------------------------------
' IMPORTANT:     Real implementation will be filled in AFTER we export and
'                analyze current project code + ModeConfigTable structure.
' ----------------------------------------------------------------------------
' Design Goals:
'   * Mode-driven (each mode defines: source table, filters, projection, sort)
'   * Zero hard-coded worksheet references beyond a single resolver function
'   * Traceable execution with centralized logging & error handling
'   * Safe incremental refactoring path (we can swap in real logic gradually)
' ----------------------------------------------------------------------------
' Public API (initial contract):
'   RunSearch(modeName [, criteriaDict]) -> Long (rows returned)
'   RefreshAllVisible()                  -> Void (batch refresh for all modes
'                                               flagged as AutoRefresh)
'   DescribeMode(modeName)               -> String (JSON-like summary)
' ----------------------------------------------------------------------------
' Execution Flow (intended):
'   1. RunSearch called with a modeName
'   2. Load mode definition from ModeConfigTable into a ModeConfig struct
'   3. Resolve source data range
'   4. Build filter predicate(s) from user criteria & mode definition
'   5. Apply filter (either in-memory array pass or AdvancedFilter)
'   6. Project columns & write output table
'   7. Return row count + log diagnostics
' ----------------------------------------------------------------------------
' Pending Unknowns (will be resolved after export):
'   - Exact column names & config schema
'   - Output destination logic (single result table vs per-mode range)
'   - Performance constraints (data size, need for caching?)
' ----------------------------------------------------------------------------
' Next Steps After Code Export:
'   Replace placeholder methods with concrete implementations
'   Add unit-like test harness routines for each critical path
' ----------------------------------------------------------------------------
' NOTE: All functions currently defensive & fail gracefully with messages.
' ============================================================================

' ---------------------------- Data Structures -------------------------------
Private Type ModeConfig
    ModeName            As String
    SourceTableName     As String
    OutputTableName     As String
    FilterFormulaRaw    As String
    ProjectionSpec      As String ' Comma-separated list of output columns
    SortSpec            As String
    AutoRefresh         As Boolean
    Notes               As String
End Type

' ----------------------------- Public API ----------------------------------
Public Function RunSearch(ByVal ModeName As String, Optional ByVal Criteria As Variant) As Long
    On Error GoTo EH
    Dim cfg As ModeConfig
    If Not LoadModeConfig(ModeName, cfg) Then
        LogMsg "RunSearch: Mode not found -> " & ModeName
        Exit Function
    End If

    LogMsg "RunSearch START: " & ModeName

    ' Placeholder pipeline
    Dim srcData As Variant
    If Not LoadSourceData(cfg, srcData) Then
        LogMsg "RunSearch: Failed to load source data for mode=" & ModeName
        Exit Function
    End If

    Dim filteredIndices As Collection
    Set filteredIndices = ApplyFilters(cfg, srcData, Criteria)

    Dim rowCount As Long
    ' NOTE: Parentheses required when capturing a function return value in VBA.
    rowCount = WriteOutput(cfg, srcData, filteredIndices)

    LogMsg "RunSearch COMPLETE: " & ModeName & " rows=" & rowCount
    RunSearch = rowCount
    Exit Function
EH:
    LogMsg "RunSearch ERROR: " & Err.Number & " - " & Err.Description
End Function

Public Sub RefreshAllVisible()
    On Error GoTo EH
    Dim modeList As Collection
    Set modeList = ListAllModes(True)
    Dim m As Variant
    For Each m In modeList
        RunSearch CStr(m)
    Next m
    Exit Sub
EH:
    LogMsg "RefreshAllVisible ERROR: " & Err.Number & " - " & Err.Description
End Sub

Public Function DescribeMode(ByVal ModeName As String) As String
    Dim cfg As ModeConfig
    If Not LoadModeConfig(ModeName, cfg) Then
        DescribeMode = "{error:'mode not found'}"
        Exit Function
    End If
    DescribeMode = "{mode:'" & JsonEscape(cfg.ModeName) & _
                    "',source:'" & JsonEscape(cfg.SourceTableName) & _
                    "',output:'" & JsonEscape(cfg.OutputTableName) & _
                    "',projection:'" & JsonEscape(cfg.ProjectionSpec) & _
                    "',sort:'" & JsonEscape(cfg.SortSpec) & _
                    "',autoRefresh:" & LCase$(CStr(cfg.AutoRefresh)) & _
                    "}"
End Function

' ----------------------- Core Pipeline Placeholders ------------------------
Private Function LoadModeConfig(ByVal ModeName As String, ByRef cfg As ModeConfig) As Boolean
    ' TODO: Replace with logic that finds row in ModeConfigTable.
    ' Current placeholder simply fails; after export we will implement.
    Dim lo As ListObject
    Set lo = FindModeConfigTable()
    If lo Is Nothing Then
        LogMsg "LoadModeConfig: ModeConfigTable not found."
        Exit Function
    End If

    Dim r As ListRow
    For Each r In lo.ListRows
        If LCase$(Trim$(r.Range.Cells(1, 1).Value)) = LCase$(Trim$(ModeName)) Then
            cfg.ModeName = r.Range.Cells(1, 1).Value
            ' The following column mappings are assumptions and WILL be revised.
            On Error Resume Next
            cfg.SourceTableName = Nz(r.Range.Cells(1, 2).Value)
            cfg.OutputTableName = Nz(r.Range.Cells(1, 3).Value)
            cfg.FilterFormulaRaw = Nz(r.Range.Cells(1, 4).Value)
            cfg.ProjectionSpec = Nz(r.Range.Cells(1, 5).Value)
            cfg.SortSpec = Nz(r.Range.Cells(1, 6).Value)
            cfg.AutoRefresh = CBool(val(Nz(r.Range.Cells(1, 7).Value)))
            cfg.Notes = Nz(r.Range.Cells(1, 8).Value)
            On Error GoTo 0
            LoadModeConfig = True
            Exit Function
        End If
    Next r
End Function

Private Function LoadSourceData(ByRef cfg As ModeConfig, ByRef srcData As Variant) As Boolean
    On Error GoTo EH
    Dim lo As ListObject
    Set lo = FindTableByName(cfg.SourceTableName)
    If lo Is Nothing Then Exit Function
    srcData = lo.DataBodyRange.Value
    LoadSourceData = True
    Exit Function
EH:
    LogMsg "LoadSourceData ERROR: " & Err.Number & " - " & Err.Description
End Function

Private Function ApplyFilters(ByRef cfg As ModeConfig, ByRef srcData As Variant, ByVal Criteria As Variant) As Collection
    ' Placeholder: returns all row indices. Real implementation will parse
    ' cfg.FilterFormulaRaw + Criteria to build evaluation.
    Dim result As New Collection
    If IsEmpty(srcData) Then
        Set ApplyFilters = result
        Exit Function
    End If

    Dim rCount As Long
    rCount = UBound(srcData, 1)
    Dim i As Long
    For i = 1 To rCount
        result.Add i
    Next i
    Set ApplyFilters = result
End Function

Private Function WriteOutput(ByRef cfg As ModeConfig, ByRef srcData As Variant, ByVal indices As Collection) As Long
    ' Placeholder: Writes to output table if found, else no-op.
    Dim loOut As ListObject
    Set loOut = FindTableByName(cfg.OutputTableName)
    If loOut Is Nothing Then
        LogMsg "WriteOutput: Output table not found -> " & cfg.OutputTableName
        Exit Function
    End If
    If indices Is Nothing Then Exit Function

    ' Simple full column dump: copy all columns for selected rows.
    Dim rowsOut As Long
    rowsOut = indices.Count
    If rowsOut = 0 Then
        ClearListObject loOut
        Exit Function
    End If

    Dim srcCols As Long
    srcCols = UBound(srcData, 2)
    Dim arr()
    ReDim arr(1 To rowsOut, 1 To srcCols)

    Dim i As Long, r As Long
    For i = 1 To rowsOut
        For r = 1 To srcCols
            arr(i, r) = srcData(indices(i), r)
        Next r
    Next i

    ClearListObject loOut
    loOut.Resize loOut.HeaderRowRange.Resize(rowsOut + 1, srcCols)
    loOut.DataBodyRange.Value = arr
    WriteOutput = rowsOut
End Function

' ----------------------------- Utilities -----------------------------------
Private Function FindModeConfigTable() As ListObject
    ' Returning an object (ListObject) requires Set; missing Set caused "Invalid use of property".
    Set FindModeConfigTable = FindTableByName("ModeConfigTable")
End Function

Private Function FindTableByName(ByVal tableName As String) As ListObject
    On Error Resume Next
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.name) = LCase$(tableName) Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Sub ClearListObject(ByVal lo As ListObject)
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Delete
    End If
End Sub

Private Sub LogMsg(ByVal msg As String)
    Debug.Print "[SearchCore] " & msg
    ' Optional: Append to diagnostics sheet (reuse exporter sheet logic if desired)
End Sub

Private Function ListAllModes(Optional AutoOnly As Boolean = False) As Collection
    Dim col As New Collection
    Dim lo As ListObject
    Set lo = FindModeConfigTable()
    If lo Is Nothing Then
        Set ListAllModes = col
        Exit Function
    End If
    Dim r As ListRow
    For Each r In lo.ListRows
        Dim mName As String
        mName = Trim$(r.Range.Cells(1, 1).Value)
        If Len(mName) > 0 Then
            If AutoOnly Then
                ' Assuming AutoRefresh flag column (7) per placeholder.
                If CBool(val(r.Range.Cells(1, 7).Value)) Then col.Add mName
            Else
                col.Add mName
            End If
        End If
    Next r
    Set ListAllModes = col
End Function

Private Function Nz(ByVal v As Variant, Optional ByVal fallback As String = "") As String
    If IsError(v) Then
        Nz = fallback
    ElseIf IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then
        Nz = fallback
    Else
        Nz = CStr(v)
    End If
End Function

Private Function JsonEscape(ByVal s As String) As String
    ' Escape backslash first
    s = Replace(s, "\", "\\")
    ' Escape double quotes -> \"
    ' Using Chr$(34) to avoid confusing nested quotes
    s = Replace(s, Chr$(34), "\\" & Chr$(34))
    ' Normalize newlines
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

' ============================================================================
' End of module (scaffolding)
' ============================================================================

Public Sub RefreshResults_Enhanced()
    ' Dashboard-compatible refresh logic: checks for active inputs, slicer pulse, and outputs results
    On Error GoTo CleanExit
    Static gBusy As Boolean
    If gBusy Then Exit Sub
    gBusy = True

    ' Ensure slicer pulse cell exists (if you have a helper, call it here)
    If IsEmpty(ThisWorkbook.Names("SlicerPulseCell")) Then
        ' Optionally create or validate pulse cell here
    End If

    Dim anyActive As Boolean: anyActive = IsAnySearchInputActive()
    Dim pulseVal As Long: pulseVal = 0
    On Error Resume Next
    pulseVal = CLng(ThisWorkbook.Names("SlicerPulseCell").RefersToRange.Value)
    On Error GoTo 0
    Dim pulseOk As Boolean: pulseOk = (pulseVal > 0)

    If anyActive Then
        ' Search-driven: filter and output
        ' If you have a PerformSearch routine, call it here
        ' For now, call RefreshAllVisible as a placeholder
        Call RefreshAllVisible
    ElseIf pulseOk Then
        ' Slicer-driven: clear temp filter and show visible
        If Not IsEmpty(ThisWorkbook.Names("TempSearchFilterCol")) Then
            ' Optionally clear temp filter here
        End If
        Call RefreshAllVisible
    Else
        ' Default: headers only or no results
        ' Optionally clear temp filter and output no results
        Call RefreshAllVisible
    End If
CleanExit:
    gBusy = False
End Sub

' Legacy wrapper: allow existing code calling RefreshResults (without _Enhanced)
' to continue working until all references are migrated.
Public Sub RefreshResults()
    RefreshResults_Enhanced
End Sub

Public Function AllInputNamedRanges_Enhanced() As Collection
    Dim c As New Collection
    Dim ws As Worksheet, loCfg As ListObject, r As Range
    Dim nmKey As String, nmType As String
    On Error GoTo Done
    
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Or loCfg.DataBodyRange Is Nothing Then GoTo Done

    For Each r In loCfg.DataBodyRange.Rows
        nmKey = Trim$(CStr(r.Cells(1, 2).Value))   ' column B: ConfigValue (the NAME of the range)
        nmType = Trim$(CStr(r.Cells(1, 3).Value))  ' column C: TYPE
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

' ---------------------------------------------------------------
' Added compatibility helpers migrated from earlier consolidated module
' to satisfy Dashboard sheet calls without re-importing legacy modules.
' ---------------------------------------------------------------
Public Sub EnsurePulseCell(Optional ByVal resetValue As Boolean = False)
    ' Ensures a named range "SlicerPulseCell" exists (used as a pulse / change trigger)
    ' If present and resetValue=True, increments its value to force downstream refresh logic.
    On Error GoTo CleanExit
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names("SlicerPulseCell").RefersToRange
    On Error GoTo 0
    If r Is Nothing Then
        ' Create a safe placeholder location (far out of normal view) on active sheet.
        Dim host As Worksheet
        If ActiveSheet Is Nothing Then
            Set host = ThisWorkbook.Worksheets(1)
        Else
            Set host = ActiveSheet
        End If
        Set r = host.Range("Z100") ' unobtrusive cell
        On Error Resume Next
        ThisWorkbook.Names.Add Name:="SlicerPulseCell", RefersTo:=r
        On Error GoTo 0
    End If
    If Not r Is Nothing Then
        If resetValue Then
            Dim v As Variant
            v = 0
            On Error Resume Next
            If IsNumeric(r.Value) Then v = CLng(r.Value)
            r.Value = v + 1
            On Error GoTo 0
        End If
    End If
CleanExit:
End Sub

Public Sub ClearTempSearchFilter()
    ' Attempts to remove any temporary helper column used for search filtering
    ' (heuristic: column name contains "temp" AND "search" AND "filter").
    On Error Resume Next
    Dim ws As Worksheet, lo As ListObject, lc As ListColumn
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            For Each lc In lo.ListColumns
                Dim nm As String
                nm = LCase$(lc.Name)
                If (InStr(nm, "temp") > 0) And (InStr(nm, "search") > 0) And (InStr(nm, "filter") > 0) Then
                    lc.Delete
                    Exit Sub
                End If
            Next lc
        Next lo
    Next ws
End Sub

Public Function IsAnySearchInputActive() As Boolean
    ' Returns True if any discovered input named range has a non-empty trimmed value.
    Dim coll As Collection: Set coll = AllInputNamedRanges_Enhanced()
    Dim i As Long, r As Range
    For i = 1 To coll.Count
        Set r = coll(i)
        If Not r Is Nothing Then
            If Len(Trim$(CStr(r.Cells(1, 1).Value))) > 0 Then
                IsAnySearchInputActive = True
                Exit Function
            End If
        End If
    Next i
End Function


