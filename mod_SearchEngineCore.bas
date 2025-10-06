'Attribute VB_Name = "mod_SearchEngineCore"  ' Commented out for manual copy/paste convenience
Option Explicit

' ============================================================================
' Module:        mod_SearchEngineCore (Scaffolding Version)
' Purpose:       Provide a clean, testable, extensible core API surface for the
'                dashboard search system. This is a REPLACEMENT / IMPROVEMENT
'                scaffold – NOT final until we ingest & review existing logic.
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
    rowCount = WriteOutput cfg, srcData, filteredIndices

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
            cfg.AutoRefresh = CBool(Val(Nz(r.Range.Cells(1, 7).Value)))
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
    FindModeConfigTable = FindTableByName("ModeConfigTable")
End Function

Private Function FindTableByName(ByVal tableName As String) As ListObject
    On Error Resume Next
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.Name) = LCase$(tableName) Then
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
                If CBool(Val(r.Range.Cells(1, 7).Value)) Then col.Add mName
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
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\"")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

' ============================================================================
' End of module (scaffolding)
' ============================================================================
