'Attribute VB_Name = "mod_MappingCleanup"
'============================================================
' MAPPING TABLE CLEANUP / STANDARD TERM USAGE ANALYZER
'============================================================
' Purpose:
'   Identify StandardTerm entries in the mapping table that NEVER occur
'   (as a full substring) in any equipment description (or other chosen
'   target text column) of the data table. Optionally flags or deletes
'   them. Designed to scale for very large data sets (640k+ rows).
'
' Core Features:
'   * Scans description column once into memory (array) for fast reuse
'   * Optional word index (token -> candidate row set) to dramatically
'     reduce substring checks for multi-word / long phrase terms
'   * Adds/updates presence column (default: StandardTerm_Present)
'   * Flexible: you can point at alternate data table, column headers,
'     mapping table name, StandardTerm column index, and toggle actions
'   * Two main modes: FLAG (mark 1/0) or DELETE (physically remove rows
'     with zero usage). Delete operates bottom-up for stability.
'   * Progress / diagnostics printed every ProgressInterval terms when
'     Diagnostic or Verbose mode is enabled.
'
' Dependencies (expected existing project helpers):
'   - lo(name As String) As ListObject            (find table by name)
'   - HeaderIndexByText(lo As ListObject, headerText As String) As Long
'   - MappingTableName(), DataTableName(), DataDescConfigKey(),
'     GetConfigValue()  (config access) -- optional; graceful fallback
'
' Safety & Performance:
'   - Turns ScreenUpdating, Events, Calculation off during batch
'   - Early exits + robust error handling with status messages
'   - Token index optional (useWordIndex:=True) but auto-disables if
'     memory could spike (heuristic on row count * avg tokens)
'
' Public Entry Points:
'   CleanMapping                (core worker with parameters)
'   CleanMapping_FlagUnused     (wrapper: FLAG)
'   CleanMapping_DeleteUnused   (wrapper: DELETE)
'
' Usage Examples (Immediate Window):
'   Call CleanMapping_FlagUnused()               ' defaults
'   Call CleanMapping("EquipmentData", "Equipment Description")
'   Call CleanMapping_DeleteUnused(useWordIndex:=False)
'
' Output:
'   - Column (creates if needed) StandardTerm_Present (or custom) with
'     1 if term found in ANY description cell, else 0.
'   - Optional deletion of rows with 0 when action="DELETE".
'
'============================================================
Option Explicit

' Module-level state (m-prefixed to avoid collisions elsewhere)
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevCalc As XlCalculation

' ============================================================
' Public API Wrappers
' ============================================================
Public Sub CleanMapping_FlagUnused()
    CleanMapping action:="FLAG"
End Sub

Public Sub CleanMapping_DeleteUnused()
    CleanMapping action:="DELETE"
End Sub

' ============================================================
' Core Procedure
' ============================================================
Public Sub CleanMapping( _
        Optional ByVal dataTableName As String = "", _
        Optional ByVal descriptionHeader As String = "", _
        Optional ByVal mappingTableName As String = "", _
        Optional ByVal standardTermColIndex As Long = 2, _
        Optional ByVal action As String = "FLAG", _
        Optional ByVal useWordIndex As Boolean = True, _
        Optional ByVal presenceColName As String = "StandardTerm_Present", _
        Optional ByVal verbose As Boolean = True, _
        Optional ByVal ProgressInterval As Long = 500 _
    )
    On Error GoTo EH

    Dim t0 As Double: t0 = Timer
    Dim diag As Boolean: diag = verbose Or GetDiagnosticModeGlobal()

    ' Capture and adjust application state early for performance
    SaveAppState

    ' Resolve tables --------------------------------------------------
    If Len(dataTableName) = 0 Then dataTableName = SafeDataTableName()
    If Len(mappingTableName) = 0 Then mappingTableName = SafeMappingTableName()

    ' Use fully-qualified helper to avoid ambiguous references (legacy module duplicates)
    Dim loData As ListObject: Set loData = mod_SearchEngine_Enhanced.lo(dataTableName)
    If loData Is Nothing Then
        MsgBox "Data table '" & dataTableName & "' not found.", vbExclamation
        Exit Sub
    End If
    If loData.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & dataTableName & "' has no rows.", vbExclamation
        Exit Sub
    End If

    ' Description column header ---------------------------------------
    If Len(descriptionHeader) = 0 Then
        ' Attempt config-based resolution using existing helpers
        descriptionHeader = GetConfigValueSafe("DataTable_EquipDescription", "Equipment Description")
    End If
    Dim descColIdx As Long: descColIdx = mod_SearchEngine_Enhanced.HeaderIndexByText(loData, descriptionHeader)
    If descColIdx = 0 Then
        MsgBox "Description header '" & descriptionHeader & "' not found in data table.", vbExclamation
        Exit Sub
    End If

    Dim loMap As ListObject: Set loMap = mod_SearchEngine_Enhanced.lo(mappingTableName)
    If loMap Is Nothing Then
        MsgBox "Mapping table '" & mappingTableName & "' not found.", vbExclamation
        Exit Sub
    End If
    If loMap.DataBodyRange Is Nothing Then
        MsgBox "Mapping table '" & mappingTableName & "' has no rows.", vbExclamation
        Exit Sub
    End If

    ' StandardTerm column index --------------------------------------
    If standardTermColIndex <= 0 Or standardTermColIndex > loMap.ListColumns.Count Then
        ' Try locate by header text "StandardTerm"
    Dim guessIdx As Long: guessIdx = mod_SearchEngine_Enhanced.HeaderIndexByText(loMap, "StandardTerm")
        If guessIdx > 0 Then
            standardTermColIndex = guessIdx
        Else
            MsgBox "StandardTerm column not found (index invalid and header guess failed).", vbExclamation
            Exit Sub
        End If
    End If

    ' Presence / output column (add if missing) -----------------------
    Dim presenceIdx As Long: presenceIdx = mod_SearchEngine_Enhanced.HeaderIndexByText(loMap, presenceColName)
    If presenceIdx = 0 Then
        Set loMap = EnsureListColumn(loMap, presenceColName, presenceIdx)
        If presenceIdx = 0 Then
            MsgBox "Failed to add presence column '" & presenceColName & "'.", vbCritical
            Exit Sub
        End If
    End If

    ' Acquire description column into array (1-based row major) ------
    Dim descArr As Variant
        descArr = loData.DataBodyRange.Columns(descColIdx).Value  ' 2D array (1..n,1..1) - qualified reference
    Dim descCount As Long: descCount = UBound(descArr, 1)
    If diag Then Debug.Print "[MAPCLEAN] Loaded " & descCount & " descriptions." & _
        " Time=" & Format(Timer - t0, "0.00s")

    ' Preprocess descriptions to lowercase; maybe build word index ----
    Dim lowerDesc() As String: ReDim lowerDesc(1 To descCount)
    Dim i As Long, raw As String
    For i = 1 To descCount
        raw = CStr(descArr(i, 1))
        If Len(raw) > 0 Then
            lowerDesc(i) = LCase$(raw)
        Else
            lowerDesc(i) = ""
        End If
    Next i

    Dim wordIndex As Object
    If useWordIndex Then
        Set wordIndex = BuildWordIndex(lowerDesc, diag)
        ' Heuristic: disable if index seems too large ( > 5 * descCount )
        If Not wordIndex Is Nothing Then
            If wordIndex.Count > descCount * 5 Then
                If diag Then Debug.Print "[MAPCLEAN] Word index too large; disabling (count=" & wordIndex.Count & ")";
                Set wordIndex = Nothing
            End If
        End If
    End If

    ' Iterate mapping terms -------------------------------------------
    Dim termArr As Variant: termArr = loMap.DataBodyRange.Columns(standardTermColIndex).Value
    Dim termCount As Long: termCount = UBound(termArr, 1)

    Dim presenceVals() As Variant: ReDim presenceVals(1 To termCount, 1 To 1)

    Dim found As Boolean, term As String, lcTerm As String, candidates As Variant
    Dim testRows() As Long, r As Long, c As Long

    Dim lastReport As Double: lastReport = Timer

    For r = 1 To termCount
        term = CStr(termArr(r, 1))
        lcTerm = LCase$(Trim$(term))
        found = False

        If Len(lcTerm) = 0 Then
            presenceVals(r, 1) = 0
        Else
            If Not wordIndex Is Nothing Then
                candidates = CandidateRowsFromWordIndex(wordIndex, lcTerm)
                If IsArray(candidates) Then
                    For c = LBound(candidates) To UBound(candidates)
                        If InStr(1, lowerDesc(candidates(c)), lcTerm, vbTextCompare) > 0 Then
                            found = True: Exit For
                        End If
                    Next c
                Else
                    ' fallback linear search if no candidates (rare)
                    For i = 1 To descCount
                        If InStr(1, lowerDesc(i), lcTerm, vbTextCompare) > 0 Then
                            found = True: Exit For
                        End If
                    Next i
                End If
            Else
                ' Linear scan (early exit)
                For i = 1 To descCount
                    If InStr(1, lowerDesc(i), lcTerm, vbTextCompare) > 0 Then
                        found = True: Exit For
                    End If
                Next i
            End If
            presenceVals(r, 1) = IIf(found, 1, 0)
        End If

        If diag Then
            If (r Mod ProgressInterval = 0) Or (Timer - lastReport > 5) Then
                Debug.Print "[MAPCLEAN] Processed " & r & "/" & termCount & _
                    ", found=" & presenceVals(r, 1) & _
                    ", elapsed=" & Format(Timer - t0, "0.0s")
                lastReport = Timer
                DoEvents
            End If
        End If
    Next r

    ' Write back presence column --------------------------------------
    loMap.DataBodyRange.Columns(presenceIdx).Value = presenceVals
    If diag Then Debug.Print "[MAPCLEAN] Presence column updated. Time=" & Format(Timer - t0, "0.00s")

    ' Optional deletion ------------------------------------------------
    If StrComp(action, "DELETE", vbTextCompare) = 0 Then
        Dim deleted As Long
        For r = termCount To 1 Step -1
            If presenceVals(r, 1) = 0 Then
                loMap.ListRows(r).Delete
                deleted = deleted + 1
                If diag And (deleted Mod 200 = 0) Then Debug.Print "[MAPCLEAN] Deleted " & deleted & " unused rows.";
            End If
        Next r
        If diag Then Debug.Print "[MAPCLEAN] Deleted total unused rows: " & deleted
    End If

    If diag Then Debug.Print "[MAPCLEAN] COMPLETE. Terms=" & termCount & _
        ", Action=" & action & ", TotalTime=" & Format(Timer - t0, "0.00s")

CleanExit:
    RestoreAppState
    Exit Sub
EH:
    Debug.Print "[MAPCLEAN][ERROR] " & Err.Number & " - " & Err.Description
    MsgBox "Mapping cleanup error: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' ============================================================
' Word Index Construction (Optional Acceleration)
' ============================================================
Private Function BuildWordIndex(ByRef lowerDesc() As String, ByVal diag As Boolean) As Object
    On Error GoTo EH
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long, txt As String, tokens As Variant, t As Variant
    For i = LBound(lowerDesc) To UBound(lowerDesc)
        txt = lowerDesc(i)
        If Len(txt) > 0 Then
            tokens = TokenizeSimple(txt)
            For Each t In tokens
                If Not dict.exists(t) Then
                    Dim rowDict As Object: Set rowDict = CreateObject("Scripting.Dictionary")
                    dict.Add t, rowDict
                End If
                If Not dict(t).exists(i) Then dict(t).Add i, True
            Next t
        End If
        If diag Then
            If i Mod 100000 = 0 Then Debug.Print "[MAPCLEAN] Indexing row " & i
        End If
    Next i
    If diag Then Debug.Print "[MAPCLEAN] Word index built. Unique tokens=" & dict.Count
    Set BuildWordIndex = dict
    Exit Function
EH:
    Debug.Print "[MAPCLEAN][WARN] Failed building word index: " & Err.Description
End Function

Private Function CandidateRowsFromWordIndex(ByVal wordIndex As Object, ByVal lcTerm As String) As Variant
    Dim tokens As Variant: tokens = TokenizeSimple(lcTerm)
    Dim t As Variant, firstSet As Object, intersectSet As Object, rowId As Variant

    For Each t In tokens
        If wordIndex.exists(t) Then
            If firstSet Is Nothing Then
                Set firstSet = wordIndex(t)
            Else
                ' Intersect existing set with new token set
                Set intersectSet = CreateObject("Scripting.Dictionary")
                For Each rowId In firstSet.Keys
                    If wordIndex(t).exists(rowId) Then intersectSet.Add rowId, True
                Next rowId
                Set firstSet = intersectSet
                If firstSet.Count = 0 Then Exit For
            End If
        Else
            ' Token absent globally => no candidates
            Exit For
        End If
    Next t

    If Not firstSet Is Nothing Then
        Dim arr() As Long, i As Long
        ReDim arr(1 To firstSet.Count)
        For Each rowId In firstSet.Keys
            i = i + 1
            arr(i) = CLng(rowId)
        Next rowId
        CandidateRowsFromWordIndex = arr
    End If
End Function

Private Function TokenizeSimple(ByVal s As String) As Variant
    ' Splits on any non-alphanumeric; collapses multiples; returns unique tokens
    Dim cleaned As String, i As Long, ch As String * 1
    cleaned = s
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If Not ((ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9")) Then Mid$(cleaned, i, 1) = " "
    Next i
    cleaned = Trim$(cleaned)
    If Len(cleaned) = 0 Then
        TokenizeSimple = Array()
        Exit Function
    End If
    Dim parts As Variant: parts = Split(cleaned, " ")
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Not dict.exists(parts(i)) Then dict.Add parts(i), True
        End If
    Next i
    TokenizeSimple = dict.Keys
End Function

' ============================================================
' Helpers & Integration Safety
' ============================================================
Private Function EnsureListColumn(ByVal loTarget As ListObject, ByVal headerName As String, ByRef newIndex As Long) As ListObject
    On Error GoTo EH
    Dim lc As ListColumn
    For Each lc In loTarget.ListColumns
        If StrComp(CStr(lc.Name), headerName, vbTextCompare) = 0 Then
            newIndex = lc.Index
            Set EnsureListColumn = loTarget
            Exit Function
        End If
    Next lc
    Set lc = loTarget.ListColumns.Add
    lc.Name = headerName
    newIndex = lc.Index
    Set EnsureListColumn = loTarget
    Exit Function
EH:
    Debug.Print "[MAPCLEAN][ERROR] EnsureListColumn: " & Err.Description
End Function

Private Function SafeDataTableName() As String
    On Error Resume Next
    SafeDataTableName = DataTableName()
    If Len(SafeDataTableName) = 0 Then SafeDataTableName = "EquipmentData"
End Function

Private Function SafeMappingTableName() As String
    On Error Resume Next
    SafeMappingTableName = MappingTableName()
    If Len(SafeMappingTableName) = 0 Then SafeMappingTableName = "tbl_Mapping"
End Function

Private Function GetConfigValueSafe(ByVal key As String, ByVal defVal As String) As String
    On Error Resume Next
    Dim v As String: v = GetConfigValue(key)
    If Len(Trim$(v)) = 0 Then v = defVal
    GetConfigValueSafe = v
End Function

Private Function GetDiagnosticModeGlobal() As Boolean
    ' Explicitly reference canonical diagnostic flag to avoid ambiguous public variable collisions
    On Error Resume Next
    GetDiagnosticModeGlobal = CBool(mod_SearchEngine_Enhanced.DiagnosticMode)
End Function

' ============================================================
' Application State Management
' ============================================================
Private Sub SaveAppState()
    On Error Resume Next
    mPrevScreenUpdating = Application.ScreenUpdating
    mPrevEnableEvents = Application.EnableEvents
    mPrevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreAppState()
    On Error Resume Next
    Application.ScreenUpdating = mPrevScreenUpdating
    Application.EnableEvents = mPrevEnableEvents
    Application.Calculation = mPrevCalc
End Sub

' ============================================================
' Initialization Hook (optional future use)
' ============================================================
Public Sub MappingCleanup_SelfTest()
    Debug.Print "[MAPCLEAN] DataTable=" & SafeDataTableName() & _
        " MappingTable=" & SafeMappingTableName()
End Sub

'============================================================
' End of mod_MappingCleanup
'============================================================
