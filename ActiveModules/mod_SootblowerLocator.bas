Attribute VB_Name = "mod_SootblowerLocator"
Option Explicit
'
' Sootblower Locator Mode
' ------------------------------------------------------------
' This module implements a mode-specific UI + search handler for locating
' sootblowers identified by Tag ID pattern: (SSB) <num> <type>.
' Groups:
'   - Retracts: Functional System = "RETRACTS" or type in {SBEL, SBIK}
'   - Wall Blower: Functional System = "WALL BLOWER" or type in {SBIR, SBWB}
' Scope:
'   - Only rows where Functional System Category = "SOOT BLOWING"
' Output:
'   - Writes rows using OutputColumns configured in ConfigTable.
'   - When Show All is requested, sort by Functional System, then Equipment Description.
'
' Integration:
'   - ModeConfigTable.CustomHandler = "Init_SootblowerLocator"
'   - Requires helpers from mod_PrimaryConsolidatedModule (lo, GetConfigValue, HeaderIndexByText, MaxOutputRows, ClearOldResults, WriteStatus).
'
' Error Handling:
'   - Comprehensive diagnostic logging to logs/Diagnostic_Notes/ folder
'   - Detailed error context and recovery suggestions
'   - Validation of all dependencies and prerequisites
'
Private Const FS_CAT_TARGET As String = "SOOT BLOWING"
Private Const TYPE_RETRACTS As String = "Retracts"
Private Const TYPE_WALL As String = "Wall"

' Diagnostic logging constants
Private Const DIAG_LOG_FILE As String = "SootblowerDiagnostics"
Private Const LOG_SEPARATOR As String = vbTab

Public Sub Init_SootblowerLocator()
    ' Simple startup diagnostic to verify function is being called
    MsgBox "Init_SootblowerLocator started", vbInformation, "Debug"
    
    Dim startTime As Double: startTime = Timer
    LogDiagnostic "INFO", "Init_SootblowerLocator", "Starting sootblower locator initialization", ""
    
    On Error GoTo ErrorHandler
    
    ' Validate environment and dependencies first
    If Not ValidateEnvironment() Then
        LogDiagnostic "ERROR", "Init_SootblowerLocator", "Environment validation failed", "Check data tables, config tables, and required functions"
        MsgBox "Environment validation failed. Check diagnostic logs for details.", vbCritical, "Sootblower Locator"
        Exit Sub
    End If
    
    ' Ensure config and mode rows/columns exist per user request
    LogDiagnostic "INFO", "Init_SootblowerLocator", "Ensuring configuration setup", ""
    On Error Resume Next
    Ensure_ConfigKeys_Sootblower
    If Err.Number <> 0 Then
        LogDiagnostic "WARN", "Init_SootblowerLocator", "Config keys setup failed", "Error: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    
    Ensure_ModeConfigEntry_SootblowerLocation
    If Err.Number <> 0 Then
        LogDiagnostic "WARN", "Init_SootblowerLocator", "Mode config setup failed", "Error: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' Optionally auto-create parsed columns based on config toggle
    LogDiagnostic "INFO", "Init_SootblowerLocator", "Checking auto-parse configuration", ""
    Dim autoParse As String: autoParse = UCase$(Trim$(GetConfigValue("SSB_AutoParseColumns")))
    If autoParse = "YES" Or autoParse = "TRUE" Or autoParse = "1" Then
        LogDiagnostic "INFO", "Init_SootblowerLocator", "Auto-creating parsed columns", ""
        EnsureSSBParsedColumns
    End If

    ' Build/update association helpers once on init as well
    LogDiagnostic "INFO", "Init_SootblowerLocator", "Updating association helper columns", ""
    EnsureSSBAssocHelperColumns
    
    ' Try to create/show the modeless UserForm (via factory). If VBIDE access is blocked, fall back.
    LogDiagnostic "INFO", "Init_SootblowerLocator", "Attempting to create UserForm", ""
    If EnsureSootblowerForm() Then
        LogDiagnostic "INFO", "Init_SootblowerLocator", "UserForm available; attempting to show modeless form", ""

        On Error Resume Next
        Dim formCreated As Boolean: formCreated = False
        Dim frm As Object

        ' Prefer dynamic creator (works even without a compiled .frm)
        Set frm = Application.Run("SootblowerFormCreator.CreateSootblowerForm")
        If Not frm Is Nothing Then
            frm.Show vbModeless
            formCreated = True
        Else
            Err.Clear
            ' Fallback to design-time form if present in project
            VBA.UserForms.Add("frmSootblowerLocator").Show vbModeless
            If Err.Number = 0 Then formCreated = True
        End If

        On Error GoTo ErrorHandler

        If formCreated Then
            LogDiagnostic "SUCCESS", "Init_SootblowerLocator", "Initialization completed successfully", "Duration: " & Format(Timer - startTime, "0.00") & " seconds"
            Exit Sub
        End If
    End If

fallback:
    LogDiagnostic "WARN", "Init_SootblowerLocator", "UserForm creation failed, using fallback dialog", ""
    ' Minimal fallback UI if form cannot be created
    Dim resp As VbMsgBoxResult, numTxt As String, grp As String
    numTxt = InputBox("Enter sootblower number (digits only, optional). Leave blank to list all.", "Sootblower Locator")
    If StrPtr(numTxt) = 0 Then ' User cancelled
        LogDiagnostic "INFO", "Init_SootblowerLocator", "User cancelled input dialog", ""
        Exit Sub
    End If
    
    resp = MsgBox("Limit by group? Yes = IK/EL Retracts, No = IR/WB Wall Blower, Cancel = both.", vbYesNoCancel + vbQuestion, "Group Filter")
    If resp = vbYes Then grp = TYPE_RETRACTS ElseIf resp = vbNo Then grp = TYPE_WALL Else grp = ""
    
    LogDiagnostic "INFO", "Init_SootblowerLocator", "Fallback input received", "Number: '" & numTxt & "', Group: '" & grp & "'"
    
    If Len(Trim$(numTxt)) = 0 Then
        SB_DisplayAll grp
    Else
        SB_ExecuteSearch numTxt, grp
    End If
    
    LogDiagnostic "SUCCESS", "Init_SootblowerLocator", "Initialization completed via fallback", "Duration: " & Format(Timer - startTime, "0.00") & " seconds"
    Exit Sub

ErrorHandler:
    ' First attempt: Try standard logging
    On Error Resume Next
    LogDiagnostic "ERROR", "Init_SootblowerLocator", "Unexpected error during initialization", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    
    ' Second attempt: Show error message
    MsgBox "Unexpected error in Init_SootblowerLocator: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
    
    ' Third attempt: Try to write to Desktop as last resort
    Dim errNum As Long, errDesc As String, errLine As String
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl
    
    Dim fso As Object, desktopPath As String, lastResortFile As String, fileStream As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso Is Nothing Then
        desktopPath = Environ$("USERPROFILE") & "\Desktop\"
        lastResortFile = desktopPath & "SootblowerLocator_CriticalError_" & Format(Now, "YYYYMMDD_HHMMSS") & ".log"
        
        Set fileStream = fso.CreateTextFile(lastResortFile, True)
        If Not fileStream Is Nothing Then
            fileStream.WriteLine "CRITICAL ERROR in mod_SootblowerLocator"
            fileStream.WriteLine "Timestamp: " & Now
            fileStream.WriteLine "Function: Init_SootblowerLocator"
            fileStream.WriteLine "Error Number: " & errNum
            fileStream.WriteLine "Description: " & errDesc
            fileStream.WriteLine "Line: " & errLine
            fileStream.WriteLine "User: " & Application.UserName
            fileStream.WriteLine "Computer: " & Environ("COMPUTERNAME")
            fileStream.WriteLine "Workbook: " & ThisWorkbook.FullName
            fileStream.Close
        End If
    End If
    
    LogDiagnostic "ERROR", "Init_SootblowerLocator", "Initialization failed", "Duration: " & Format(Timer - startTime, "0.00") & " seconds"
End Sub

Public Sub SB_ExecuteSearch(ByVal numberText As String, ByVal groupSel As String)
    Dim startTime As Double: startTime = Timer
    LogDiagnostic "INFO", "SB_ExecuteSearch", "Starting search execution", "Number: '" & numberText & "', Group: '" & groupSel & "'"
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Len(Trim$(numberText)) = 0 Then
        LogDiagnostic "WARN", "SB_ExecuteSearch", "Empty number provided", "Redirecting to display all"
        SB_DisplayAll groupSel
        Exit Sub
    End If
    
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Then
        LogDiagnostic "ERROR", "SB_ExecuteSearch", "Data table not found", "Table name: " & DATA_TABLE_NAME
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found. Check your configuration.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If
    
    If dataLo.DataBodyRange Is Nothing Then
        LogDiagnostic "ERROR", "SB_ExecuteSearch", "Data table is empty", "No data rows found"
        MsgBox "Data table '" & DATA_TABLE_NAME & "' is empty.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If

    LogDiagnostic "INFO", "SB_ExecuteSearch", "Data validation passed", "Rows: " & dataLo.DataBodyRange.Rows.Count

    Dim idxs As Variant
    idxs = FindSootblowerMatches(dataLo, numberText, groupSel)

    If IsEmpty(idxs) Then
        LogDiagnostic "INFO", "SB_ExecuteSearch", "No matches found", "Number: " & numberText & ", Group: " & groupSel
        SB_Log "NoMatch", numberText, groupSel, 0, "No sootblower matched."
        If MsgBox("No match found. Try a different search?", vbQuestion + vbYesNo, "Sootblower Locator") = vbYes Then 
            LogDiagnostic "INFO", "SB_ExecuteSearch", "User chose to retry search", ""
            Exit Sub
        End If
        If MsgBox("Display all sootblowers?", vbQuestion + vbYesNo, "Sootblower Locator") = vbYes Then
            LogDiagnostic "INFO", "SB_ExecuteSearch", "User chose to display all", ""
            SB_DisplayAll ""
        End If
        Exit Sub
    End If

    LogDiagnostic "INFO", "SB_ExecuteSearch", "Initial matches found", "Count: " & (UBound(idxs) + 1)

    ' If no group selected and more than one match with the same number across groups, prompt
    If Len(Trim$(groupSel)) = 0 Then
        Dim num As String: num = DigitsOnly(numberText)
        Dim cntRetr As Long, cntWall As Long
        CountByGroup dataLo, idxs, cntRetr, cntWall
        LogDiagnostic "INFO", "SB_ExecuteSearch", "Group count analysis", "Retracts: " & cntRetr & ", Wall: " & cntWall
        
        If cntRetr > 0 And cntWall > 0 Then
            LogDiagnostic "WARN", "SB_ExecuteSearch", "Ambiguous number across multiple groups", "Number: " & num
            SB_Log "AmbiguousNumber", num, groupSel, UBound(idxs), "Multiple groups match this number."
            Dim choose As VbMsgBoxResult
            choose = MsgBox("More than one sootblower uses this number. Choose a group: Yes = IK/EL (Retracts), No = IR/WB (Wall Blower).", vbYesNoCancel + vbQuestion, "Choose Group")
            If choose = vbYes Then 
                groupSel = TYPE_RETRACTS
                LogDiagnostic "INFO", "SB_ExecuteSearch", "User selected Retracts group", ""
            ElseIf choose = vbNo Then 
                groupSel = TYPE_WALL
                LogDiagnostic "INFO", "SB_ExecuteSearch", "User selected Wall group", ""
            ElseIf choose = vbCancel Then 
                LogDiagnostic "INFO", "SB_ExecuteSearch", "User cancelled group selection", ""
                Exit Sub
            End If
            
            ' Recompute with group filter
            LogDiagnostic "INFO", "SB_ExecuteSearch", "Recomputing matches with group filter", "Selected group: " & groupSel
            idxs = FindSootblowerMatches(dataLo, numberText, groupSel)
            If IsEmpty(idxs) Then
                LogDiagnostic "WARN", "SB_ExecuteSearch", "No results in selected group", ""
                MsgBox "No results in selected group.", vbInformation
                Exit Sub
            End If
            LogDiagnostic "INFO", "SB_ExecuteSearch", "Refined matches found", "Count: " & (UBound(idxs) + 1)
        End If
    End If

    OutputRowsToDashboard dataLo, idxs, False
    LogDiagnostic "SUCCESS", "SB_ExecuteSearch", "Search completed successfully", "Results: " & (UBound(idxs) + 1) & ", Duration: " & Format(Timer - startTime, "0.00") & " seconds"
    SB_Log "Output", numberText, groupSel, UBound(idxs), "Displayed matches"
    Exit Sub
    
ErrorHandler:
    LogDiagnostic "ERROR", "SB_ExecuteSearch", "Unexpected error during search", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    SB_Log "Error", numberText, groupSel, 0, CStr(Err.Number) & ": " & Err.Description
    MsgBox "Unexpected error in SB_ExecuteSearch: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
End Sub

Public Sub SB_DisplayAll(ByVal groupSel As String)
    Dim startTime As Double: startTime = Timer
    LogDiagnostic "INFO", "SB_DisplayAll", "Starting display all operation", "Group: '" & groupSel & "'"
    
    On Error GoTo ErrorHandler
    
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Then
        LogDiagnostic "ERROR", "SB_DisplayAll", "Data table not found", "Table name: " & DATA_TABLE_NAME
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found. Check your configuration.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If
    
    If dataLo.DataBodyRange Is Nothing Then
        LogDiagnostic "ERROR", "SB_DisplayAll", "Data table is empty", "No data rows found"
        MsgBox "Data table '" & DATA_TABLE_NAME & "' is empty.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If
    
    LogDiagnostic "INFO", "SB_DisplayAll", "Data validation passed", "Rows: " & dataLo.DataBodyRange.Rows.Count
    
    Dim idxs As Variant
    idxs = FindSootblowerMatches(dataLo, "", groupSel) ' no number filter
    
    If IsEmpty(idxs) Then
        LogDiagnostic "INFO", "SB_DisplayAll", "No sootblowers found", "Group filter: '" & groupSel & "'"
        MsgBox "No sootblowers found" & IIf(Len(groupSel) > 0, " for group '" & groupSel & "'", "") & ".", vbInformation, "Sootblower Locator"
        Exit Sub
    End If
    
    LogDiagnostic "INFO", "SB_DisplayAll", "Sootblowers found", "Count: " & (UBound(idxs) + 1)
    
    OutputRowsToDashboard dataLo, idxs, True
    LogDiagnostic "SUCCESS", "SB_DisplayAll", "Display all completed successfully", "Results: " & (UBound(idxs) + 1) & ", Duration: " & Format(Timer - startTime, "0.00") & " seconds"
    SB_Log "ShowAll", "", groupSel, UBound(idxs), "Displayed all sootblowers"
    Exit Sub
    
ErrorHandler:
    LogDiagnostic "ERROR", "SB_DisplayAll", "Unexpected error during display all", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    SB_Log "Error", "", groupSel, 0, CStr(Err.Number) & ": " & Err.Description
    MsgBox "Unexpected error in SB_DisplayAll: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
End Sub

' Core finder: returns 1-based row indexes relative to DataBodyRange
Private Function FindSootblowerMatches(ByVal dataLo As ListObject, ByVal numberText As String, ByVal groupSel As String) As Variant
    LogDiagnostic "INFO", "FindSootblowerMatches", "Starting search for matches", "Number: '" & numberText & "', Group: '" & groupSel & "'"
    
    On Error GoTo ErrorHandler
    
    ' Get column indices with validation
    Dim fsCatIdx As Long, tagIdx As Long, fsIdx As Long
    fsCatIdx = GetColumnIndex("DataTable_FunctionalSystemCategory", dataLo): If fsCatIdx = 0 Then fsCatIdx = HeaderIndexByText(dataLo, "Functional System Category")
    fsIdx = GetColumnIndex("DataTable_FunctionalSystem", dataLo):        If fsIdx = 0 Then fsIdx = HeaderIndexByText(dataLo, "Functional System")
    tagIdx = GetColumnIndex("DataTable_TagID", dataLo):                    If tagIdx = 0 Then tagIdx = HeaderIndexByText(dataLo, "Tag ID")
    
    If fsCatIdx = 0 Then
        LogDiagnostic "ERROR", "FindSootblowerMatches", "Functional System Category column not found", ""
        Exit Function
    End If
    If fsIdx = 0 Then
        LogDiagnostic "ERROR", "FindSootblowerMatches", "Functional System column not found", ""
        Exit Function
    End If
    If tagIdx = 0 Then
        LogDiagnostic "ERROR", "FindSootblowerMatches", "Tag ID column not found", ""
        Exit Function
    End If
    
    LogDiagnostic "INFO", "FindSootblowerMatches", "Column indices resolved", "FsCat: " & fsCatIdx & ", FS: " & fsIdx & ", Tag: " & tagIdx

    Dim wantNum As String: wantNum = DigitsOnly(numberText)
    Dim n As Long: n = dataLo.DataBodyRange.Rows.Count
    Dim r As Long, keep As Boolean
    Dim fsCat As String, tagText As String, sb As Boolean, num As String, tcode As String, fs As String
    Dim tmpIdx() As Long, cnt As Long
    ReDim tmpIdx(1 To n)

    LogDiagnostic "INFO", "FindSootblowerMatches", "Starting row scan", "Rows to scan: " & n & ", Target number: '" & wantNum & "'"

    For r = 1 To n
        On Error Resume Next
        fsCat = UCase$(Trim$(CStr(dataLo.DataBodyRange.Cells(r, fsCatIdx).Value)))
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "FindSootblowerMatches", "Error reading FunctionalSystemCategory", "Row: " & r & ", Error: " & Err.Description
            Err.Clear
            GoTo NextRow
        End If
        On Error GoTo ErrorHandler
        
        If fsCat <> UCase$(FS_CAT_TARGET) Then GoTo NextRow

        On Error Resume Next
        tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "FindSootblowerMatches", "Error reading Tag ID", "Row: " & r & ", Error: " & Err.Description
            Err.Clear
            GoTo NextRow
        End If
        On Error GoTo ErrorHandler
        
        ParseSSBTag tagText, sb, num, tcode
        If Not sb Then GoTo NextRow

        On Error Resume Next
        fs = UCase$(Trim$(CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value)))
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "FindSootblowerMatches", "Error reading Functional System", "Row: " & r & ", Error: " & Err.Description
            Err.Clear
            fs = ""
        End If
        On Error GoTo ErrorHandler
        
        keep = True

        ' Number filter (exact match on middle segment if provided)
        If Len(wantNum) > 0 Then
            If StrComp(num, wantNum, vbTextCompare) <> 0 Then keep = False
        End If

        ' Group filter if provided
        If keep And Len(Trim$(groupSel)) > 0 Then
            If StrComp(groupSel, TYPE_RETRACTS, vbTextCompare) = 0 Then
                keep = IsRetracts(fs, tcode)
            ElseIf StrComp(groupSel, TYPE_WALL, vbTextCompare) = 0 Then
                keep = IsWall(fs, tcode)
            End If
        End If

        If keep Then
            cnt = cnt + 1
            tmpIdx(cnt) = r
        End If
NextRow:
    Next r

    LogDiagnostic "INFO", "FindSootblowerMatches", "Row scan completed", "Matches found: " & cnt

    If cnt = 0 Then 
        LogDiagnostic "INFO", "FindSootblowerMatches", "No matches found", ""
        Exit Function
    End If
    
    ReDim Preserve tmpIdx(1 To cnt)
    FindSootblowerMatches = tmpIdx
    
    LogDiagnostic "SUCCESS", "FindSootblowerMatches", "Matches found and returned", "Count: " & cnt
    Exit Function
    
ErrorHandler:
    LogDiagnostic "ERROR", "FindSootblowerMatches", "Unexpected error during search", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    ' Return empty result on error
End Function

Private Function IsRetracts(ByVal fs As String, ByVal tcode As String) As Boolean
    fs = UCase$(Trim$(fs)): tcode = UCase$(Trim$(tcode))
    IsRetracts = (fs = "RETRACTS" Or tcode = "SBEL" Or tcode = "SBIK")
End Function

Private Function IsWall(ByVal fs As String, ByVal tcode As String) As Boolean
    fs = UCase$(Trim$(fs)): tcode = UCase$(Trim$(tcode))
    IsWall = (fs = "WALL BLOWER" Or tcode = "SBIR" Or tcode = "SBWB")
End Function

Private Sub CountByGroup(ByVal dataLo As ListObject, ByVal idxs As Variant, ByRef outRetr As Long, ByRef outWall As Long)
    Dim fsIdx As Long, tagIdx As Long
    fsIdx = GetColumnIndex("DataTable_FunctionalSystem", dataLo): If fsIdx = 0 Then fsIdx = HeaderIndexByText(dataLo, "Functional System")
    tagIdx = GetColumnIndex("DataTable_TagID", dataLo):                  If tagIdx = 0 Then tagIdx = HeaderIndexByText(dataLo, "Tag ID")
    Dim i As Long, r As Long, fs As String, tagText As String, sb As Boolean, num As String, tcode As String
    For i = 1 To UBound(idxs)
        r = idxs(i)
        fs = CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value)
        tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
        ParseSSBTag tagText, sb, num, tcode
        If IsRetracts(fs, tcode) Then outRetr = outRetr + 1 ElseIf IsWall(fs, tcode) Then outWall = outWall + 1
    Next i
End Sub

Private Sub OutputRowsToDashboard(ByVal dataLo As ListObject, ByVal rowIdxs As Variant, ByVal doSortAll As Boolean)
    LogDiagnostic "INFO", "OutputRowsToDashboard", "Starting output to dashboard", "Rows: " & (UBound(rowIdxs) + 1) & ", Sort: " & doSortAll
    
    On Error GoTo ErrorHandler
    
    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = nr(GetConfigValue("ResultsStartCell"))
    Set statusRng = nr(GetConfigValue("StatusCell"))
    If resultsStart Is Nothing Then
        LogDiagnostic "ERROR", "OutputRowsToDashboard", "ResultsStartCell range not found", "Config value: " & GetConfigValue("ResultsStartCell")
        MsgBox "Named range 'ResultsStartCell' not found. Check your configuration.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If
    
    LogDiagnostic "INFO", "OutputRowsToDashboard", "Output ranges resolved", "Results: " & resultsStart.Address & ", Status: " & IIf(statusRng Is Nothing, "None", statusRng.Address)

    ' Build dynamic output columns from Config
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
            Else
                LogDiagnostic "WARN", "OutputRowsToDashboard", "Output column not found in data", "Key: " & key & ", Header: " & headerName
            End If
        End If
    Next i
    
    LogDiagnostic "INFO", "OutputRowsToDashboard", "Output columns resolved", "Valid columns: " & colCount
    
    If colCount = 0 Then
        LogDiagnostic "ERROR", "OutputRowsToDashboard", "No valid output columns found", "Check ConfigTable output column configuration"
        MsgBox "No valid output columns found in ConfigTable. Check your configuration.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If
    ReDim Preserve outCols(1 To colCount)

    ' Headers
    Dim hdr() As Variant
    ReDim hdr(1 To 1, 1 To colCount)
    For i = 1 To colCount
        hdr(1, i) = CStr(dataLo.HeaderRowRange.Cells(1, outCols(i)).Value)
    Next i
    resultsStart.Resize(1, colCount).Value = hdr
    
    LogDiagnostic "INFO", "OutputRowsToDashboard", "Headers written", "Columns: " & colCount

    ' Clear old results
    ClearOldResults resultsStart, colCount

    Dim maxRows As Long: maxRows = MaxOutputRows()
    Dim take As Long
    If maxRows > 0 Then
        If UBound(rowIdxs) < maxRows Then take = UBound(rowIdxs) Else take = maxRows
        LogDiagnostic "INFO", "OutputRowsToDashboard", "Row limit applied", "Max: " & maxRows & ", Taking: " & take
    Else
        take = UBound(rowIdxs)
        LogDiagnostic "INFO", "OutputRowsToDashboard", "No row limit", "Taking all: " & take
    End If

    Dim outArr() As Variant, ri As Long, j As Long
    ReDim outArr(1 To take, 1 To colCount)
    For i = 1 To take
        ri = rowIdxs(i)
        On Error Resume Next
        For j = 1 To colCount
            outArr(i, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).Value)
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "OutputRowsToDashboard", "Error reading cell data", "Row: " & ri & ", Col: " & outCols(j) & ", Error: " & Err.Description
                outArr(i, j) = "#READ_ERROR#"
                Err.Clear
            End If
        Next j
        On Error GoTo ErrorHandler
    Next i
    
    LogDiagnostic "INFO", "OutputRowsToDashboard", "Data array populated", "Rows: " & take & ", Columns: " & colCount

    If doSortAll And take > 1 Then
        LogDiagnostic "INFO", "OutputRowsToDashboard", "Applying sort", "Sorting by Functional System and Equipment Description"
        Dim fsIdx As Long, descIdx As Long
        fsIdx = GetColumnIndex("DataTable_FunctionalSystem", dataLo): If fsIdx = 0 Then fsIdx = HeaderIndexByText(dataLo, "Functional System")
        descIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)
        Dim fsOutPos As Long: fsOutPos = FindOutPos(outCols, colCount, fsIdx)
        Dim descOutPos As Long: descOutPos = FindOutPos(outCols, colCount, descIdx)
        ' Two-key sort: stable approach by sorting by second key first, then first key
        On Error Resume Next
        If descOutPos > 0 Then 
            QuickSort2D_N outArr, 1, take, descOutPos
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "OutputRowsToDashboard", "Sort by description failed", "Error: " & Err.Description
                Err.Clear
            End If
        End If
        If fsOutPos > 0 Then 
            QuickSort2D_N outArr, 1, take, fsOutPos
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "OutputRowsToDashboard", "Sort by functional system failed", "Error: " & Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo ErrorHandler
        LogDiagnostic "INFO", "OutputRowsToDashboard", "Sort completed", ""
    End If

    resultsStart.Offset(1, 0).Resize(take, colCount).Value = outArr
    
    On Error Resume Next
    If Not statusRng Is Nothing Then
        WriteStatus statusRng, "Displayed " & take & " row(s).", "Sootblower Locator"
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "OutputRowsToDashboard", "Status write failed", "Error: " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo ErrorHandler
    
    LogDiagnostic "SUCCESS", "OutputRowsToDashboard", "Output completed successfully", "Displayed: " & take & " rows"
    Exit Sub
    
ErrorHandler:
    LogDiagnostic "ERROR", "OutputRowsToDashboard", "Unexpected error during output", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "Error in OutputRowsToDashboard: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
End Sub

Private Function FindOutPos(ByRef outCols() As Long, ByVal colCount As Long, ByVal dataColIdx As Long) As Long
    Dim p As Long
    If dataColIdx <= 0 Then Exit Function
    For p = 1 To colCount
        If outCols(p) = dataColIdx Then FindOutPos = p: Exit Function
    Next p
End Function

Public Sub ParseSSBTag(ByVal s As String, ByRef isSB As Boolean, ByRef num As String, ByRef tcode As String)
    isSB = False: num = "": tcode = ""
    Dim rx As Object, m As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False: rx.IgnoreCase = True
    rx.pattern = "^\(SSB\)\s*(\d{1,3})\s+([A-Za-z0-9_\-]+)"
    Dim ms As Object: Set ms = rx.Execute(CStr(s))
    If Not ms Is Nothing And ms.count > 0 Then
        Set m = ms(0)
        isSB = True
        num = m.SubMatches(0)
        tcode = UCase$(m.SubMatches(1))
    End If
End Sub

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Private Sub SB_Log(ByVal action As String, ByVal numberText As String, ByVal groupSel As String, ByVal count As Long, ByVal msg As String)
    On Error Resume Next
    Dim ws As Worksheet, nr As Long
    Set ws = ThisWorkbook.Worksheets("SootblowerLog")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "SootblowerLog"
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Action"
        ws.Cells(1, 3).Value = "Number"
        ws.Cells(1, 4).Value = "Group"
        ws.Cells(1, 5).Value = "Count"
        ws.Cells(1, 6).Value = "Message"
    End If
    nr = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(nr, 1).Value = Now
    ws.Cells(nr, 2).Value = action
    ws.Cells(nr, 3).Value = Trim$(numberText)
    ws.Cells(nr, 4).Value = groupSel
    ws.Cells(nr, 5).Value = count
    ws.Cells(nr, 6).Value = msg
End Sub

 ' NOTE: Duplicate EnsureSootblowerForm removed. See the comprehensive
 ' Private Function EnsureSootblowerForm further below for the canonical
 ' implementation (checks existing .frm and falls back to dynamic creator).

Public Sub SB_ShowAssociated(ByVal numberText As String, ByVal groupSel As String)
    Dim startTime As Double: startTime = Timer
    LogDiagnostic "INFO", "SB_ShowAssociated", "Starting association display", "Number: '" & numberText & "', Group: '" & groupSel & "'"
    
    On Error GoTo ErrorHandler
    
    ' Resolve target sootblower set and display associated equipment using helper columns
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        LogDiagnostic "ERROR", "SB_ShowAssociated", "Data table not found or empty", "Table: " & DATA_TABLE_NAME
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If

    Dim assocMode As String: assocMode = GetConfigValue("SSB_Assoc_Mode")
    Dim assocMax As Long: assocMax = CLngSafe(GetConfigValue("SSB_Assoc_MaxRows"))
    Dim grp As String: grp = Trim$(groupSel)
    Dim num As String: num = DigitsOnly(numberText)
    
    LogDiagnostic "INFO", "SB_ShowAssociated", "Configuration loaded", "Mode: " & assocMode & ", MaxRows: " & assocMax & ", ProcessedNum: " & num

    ' Ensure helpers exist
    LogDiagnostic "INFO", "SB_ShowAssociated", "Ensuring helper columns exist", ""
    EnsureSSBParsedColumns
    EnsureSSBAssocHelperColumns

    ' If number empty, prompt to pick from displayed or ask for a number
    If Len(num) = 0 Then
        LogDiagnostic "INFO", "SB_ShowAssociated", "No number provided, prompting user", ""
        Dim resp As VbMsgBoxResult
        resp = MsgBox("No number entered. Show all associated rows for all displayed sootblowers?", vbYesNoCancel + vbQuestion, "Associated Equipment")
        If resp = vbCancel Then 
            LogDiagnostic "INFO", "SB_ShowAssociated", "User cancelled operation", ""
            Exit Sub
        End If
        If resp = vbNo Then 
            LogDiagnostic "INFO", "SB_ShowAssociated", "User chose to refine search first", ""
            Exit Sub
        End If
    End If

    ' Find source sootblowers (to derive numbers and groups)
    LogDiagnostic "INFO", "SB_ShowAssociated", "Finding source sootblowers", "Number: " & num & ", Group: " & grp
    Dim srcIdxs As Variant: srcIdxs = FindSootblowerMatches(dataLo, num, grp)
    If IsEmpty(srcIdxs) Then
        LogDiagnostic "INFO", "SB_ShowAssociated", "No source sootblowers found", ""
        MsgBox "No sootblowers found to associate.", vbInformation
        Exit Sub
    End If
    
    LogDiagnostic "INFO", "SB_ShowAssociated", "Source sootblowers found", "Count: " & (UBound(srcIdxs) + 1)

    Dim fsHeader As String: fsHeader = Nz(GetConfigValue("DataTable_FunctionalSystem"), "Functional System")
    Dim tagHeader As String: tagHeader = Nz(GetConfigValue("DataTable_TagID"), "Tag ID")
    Dim fsIdx As Long: fsIdx = HeaderIndexByText(dataLo, fsHeader)
    Dim tagIdx As Long: tagIdx = HeaderIndexByText(dataLo, tagHeader)

    Dim assocNumCol As String: assocNumCol = Nz(GetConfigValue("SSB_AssocHelper_NumberCol"), "Assoc SSB Number")
    Dim assocCatCol As String: assocCatCol = Nz(GetConfigValue("SSB_AssocHelper_CategoryCol"), "Assoc SSB Category")
    Dim assocNumIdx As Long: assocNumIdx = HeaderIndexByText(dataLo, assocNumCol)
    Dim assocCatIdx As Long: assocCatIdx = HeaderIndexByText(dataLo, assocCatCol)
    
    If assocNumIdx = 0 Or assocCatIdx = 0 Then
        LogDiagnostic "ERROR", "SB_ShowAssociated", "Association helper columns not found", "Number col: " & assocNumIdx & ", Category col: " & assocCatIdx
        MsgBox "Association helper columns not found. Please run the setup first.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If

    ' Build a set of target numbers and group preferences
    Dim targetNums As Object: Set targetNums = CreateObject("Scripting.Dictionary")
    targetNums.CompareMode = vbTextCompare
    Dim preferredGroup As String: preferredGroup = Trim$(grp)

    Dim i As Long, r As Long, isSB As Boolean, sbNum As String, tcode As String, fs As String
    LogDiagnostic "INFO", "SB_ShowAssociated", "Building target number set", ""
    
    For i = 1 To UBound(srcIdxs)
        r = srcIdxs(i)
        On Error Resume Next
        ParseSSBTag CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value), isSB, sbNum, tcode
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "SB_ShowAssociated", "Error parsing source row", "Row: " & r & ", Error: " & Err.Description
            Err.Clear
            GoTo NextSourceRow
        End If
        On Error GoTo ErrorHandler
        
        If isSB Then
            targetNums(sbNum) = True
            If Len(preferredGroup) = 0 Then
                On Error Resume Next
                fs = UCase$(CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value))
                If Err.Number <> 0 Then
                    fs = ""
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                If IsRetracts(fs, tcode) Then preferredGroup = "Retracts"
                If IsWall(fs, tcode) Then preferredGroup = IIf(Len(preferredGroup) = 0, "Wall", preferredGroup)
            End If
        End If
NextSourceRow:
    Next i
    
    LogDiagnostic "INFO", "SB_ShowAssociated", "Target numbers built", "Count: " & targetNums.Count & ", Preferred group: " & preferredGroup

    ' Scan helper columns to select associated rows: match on number, then choose closest group
    Dim n As Long: n = dataLo.DataBodyRange.Rows.Count
    Dim pickIdx() As Long, cntPick As Long
    ReDim pickIdx(1 To n)
    Dim assocNum As String, assocCat As String
    
    LogDiagnostic "INFO", "SB_ShowAssociated", "Scanning for associated rows", "Max rows: " & assocMax
    
    For r = 1 To n
        On Error Resume Next
        assocNum = CStr(dataLo.DataBodyRange.Cells(r, assocNumIdx).Value)
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "SB_ShowAssociated", "Error reading association number", "Row: " & r & ", Error: " & Err.Description
            Err.Clear
            GoTo NextAssocScanRow
        End If
        On Error GoTo ErrorHandler
        
        If Len(assocNum) > 0 Then
            If targetNums.exists(assocNum) Then
                ' group resolution: prefer preferredGroup if set; else accept any
                On Error Resume Next
                assocCat = CStr(dataLo.DataBodyRange.Cells(r, assocCatIdx).Value)
                If Err.Number <> 0 Then
                    assocCat = ""
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
                If Len(preferredGroup) = 0 Or StrComp(assocCat, preferredGroup, vbTextCompare) = 0 Then
                    cntPick = cntPick + 1
                    pickIdx(cntPick) = r
                    If assocMax > 0 And cntPick >= assocMax Then Exit For
                End If
            End If
        End If
NextAssocScanRow:
    Next r

    If cntPick = 0 Then
        LogDiagnostic "INFO", "SB_ShowAssociated", "No associated rows found", ""
        SB_Log "Assoc", num, grp, 0, "No associated rows"
        MsgBox "No associated equipment found.", vbInformation
        Exit Sub
    End If

    LogDiagnostic "INFO", "SB_ShowAssociated", "Associated rows found", "Count: " & cntPick
    
    ReDim Preserve pickIdx(1 To cntPick)
    OutputRowsToDashboard dataLo, pickIdx, True ' sorted view helpful here
    LogDiagnostic "SUCCESS", "SB_ShowAssociated", "Association display completed", "Results: " & cntPick & ", Duration: " & Format(Timer - startTime, "0.00") & " seconds"
    SB_Log "Assoc", num, grp, cntPick, "Displayed associated equipment"
    Exit Sub
    
ErrorHandler:
    LogDiagnostic "ERROR", "SB_ShowAssociated", "Unexpected error during association display", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    SB_Log "Error", numberText, groupSel, 0, CStr(Err.Number) & ": " & Err.DESCRIPTION
    MsgBox "Unexpected error in SB_ShowAssociated: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
End Sub

Public Sub EnsureSSBParsedColumns()
    ' Create/refresh three supporting parsed columns next to [Tag ID] for rows beginning with (SSB)
    LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Starting parsed columns setup", ""
    
    On Error GoTo ErrorHandler
    
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        LogDiagnostic "ERROR", "EnsureSSBParsedColumns", "Data table not found or empty", "Table: " & DATA_TABLE_NAME
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If

    Dim tagHeader As String: tagHeader = GetConfigValue("DataTable_TagID"): If Len(Trim$(tagHeader)) = 0 Then tagHeader = "Tag ID"
    Dim tagIdx As Long: tagIdx = HeaderIndexByText(dataLo, tagHeader)
    If tagIdx = 0 Then 
        LogDiagnostic "ERROR", "EnsureSSBParsedColumns", "Tag ID column not found", "Header: " & tagHeader
        MsgBox "Tag ID column not found.", vbExclamation, "Sootblower Locator"
        Exit Sub
    End If
    
    LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Tag ID column found", "Index: " & tagIdx

    Dim colPrefix As String: colPrefix = Nz(GetConfigValue("SSB_ParsedPrefixCol"), "SSB Prefix")
    Dim colNumber As String: colNumber = Nz(GetConfigValue("SSB_ParsedNumberCol"), "SSB Number")
    Dim colType As String:   colType = Nz(GetConfigValue("SSB_ParsedTypeCol"), "SSB Type")

    ' Ensure columns exist directly to the right of Tag ID, creating if missing
    Dim c As Long, needNames As Variant
    needNames = Array(colPrefix, colNumber, colType)

    ' If headers present somewhere else, leave them; otherwise insert new columns after Tag ID
    Dim haveAll As Boolean: haveAll = True
    For c = LBound(needNames) To UBound(needNames)
        If HeaderIndexByText(dataLo, CStr(needNames(c))) = 0 Then haveAll = False
    Next c
    
    LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Column existence check", "Have all columns: " & haveAll

    If Not haveAll Then
        LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Inserting missing columns", "After Tag ID index: " & tagIdx
        ' Insert three columns after Tag ID
        On Error Resume Next
        dataLo.ListColumns(tagIdx).Range.Offset(0, 1).Resize(1, 3).EntireColumn.Insert
        If Err.Number <> 0 Then
            LogDiagnostic "ERROR", "EnsureSSBParsedColumns", "Failed to insert columns", "Error: " & Err.Description
            MsgBox "Failed to insert parsed columns: " & Err.Description, vbCritical, "Sootblower Locator"
            Exit Sub
        End If
        
        dataLo.HeaderRowRange.Cells(1, tagIdx + 1).Value = colPrefix
        dataLo.HeaderRowRange.Cells(1, tagIdx + 2).Value = colNumber
        dataLo.HeaderRowRange.Cells(1, tagIdx + 3).Value = colType
        
        ' Resize table to include new columns
        dataLo.Resize dataLo.Range.Resize(dataLo.Range.Rows.Count, dataLo.Range.Columns.Count + 3)
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "EnsureSSBParsedColumns", "Table resize warning", "Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Columns inserted successfully", ""
    End If

    ' Re-resolve indices after potential insert
    Dim prefixIdx As Long: prefixIdx = HeaderIndexByText(dataLo, colPrefix)
    Dim numberIdx As Long: numberIdx = HeaderIndexByText(dataLo, colNumber)
    Dim typeIdx As Long:   typeIdx = HeaderIndexByText(dataLo, colType)
    
    LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Final column indices", "Prefix: " & prefixIdx & ", Number: " & numberIdx & ", Type: " & typeIdx

    ' Fill parsed values for rows with (SSB) pattern only
    Dim n As Long: n = dataLo.DataBodyRange.Rows.Count
    Dim r As Long, tagText As String, isSB As Boolean, num As String, tcode As String
    Dim processedCount As Long: processedCount = 0
    
    LogDiagnostic "INFO", "EnsureSSBParsedColumns", "Starting row processing", "Total rows: " & n
    
    For r = 1 To n
        On Error Resume Next
        tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
        If Err.Number <> 0 Then
            LogDiagnostic "WARN", "EnsureSSBParsedColumns", "Error reading tag text", "Row: " & r & ", Error: " & Err.Description
            Err.Clear
            GoTo NextParseRow
        End If
        On Error GoTo ErrorHandler
        
        ParseSSBTag tagText, isSB, num, tcode
        If isSB Then
            On Error Resume Next
            dataLo.DataBodyRange.Cells(r, prefixIdx).Value = "(SSB)"
            dataLo.DataBodyRange.Cells(r, numberIdx).Value = num
            dataLo.DataBodyRange.Cells(r, typeIdx).Value = tcode
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "EnsureSSBParsedColumns", "Error writing parsed values", "Row: " & r & ", Error: " & Err.Description
                Err.Clear
            Else
                processedCount = processedCount + 1
            End If
            On Error GoTo ErrorHandler
        End If
NextParseRow:
    Next r

    LogDiagnostic "SUCCESS", "EnsureSSBParsedColumns", "Parsed columns setup completed", "Processed SSB rows: " & processedCount
    MsgBox "SSB parsed columns updated. Processed " & processedCount & " SSB rows.", vbInformation, "Sootblower Locator"
    Exit Sub
    
ErrorHandler:
    LogDiagnostic "ERROR", "EnsureSSBParsedColumns", "Unexpected error during parsed columns setup", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "Error in EnsureSSBParsedColumns: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
End Sub

Private Function Nz(ByVal s As String, ByVal fallback As String) As String
    If Len(Trim$(s)) = 0 Then Nz = fallback Else Nz = s
End Function

Public Sub EnsureSSBAssocHelperColumns()
    LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "Starting association helper columns setup", ""
    
    On Error GoTo ErrorHandler
    
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then 
        LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Data table not found or empty", "Table: " & DATA_TABLE_NAME
        Exit Sub
    End If

    Dim catHeader As String: catHeader = Nz(GetConfigValue("DataTable_FunctionalSystemCategory"), "Functional System Category")
    Dim fsHeader As String: fsHeader = Nz(GetConfigValue("DataTable_FunctionalSystem"), "Functional System")
    Dim tagHeader As String: tagHeader = Nz(GetConfigValue("DataTable_TagID"), "Tag ID")
    Dim descHeader As String: descHeader = Nz(GetConfigValue("DataTable_EquipDescription"), "Equipment Description")
    Dim catVal As String: catVal = Nz(GetConfigValue("SSB_Assoc_FilterCategory"), Nz(GetConfigValue("SSB_FunctionalSystemCategoryValue"), "SOOT BLOWING"))

    Dim assocNumCol As String: assocNumCol = Nz(GetConfigValue("SSB_AssocHelper_NumberCol"), "Assoc SSB Number")
    Dim assocCatCol As String: assocCatCol = Nz(GetConfigValue("SSB_AssocHelper_CategoryCol"), "Assoc SSB Category")

    Dim catIdx As Long: catIdx = HeaderIndexByText(dataLo, catHeader)
    Dim fsIdx As Long: fsIdx = HeaderIndexByText(dataLo, fsHeader)
    Dim tagIdx As Long: tagIdx = HeaderIndexByText(dataLo, tagHeader)
    Dim descIdx As Long: descIdx = HeaderIndexByText(dataLo, descHeader)
    
    If catIdx = 0 Then
        LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Functional System Category column not found", "Header: " & catHeader
        Exit Sub
    End If
    If fsIdx = 0 Then
        LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Functional System column not found", "Header: " & fsHeader
        Exit Sub
    End If
    If tagIdx = 0 Then
        LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Tag ID column not found", "Header: " & tagHeader
        Exit Sub
    End If
    If descIdx = 0 Then
        LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Equipment Description column not found", "Header: " & descHeader
        Exit Sub
    End If
    
    LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "Required columns validated", "Cat: " & catIdx & ", FS: " & fsIdx & ", Tag: " & tagIdx & ", Desc: " & descIdx

    ' Ensure helper columns exist (append to far right if missing)
    Dim needNames As Variant: needNames = Array(assocNumCol, assocCatCol)
    Dim i As Long
    For i = LBound(needNames) To UBound(needNames)
        If HeaderIndexByText(dataLo, CStr(needNames(i))) = 0 Then
            LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "Adding missing helper column", "Column: " & CStr(needNames(i))
            On Error Resume Next
            dataLo.ListColumns.Add
            If Err.Number <> 0 Then
                LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Failed to add column", "Column: " & CStr(needNames(i)) & ", Error: " & Err.Description
                Exit Sub
            End If
            dataLo.HeaderRowRange.Cells(1, dataLo.ListColumns.count).Value = CStr(needNames(i))
            dataLo.Resize dataLo.Range.Resize(dataLo.Range.Rows.Count, dataLo.Range.Columns.Count)
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "EnsureSSBAssocHelperColumns", "Table resize warning", "Error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next i

    Dim assocNumIdx As Long: assocNumIdx = HeaderIndexByText(dataLo, assocNumCol)
    Dim assocCatIdx As Long: assocCatIdx = HeaderIndexByText(dataLo, assocCatCol)
    
    LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "Helper column indices", "Number: " & assocNumIdx & ", Category: " & assocCatIdx

    ' Pre-gather SSB numbers from (SSB) Tag rows to speed lookups
    Dim n As Long: n = dataLo.DataBodyRange.Rows.Count
    Dim ssbNumbers As Object: Set ssbNumbers = CreateObject("Scripting.Dictionary")
    ssbNumbers.CompareMode = vbTextCompare
    Dim r As Long, tagText As String, isSB As Boolean, num As String, tcode As String
    
    LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "Gathering SSB numbers", "Scanning " & n & " rows"
    
    For r = 1 To n
        On Error Resume Next
        If StrComp(CStr(dataLo.DataBodyRange.Cells(r, catIdx).Value), catVal, vbTextCompare) = 0 Then
            tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "EnsureSSBAssocHelperColumns", "Error reading row data", "Row: " & r & ", Error: " & Err.Description
                Err.Clear
                GoTo NextGatherRow
            End If
            On Error GoTo ErrorHandler
            ParseSSBTag tagText, isSB, num, tcode
            If isSB Then If Not ssbNumbers.exists(num) Then ssbNumbers.Add num, True
        End If
NextGatherRow:
    Next r
    
    LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "SSB numbers gathered", "Unique numbers: " & ssbNumbers.Count

    ' Keyword groups from config
    Dim kwRetr As Variant, kwWall As Variant
    kwRetr = Split(Nz(GetConfigValue("SSB_AssocKeywords_Retracts"), "IK,EL,RETRACT"), ",")
    kwWall = Split(Nz(GetConfigValue("SSB_AssocKeywords_Wall"), "IR,WB,WALL,WATER"), ",")

    ' Walk all rows and populate helper columns for those in the target category
    Dim descText As String, udesc As String, foundNum As String, catGuess As String, k As Long
    Dim processedCount As Long: processedCount = 0
    
    LogDiagnostic "INFO", "EnsureSSBAssocHelperColumns", "Starting association processing", "Target category: " & catVal
    
    For r = 1 To n
        On Error Resume Next
        If StrComp(CStr(dataLo.DataBodyRange.Cells(r, catIdx).Value), catVal, vbTextCompare) = 0 Then
            ' number detection: exact token match for any known SSB number
            descText = CStr(dataLo.DataBodyRange.Cells(r, descIdx).Value)
            If Err.Number <> 0 Then
                LogDiagnostic "WARN", "EnsureSSBAssocHelperColumns", "Error reading description", "Row: " & r & ", Error: " & Err.Description
                Err.Clear
                GoTo NextAssocRow
            End If
            On Error GoTo ErrorHandler
            
            foundNum = ""
            For Each num In ssbNumbers.Keys
                If HasNumberToken(descText, CStr(num)) Then
                    foundNum = CStr(num): Exit For
                End If
            Next num

            If Len(foundNum) > 0 Then
                On Error Resume Next
                dataLo.DataBodyRange.Cells(r, assocNumIdx).Value = foundNum
                If Err.Number <> 0 Then
                    LogDiagnostic "WARN", "EnsureSSBAssocHelperColumns", "Error writing number", "Row: " & r & ", Error: " & Err.Description
                    Err.Clear
                    GoTo NextAssocRow
                End If
                On Error GoTo ErrorHandler
                
                ' category guess via keywords
                udesc = UCase$(descText)
                catGuess = ""
                For k = LBound(kwRetr) To UBound(kwRetr)
                    If InStr(1, udesc, UCase$(Trim$(CStr(kwRetr(k)))), vbTextCompare) > 0 Then catGuess = "Retracts": Exit For
                Next k
                If Len(catGuess) = 0 Then
                    For k = LBound(kwWall) To UBound(kwWall)
                        If InStr(1, udesc, UCase$(Trim$(CStr(kwWall(k)))), vbTextCompare) > 0 Then catGuess = "Wall": Exit For
                    Next k
                End If
                If Len(catGuess) = 0 Then
                    ' fall back to Functional System name heuristic
                    On Error Resume Next
                    Dim fs As String: fs = UCase$(CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value))
                    If Err.Number <> 0 Then
                        fs = ""
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                    If InStr(1, fs, "RETRACT", vbTextCompare) > 0 Then catGuess = "Retracts"
                    If InStr(1, fs, "WALL", vbTextCompare) > 0 Or InStr(1, fs, "WATER", vbTextCompare) > 0 Then catGuess = "Wall"
                End If
                If Len(catGuess) > 0 Then 
                    On Error Resume Next
                    dataLo.DataBodyRange.Cells(r, assocCatIdx).Value = catGuess
                    If Err.Number <> 0 Then
                        LogDiagnostic "WARN", "EnsureSSBAssocHelperColumns", "Error writing category", "Row: " & r & ", Error: " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                End If
                processedCount = processedCount + 1
            End If
        End If
NextAssocRow:
    Next r
    
    LogDiagnostic "SUCCESS", "EnsureSSBAssocHelperColumns", "Association helper columns setup completed", "Processed associations: " & processedCount
    Exit Sub
    
ErrorHandler:
    LogDiagnostic "ERROR", "EnsureSSBAssocHelperColumns", "Unexpected error during association setup", "Error: " & Err.Number & " - " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "Error in EnsureSSBAssocHelperColumns: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Check diagnostic logs for details.", vbCritical, "Sootblower Locator Error"
End Sub

Private Function HasNumberToken(ByVal text As String, ByVal num As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False: rx.IgnoreCase = True
    rx.pattern = "(^|[^0-9])" & num & "($|[^0-9])"
    HasNumberToken = rx.Test(CStr(text))
End Function

' ============================================================================
' COMPREHENSIVE DIAGNOSTIC LOGGING AND VALIDATION SYSTEM
' ============================================================================

Private Function EnsureSootblowerForm() As Boolean
    ' Function to ensure the Sootblower Form exists or can be created
    ' Returns True if the form is available, False otherwise
    
    On Error Resume Next
    LogDiagnostic "INFO", "EnsureSootblowerForm", "Checking for UserForm availability", ""
    
    ' First try to reference the form directly - if it exists in the project
    Dim testForm As Object
    Set testForm = VBA.UserForms.Add("frmSootblowerLocator")
    
    If Err.Number = 0 Then
        ' Form exists in the project
        testForm.Hide  ' Hide the test instance
        LogDiagnostic "INFO", "EnsureSootblowerForm", "UserForm exists in project", ""
        EnsureSootblowerForm = True
        Exit Function
    End If
    
    Err.Clear
    
    ' Next, try to use the SootblowerFormCreator if available
    If FormCreatorAvailable() Then
        LogDiagnostic "INFO", "EnsureSootblowerForm", "Using SootblowerFormCreator", ""
        EnsureSootblowerForm = True  ' The form will be created when needed
        Exit Function
    End If
    
    ' Neither option available, log the issue
    LogDiagnostic "WARN", "EnsureSootblowerForm", "Form not available", "Will use fallback dialog"
    EnsureSootblowerForm = False
End Function

Private Function FormCreatorAvailable() As Boolean
    ' Check if the dynamic form creator function is callable
    On Error GoTo Nope
    Dim obj As Object
    Set obj = Application.Run("SootblowerFormCreator.CreateSootblowerForm")
    ' If we got here without error, the procedure exists
    FormCreatorAvailable = True
    Exit Function
Nope:
    FormCreatorAvailable = False
End Function

Private Function ValidateEnvironment() As Boolean
    ' Debug verification that this function is being called
    MsgBox "ValidateEnvironment function called", vbInformation, "Debug"
    
    ' Comprehensive environment validation with detailed diagnostics
    LogDiagnostic "INFO", "ValidateEnvironment", "Starting environment validation", ""
    
    On Error GoTo ValidationError
    
    ' Check if DATA_TABLE_NAME constant is defined and accessible
    Dim tableName As String
    On Error Resume Next
    tableName = DATA_TABLE_NAME
    On Error GoTo ValidationError
    If Len(tableName) = 0 Then
        LogDiagnostic "ERROR", "ValidateEnvironment", "DATA_TABLE_NAME constant not defined", ""
        ValidateEnvironment = False
        Exit Function
    End If
    LogDiagnostic "INFO", "ValidateEnvironment", "DATA_TABLE_NAME validated", "Value: " & tableName
    
    ' Check if data table exists and is accessible
    Dim dataLo As ListObject: Set dataLo = lo(tableName)
    If dataLo Is Nothing Then
        LogDiagnostic "ERROR", "ValidateEnvironment", "Data table not found", "Table: " & tableName
        ValidateEnvironment = False
        Exit Function
    End If
    LogDiagnostic "INFO", "ValidateEnvironment", "Data table found", "Columns: " & dataLo.ListColumns.Count
    
    ' Check if data table has required columns
    Dim requiredHeaders As Variant
    requiredHeaders = Array("Tag ID", "Functional System Category", "Functional System", "Equipment Description")
    
    Dim i As Long, headerName As String, headerIdx As Long
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        headerName = CStr(requiredHeaders(i))
        headerIdx = HeaderIndexByText(dataLo, headerName)
        If headerIdx = 0 Then
            LogDiagnostic "ERROR", "ValidateEnvironment", "Required column missing", "Column: " & headerName
            ValidateEnvironment = False
            Exit Function
        End If
        LogDiagnostic "INFO", "ValidateEnvironment", "Required column found", "Column: " & headerName & " (Index: " & headerIdx & ")")
    Next i
    
    ' Check if required functions are available
    Dim testResult As Variant
    On Error Resume Next
    testResult = GetConfigValue("Test")
    If Err.Number <> 0 Then
        On Error GoTo ValidationError
        LogDiagnostic "ERROR", "ValidateEnvironment", "GetConfigValue function not available", "Error: " & Err.Description
        ValidateEnvironment = False
        Exit Function
    End If
    Err.Clear
    On Error GoTo ValidationError
    LogDiagnostic "INFO", "ValidateEnvironment", "GetConfigValue function validated", ""
    
    ' Check if output ranges are configured
    Dim resultsCell As String: resultsCell = GetConfigValue("ResultsStartCell")
    If Len(Trim$(resultsCell)) = 0 Then
        LogDiagnostic "WARN", "ValidateEnvironment", "ResultsStartCell not configured", "May cause output issues"
    Else
        LogDiagnostic "INFO", "ValidateEnvironment", "ResultsStartCell configured", "Value: " & resultsCell
    End If
    
    ' Validate named range function
    On Error Resume Next
    Dim testRange As Range: Set testRange = nr(resultsCell)
    If Err.Number <> 0 Then
        On Error GoTo ValidationError
        LogDiagnostic "WARN", "ValidateEnvironment", "Named range function issues", "Error: " & Err.Description
    Else
        LogDiagnostic "INFO", "ValidateEnvironment", "Named range function validated", ""
    End If
    Err.Clear
    On Error GoTo ValidationError
    
    LogDiagnostic "SUCCESS", "ValidateEnvironment", "Environment validation completed", "All critical components validated"
    ValidateEnvironment = True
    Exit Function
    
ValidationError:
    LogDiagnostic "ERROR", "ValidateEnvironment", "Validation failed with error", "Error: " & Err.Number & " - " & Err.Description
    ValidateEnvironment = False
End Function

Private Sub LogDiagnostic(ByVal severity As String, ByVal functionName As String, ByVal message As String, ByVal details As String)
    ' Comprehensive diagnostic logging to file with error handling
    On Error Resume Next
    
    Dim logEntry As String
    logEntry = Now & LOG_SEPARATOR & _
               severity & LOG_SEPARATOR & _
               "mod_SootblowerLocator." & functionName & LOG_SEPARATOR & _
               message & LOG_SEPARATOR & _
               details & LOG_SEPARATOR & _
               Application.UserName & LOG_SEPARATOR & _
               Environ("COMPUTERNAME")
    
    ' Try to write to file first
    Dim success As Boolean: success = WriteLogToFile(logEntry)
    
    ' Also try worksheet logging as backup
    If Not success Then WriteLogToWorksheet severity, functionName, message, details
    
    ' For critical errors, also try immediate window if available
    If severity = "ERROR" Then
        Debug.Print "[" & Now & "] ERROR in " & functionName & ": " & message & " | " & details
    End If
End Sub

Private Function WriteLogToFile(ByVal logEntry As String) As Boolean
    On Error Resume Next
    
    Dim logPath As String, fileName As String, baseLogPath As String
    baseLogPath = ThisWorkbook.Path & "\logs"
    logPath = baseLogPath & "\Diagnostic_Notes\"
    fileName = DIAG_LOG_FILE & "_" & Format(Now, "YYYYMMDD") & ".log"
    
    ' Ensure directories exist with more verbose error handling
    If Dir(baseLogPath, vbDirectory) = "" Then
        MkDir baseLogPath
        If Err.Number <> 0 Then
            Debug.Print "Failed to create logs directory: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
    End If
    
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
        If Err.Number <> 0 Then
            Debug.Print "Failed to create Diagnostic_Notes directory: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open logPath & fileName For Append As #fileNum
    If Err.Number = 0 Then
        Print #fileNum, logEntry
        Close #fileNum
        WriteLogToFile = True
    Else
        WriteLogToFile = False
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub WriteLogToWorksheet(ByVal severity As String, ByVal functionName As String, ByVal message As String, ByVal details As String)
    On Error Resume Next
    
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("SootblowerDiagnostics")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "SootblowerDiagnostics"
        ' Create headers
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Severity"
        ws.Cells(1, 3).Value = "Function"
        ws.Cells(1, 4).Value = "Message"
        ws.Cells(1, 5).Value = "Details"
        ws.Cells(1, 6).Value = "User"
        ws.Cells(1, 7).Value = "Computer"
        
        ' Format headers
        With ws.Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .AutoFilter
        End With
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = severity
    ws.Cells(lastRow, 3).Value = "mod_SootblowerLocator." & functionName
    ws.Cells(lastRow, 4).Value = message
    ws.Cells(lastRow, 5).Value = details
    ws.Cells(lastRow, 6).Value = Application.UserName
    ws.Cells(lastRow, 7).Value = Environ("COMPUTERNAME")
    
    ' Color coding by severity
    Select Case UCase(severity)
        Case "ERROR"
            ws.Cells(lastRow, 2).Interior.Color = RGB(255, 200, 200)
        Case "WARN"
            ws.Cells(lastRow, 2).Interior.Color = RGB(255, 255, 200)
        Case "SUCCESS"
            ws.Cells(lastRow, 2).Interior.Color = RGB(200, 255, 200)
    End Select
    
    ' Auto-size columns periodically
    If lastRow Mod 10 = 0 Then ws.Columns("A:G").AutoFit
End Sub

Private Function CLngSafe(ByVal value As Variant) As Long
    ' Safe conversion to Long with error handling
    On Error Resume Next
    CLngSafe = CLng(value)
    If Err.Number <> 0 Then
        CLngSafe = 0
        Err.Clear
    End If
End Function

Private Function SafeCellText(ByVal cellValue As Variant) As String
    ' Safe conversion of cell value to string with error handling
    On Error Resume Next
    If IsNull(cellValue) Then
        SafeCellText = ""
    ElseIf IsError(cellValue) Then
        SafeCellText = "#ERROR#"
    Else
        SafeCellText = CStr(cellValue)
    End If
    If Err.Number <> 0 Then
        SafeCellText = "#CONVERT_ERROR#"
        Err.Clear
    End If
End Function
