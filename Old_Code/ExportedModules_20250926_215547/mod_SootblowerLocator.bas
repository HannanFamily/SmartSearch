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
Private Const FS_CAT_TARGET As String = "SOOT BLOWING"
Private Const TYPE_RETRACTS As String = "Retracts"
Private Const TYPE_WALL As String = "Wall"

Public Sub Init_SootblowerLocator()
    On Error GoTo fallback
    ' Ensure config and mode rows/columns exist per user request
    On Error Resume Next
    Ensure_ConfigKeys_Sootblower
    Ensure_ModeConfigEntry_SootblowerLocation
    On Error GoTo fallback

    ' Optionally auto-create parsed columns based on config toggle
    Dim autoParse As String: autoParse = UCase$(Trim$(GetConfigValue("SSB_AutoParseColumns")))
    If autoParse = "YES" Or autoParse = "TRUE" Or autoParse = "1" Then
        EnsureSSBParsedColumns
    End If

    ' Build/update association helpers once on init as well
    EnsureSSBAssocHelperColumns
    ' Try to create/show the modeless UserForm (via factory). If VBIDE access is blocked, fall back.
    If EnsureSootblowerForm() Then
        VBA.UserForms.Add("frmSootblowerLocator").Show vbModeless
        Exit Sub
    End If

fallback:
    ' Minimal fallback UI if form cannot be created
    Dim resp As VbMsgBoxResult, numTxt As String, grp As String
    numTxt = InputBox("Enter sootblower number (digits only, optional). Leave blank to list all.", "Sootblower Locator")
    resp = MsgBox("Limit by group? Yes = IK/EL Retracts, No = IR/WB Wall Blower, Cancel = both.", vbYesNoCancel + vbQuestion, "Group Filter")
    If resp = vbYes Then grp = TYPE_RETRACTS ElseIf resp = vbNo Then grp = TYPE_WALL Else grp = ""
    If Len(Trim$(numTxt)) = 0 Then
        SB_DisplayAll grp
    Else
        SB_ExecuteSearch numTxt, grp
    End If
End Sub

Public Sub SB_ExecuteSearch(ByVal numberText As String, ByVal groupSel As String)
    On Error GoTo EH
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim idxs As Variant
    idxs = FindSootblowerMatches(dataLo, numberText, groupSel)

    If IsEmpty(idxs) Then
        SB_Log "NoMatch", numberText, groupSel, 0, "No sootblower matched."
        If MsgBox("No match found. Try a different search?", vbQuestion + vbYesNo, "Sootblower Locator") = vbYes Then Exit Sub
        If MsgBox("Display all sootblowers?", vbQuestion + vbYesNo, "Sootblower Locator") = vbYes Then
            SB_DisplayAll ""
        End If
        Exit Sub
    End If

    ' If no group selected and more than one match with the same number across groups, prompt
    If Len(Trim$(groupSel)) = 0 Then
        Dim num As String: num = DigitsOnly(numberText)
        Dim cntRetr As Long, cntWall As Long
        CountByGroup dataLo, idxs, cntRetr, cntWall
        If cntRetr > 0 And cntWall > 0 Then
            SB_Log "AmbiguousNumber", num, groupSel, UBound(idxs), "Multiple groups match this number."
            Dim choose As VbMsgBoxResult
            choose = MsgBox("More than one sootblower uses this number. Choose a group: Yes = IK/EL (Retracts), No = IR/WB (Wall Blower).", vbYesNoCancel + vbQuestion, "Choose Group")
            If choose = vbYes Then groupSel = TYPE_RETRACTS
            If choose = vbNo Then groupSel = TYPE_WALL
            If choose = vbCancel Then Exit Sub
            ' Recompute with group filter
            idxs = FindSootblowerMatches(dataLo, numberText, groupSel)
            If IsEmpty(idxs) Then
                MsgBox "No results in selected group.", vbInformation
                Exit Sub
            End If
        End If
    End If

    OutputRowsToDashboard dataLo, idxs, False
    SB_Log "Output", numberText, groupSel, UBound(idxs), "Displayed matches"
    Exit Sub
EH:
    SB_Log "Error", numberText, groupSel, 0, CStr(Err.Number) & ": " & Err.DESCRIPTION
    MsgBox "Unexpected error in SB_ExecuteSearch: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
End Sub

Public Sub SB_DisplayAll(ByVal groupSel As String)
    On Error GoTo EH
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If
    Dim idxs As Variant
    idxs = FindSootblowerMatches(dataLo, "", groupSel) ' no number filter
    If IsEmpty(idxs) Then
        MsgBox "No sootblowers found.", vbInformation
        Exit Sub
    End If
    OutputRowsToDashboard dataLo, idxs, True
    SB_Log "ShowAll", "", groupSel, UBound(idxs), "Displayed all sootblowers"
    Exit Sub
EH:
    SB_Log "Error", "", groupSel, 0, CStr(Err.Number) & ": " & Err.DESCRIPTION
    MsgBox "Unexpected error in SB_DisplayAll: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
End Sub

' Core finder: returns 1-based row indexes relative to DataBodyRange
Private Function FindSootblowerMatches(ByVal dataLo As ListObject, ByVal numberText As String, ByVal groupSel As String) As Variant
    Dim fsCatIdx As Long, tagIdx As Long, fsIdx As Long
    fsCatIdx = GetColumnIndex("DataTable_FunctionalSystemCategory", dataLo): If fsCatIdx = 0 Then fsCatIdx = HeaderIndexByText(dataLo, "Functional System Category")
    fsIdx = GetColumnIndex("DataTable_FunctionalSystem", dataLo):        If fsIdx = 0 Then fsIdx = HeaderIndexByText(dataLo, "Functional System")
    tagIdx = GetColumnIndex("DataTable_TagID", dataLo):                    If tagIdx = 0 Then tagIdx = HeaderIndexByText(dataLo, "Tag ID")
    If fsCatIdx = 0 Or tagIdx = 0 Or fsIdx = 0 Then Exit Function

    Dim wantNum As String: wantNum = DigitsOnly(numberText)
    Dim n As Long: n = dataLo.DataBodyRange.Rows.count
    Dim r As Long, keep As Boolean
    Dim fsCat As String, tagText As String, sb As Boolean, num As String, tcode As String, fs As String
    Dim tmpIdx() As Long, cnt As Long
    ReDim tmpIdx(1 To n)

    For r = 1 To n
        fsCat = UCase$(Trim$(CStr(dataLo.DataBodyRange.Cells(r, fsCatIdx).Value)))
        If fsCat <> UCase$(FS_CAT_TARGET) Then GoTo NextRow

        tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
        ParseSSBTag tagText, sb, num, tcode
        If Not sb Then GoTo NextRow

        fs = UCase$(Trim$(CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value)))
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

    If cnt = 0 Then Exit Function
    ReDim Preserve tmpIdx(1 To cnt)
    FindSootblowerMatches = tmpIdx
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
    On Error GoTo EH
    Dim resultsStart As Range, statusRng As Range
    Set resultsStart = nr(GetConfigValue("ResultsStartCell"))
    Set statusRng = nr(GetConfigValue("StatusCell"))
    If resultsStart Is Nothing Then
        MsgBox "Named range 'ResultsStartCell' not found.", vbExclamation
        Exit Sub
    End If

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
            End If
        End If
    Next i
    If colCount = 0 Then
        MsgBox "No valid output columns found in ConfigTable.", vbExclamation
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

    ' Clear old results
    ClearOldResults resultsStart, colCount

    Dim maxRows As Long: maxRows = MaxOutputRows()
    Dim take As Long
    If maxRows > 0 Then
        If UBound(rowIdxs) < maxRows Then take = UBound(rowIdxs) Else take = maxRows
    Else
        take = UBound(rowIdxs)
    End If

    Dim outArr() As Variant, ri As Long, j As Long
    ReDim outArr(1 To take, 1 To colCount)
    For i = 1 To take
        ri = rowIdxs(i)
        For j = 1 To colCount
            outArr(i, j) = SafeCellText(dataLo.DataBodyRange.Cells(ri, outCols(j)).Value)
        Next j
    Next i

    If doSortAll And take > 1 Then
        Dim fsIdx As Long, descIdx As Long
        fsIdx = GetColumnIndex("DataTable_FunctionalSystem", dataLo): If fsIdx = 0 Then fsIdx = HeaderIndexByText(dataLo, "Functional System")
        descIdx = GetColumnIndex("DataTable_EquipDescription", dataLo)
        Dim fsOutPos As Long: fsOutPos = FindOutPos(outCols, colCount, fsIdx)
        Dim descOutPos As Long: descOutPos = FindOutPos(outCols, colCount, descIdx)
        ' Two-key sort: stable approach by sorting by second key first, then first key
        If descOutPos > 0 Then QuickSort2D_N outArr, 1, take, descOutPos
        If fsOutPos > 0 Then QuickSort2D_N outArr, 1, take, fsOutPos
    End If

    resultsStart.Offset(1, 0).Resize(take, colCount).Value = outArr
    WriteStatus statusRng, "Displayed " & take & " row(s).", "Sootblower Locator"
    Exit Sub
EH:
    MsgBox "Error in OutputRowsToDashboard: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
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

Public Function EnsureSootblowerForm() As Boolean
    ' Delegates to factory in SootblowerFormFactory.bas
    On Error Resume Next
    EnsureSootblowerForm = CreateSootblowerUserForm()
End Function

Public Sub SB_ShowAssociated(ByVal numberText As String, ByVal groupSel As String)
    On Error GoTo EH
    ' Resolve target sootblower set and display associated equipment using helper columns
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim assocMode As String: assocMode = GetConfigValue("SSB_Assoc_Mode")
    Dim assocMax As Long: assocMax = CLngSafe(GetConfigValue("SSB_Assoc_MaxRows"))
    Dim grp As String: grp = Trim$(groupSel)
    Dim num As String: num = DigitsOnly(numberText)

    ' Ensure helpers exist
    EnsureSSBParsedColumns
    EnsureSSBAssocHelperColumns

    ' If number empty, prompt to pick from displayed or ask for a number
    If Len(num) = 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("No number entered. Show all associated rows for all displayed sootblowers?", vbYesNoCancel + vbQuestion, "Associated Equipment")
        If resp = vbCancel Then Exit Sub
        If resp = vbNo Then Exit Sub ' user will refine first
    End If

    ' Find source sootblowers (to derive numbers and groups)
    Dim srcIdxs As Variant: srcIdxs = FindSootblowerMatches(dataLo, num, grp)
    If IsEmpty(srcIdxs) Then
        MsgBox "No sootblowers found to associate.", vbInformation
        Exit Sub
    End If

    Dim fsHeader As String: fsHeader = Nz(GetConfigValue("DataTable_FunctionalSystem"), "Functional System")
    Dim tagHeader As String: tagHeader = Nz(GetConfigValue("DataTable_TagID"), "Tag ID")
    Dim fsIdx As Long: fsIdx = HeaderIndexByText(dataLo, fsHeader)
    Dim tagIdx As Long: tagIdx = HeaderIndexByText(dataLo, tagHeader)

    Dim assocNumCol As String: assocNumCol = Nz(GetConfigValue("SSB_AssocHelper_NumberCol"), "Assoc SSB Number")
    Dim assocCatCol As String: assocCatCol = Nz(GetConfigValue("SSB_AssocHelper_CategoryCol"), "Assoc SSB Category")
    Dim assocNumIdx As Long: assocNumIdx = HeaderIndexByText(dataLo, assocNumCol)
    Dim assocCatIdx As Long: assocCatIdx = HeaderIndexByText(dataLo, assocCatCol)

    ' Build a set of target numbers and group preferences
    Dim targetNums As Object: Set targetNums = CreateObject("Scripting.Dictionary")
    targetNums.CompareMode = vbTextCompare
    Dim preferredGroup As String: preferredGroup = Trim$(grp)

    Dim i As Long, r As Long, isSB As Boolean, sbNum As String, tcode As String, fs As String
    For i = 1 To UBound(srcIdxs)
        r = srcIdxs(i)
        ParseSSBTag CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value), isSB, sbNum, tcode
        If isSB Then
            targetNums(sbNum) = True
            If Len(preferredGroup) = 0 Then
                fs = UCase$(CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value))
                If IsRetracts(fs, tcode) Then preferredGroup = "Retracts"
                If IsWall(fs, tcode) Then preferredGroup = IIf(Len(preferredGroup) = 0, "Wall", preferredGroup)
            End If
        End If
    Next i

    ' Scan helper columns to select associated rows: match on number, then choose closest group
    Dim n As Long: n = dataLo.DataBodyRange.Rows.count
    Dim pickIdx() As Long, cntPick As Long
    ReDim pickIdx(1 To n)
    Dim assocNum As String, assocCat As String
    For r = 1 To n
        assocNum = CStr(dataLo.DataBodyRange.Cells(r, assocNumIdx).Value)
        If Len(assocNum) > 0 Then
            If targetNums.exists(assocNum) Then
                ' group resolution: prefer preferredGroup if set; else accept any
                assocCat = CStr(dataLo.DataBodyRange.Cells(r, assocCatIdx).Value)
                If Len(preferredGroup) = 0 Or StrComp(assocCat, preferredGroup, vbTextCompare) = 0 Then
                    cntPick = cntPick + 1
                    pickIdx(cntPick) = r
                    If assocMax > 0 And cntPick >= assocMax Then Exit For
                End If
            End If
        End If
    Next r

    If cntPick = 0 Then
        SB_Log "Assoc", num, grp, 0, "No associated rows"
        MsgBox "No associated equipment found.", vbInformation
        Exit Sub
    End If

    ReDim Preserve pickIdx(1 To cntPick)
    OutputRowsToDashboard dataLo, pickIdx, True ' sorted view helpful here
    SB_Log "Assoc", num, grp, cntPick, "Displayed associated equipment"
    Exit Sub
EH:
    SB_Log "Error", numberText, groupSel, 0, CStr(Err.Number) & ": " & Err.DESCRIPTION
    MsgBox "Unexpected error in SB_ShowAssociated: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
End Sub

Public Sub EnsureSSBParsedColumns()
    ' Create/refresh three supporting parsed columns next to [Tag ID] for rows beginning with (SSB)
    On Error GoTo EH
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then
        MsgBox "Data table '" & DATA_TABLE_NAME & "' not found or empty.", vbExclamation
        Exit Sub
    End If

    Dim tagHeader As String: tagHeader = GetConfigValue("DataTable_TagID"): If Len(Trim$(tagHeader)) = 0 Then tagHeader = "Tag ID"
    Dim tagIdx As Long: tagIdx = HeaderIndexByText(dataLo, tagHeader)
    If tagIdx = 0 Then MsgBox "Tag ID column not found.", vbExclamation: Exit Sub

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

    If Not haveAll Then
        ' Insert three columns after Tag ID
        dataLo.ListColumns(tagIdx).Range.Offset(0, 1).Resize(1, 3).EntireColumn.Insert
        dataLo.HeaderRowRange.Cells(1, tagIdx + 1).Value = colPrefix
        dataLo.HeaderRowRange.Cells(1, tagIdx + 2).Value = colNumber
        dataLo.HeaderRowRange.Cells(1, tagIdx + 3).Value = colType
        ' Resize table to include new columns
        dataLo.Resize dataLo.Range.Resize(dataLo.Range.Rows.count, dataLo.Range.Columns.count + 3)
    End If

    ' Re-resolve indices after potential insert
    Dim prefixIdx As Long: prefixIdx = HeaderIndexByText(dataLo, colPrefix)
    Dim numberIdx As Long: numberIdx = HeaderIndexByText(dataLo, colNumber)
    Dim typeIdx As Long:   typeIdx = HeaderIndexByText(dataLo, colType)

    ' Fill parsed values for rows with (SSB) pattern only
    Dim n As Long: n = dataLo.DataBodyRange.Rows.count
    Dim r As Long, tagText As String, isSB As Boolean, num As String, tcode As String
    For r = 1 To n
        tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
        ParseSSBTag tagText, isSB, num, tcode
        If isSB Then
            dataLo.DataBodyRange.Cells(r, prefixIdx).Value = "(SSB)"
            dataLo.DataBodyRange.Cells(r, numberIdx).Value = num
            dataLo.DataBodyRange.Cells(r, typeIdx).Value = tcode
        End If
    Next r

    MsgBox "SSB parsed columns updated.", vbInformation
    Exit Sub
EH:
    MsgBox "Error in EnsureSSBParsedColumns: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
End Sub

Private Function Nz(ByVal s As String, ByVal fallback As String) As String
    If Len(Trim$(s)) = 0 Then Nz = fallback Else Nz = s
End Function

Public Sub EnsureSSBAssocHelperColumns()
    On Error GoTo EH
    Dim dataLo As ListObject: Set dataLo = lo(DATA_TABLE_NAME)
    If dataLo Is Nothing Or dataLo.DataBodyRange Is Nothing Then Exit Sub

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
    If catIdx = 0 Or fsIdx = 0 Or tagIdx = 0 Or descIdx = 0 Then Exit Sub

    ' Ensure helper columns exist (append to far right if missing)
    Dim needNames As Variant: needNames = Array(assocNumCol, assocCatCol)
    Dim i As Long
    For i = LBound(needNames) To UBound(needNames)
        If HeaderIndexByText(dataLo, CStr(needNames(i))) = 0 Then
            dataLo.ListColumns.Add
            dataLo.HeaderRowRange.Cells(1, dataLo.ListColumns.count).Value = CStr(needNames(i))
            dataLo.Resize dataLo.Range.Resize(dataLo.Range.Rows.count, dataLo.Range.Columns.count)
        End If
    Next i

    Dim assocNumIdx As Long: assocNumIdx = HeaderIndexByText(dataLo, assocNumCol)
    Dim assocCatIdx As Long: assocCatIdx = HeaderIndexByText(dataLo, assocCatCol)

    ' Pre-gather SSB numbers from (SSB) Tag rows to speed lookups
    Dim n As Long: n = dataLo.DataBodyRange.Rows.count
    Dim ssbNumbers As Object: Set ssbNumbers = CreateObject("Scripting.Dictionary")
    ssbNumbers.CompareMode = vbTextCompare
    Dim r As Long, tagText As String, isSB As Boolean, num As String, tcode As String
    For r = 1 To n
        If StrComp(CStr(dataLo.DataBodyRange.Cells(r, catIdx).Value), catVal, vbTextCompare) = 0 Then
            tagText = CStr(dataLo.DataBodyRange.Cells(r, tagIdx).Value)
            ParseSSBTag tagText, isSB, num, tcode
            If isSB Then If Not ssbNumbers.exists(num) Then ssbNumbers.Add num, True
        End If
    Next r

    ' Keyword groups from config
    Dim kwRetr As Variant, kwWall As Variant
    kwRetr = Split(Nz(GetConfigValue("SSB_AssocKeywords_Retracts"), "IK,EL,RETRACT"), ",")
    kwWall = Split(Nz(GetConfigValue("SSB_AssocKeywords_Wall"), "IR,WB,WALL,WATER"), ",")

    ' Walk all rows and populate helper columns for those in the target category
    Dim descText As String, udesc As String, foundNum As String, catGuess As String, k As Long
    For r = 1 To n
        If StrComp(CStr(dataLo.DataBodyRange.Cells(r, catIdx).Value), catVal, vbTextCompare) = 0 Then
            ' number detection: exact token match for any known SSB number
            descText = CStr(dataLo.DataBodyRange.Cells(r, descIdx).Value)
            foundNum = ""
            For Each num In ssbNumbers.Keys
                If HasNumberToken(descText, CStr(num)) Then
                    foundNum = CStr(num): Exit For
                End If
            Next num

            If Len(foundNum) > 0 Then
                dataLo.DataBodyRange.Cells(r, assocNumIdx).Value = foundNum
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
                    Dim fs As String: fs = UCase$(CStr(dataLo.DataBodyRange.Cells(r, fsIdx).Value))
                    If InStr(1, fs, "RETRACT", vbTextCompare) > 0 Then catGuess = "Retracts"
                    If InStr(1, fs, "WALL", vbTextCompare) > 0 Or InStr(1, fs, "WATER", vbTextCompare) > 0 Then catGuess = "Wall"
                End If
                If Len(catGuess) > 0 Then dataLo.DataBodyRange.Cells(r, assocCatIdx).Value = catGuess
            End If
        End If
    Next r
    Exit Sub
EH:
    MsgBox "Error in EnsureSSBAssocHelperColumns: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
End Sub

Private Function HasNumberToken(ByVal text As String, ByVal num As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False: rx.IgnoreCase = True
    rx.pattern = "(^|[^0-9])" & num & "($|[^0-9])"
    HasNumberToken = rx.Test(CStr(text))
End Function
