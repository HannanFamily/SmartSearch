Attribute VB_Name = "temp_mod_ConfigTableTools"
' =====================================================================================
' Module: mod_ConfigTableTools
' Purpose: Maintain and verify ModeConfigTable entries for ModeDrivenSearch
' =====================================================================================

Public Sub Ensure_ModeConfigEntry_SootblowerLocation()
    Const MODE_NAME As String = "Sootblower Location"
    Const SEARCH_FIELDS As String = "Tag, Description"
    Const FILTER_FIELDS As String = "Location, System"
    Const DESCRIPTION As String = "Search by physical sootblower location"

    Dim wsConfig As Worksheet
    Dim tblConfig As ListObject
    Dim r As ListRow
    Dim foundRow As ListRow
    Dim colModeName As ListColumn, colSearchFields As ListColumn
    Dim colFilterFields As ListColumn, colDescription As ListColumn
    Dim colCustomHandler As ListColumn

    ' Ensure sheet and table exist and have required columns
    Set wsConfig = GetOrCreateWorksheet("ModeConfig")
    Set tblConfig = GetOrCreateListObject(wsConfig, "ModeConfigTable", Array("ModeName", "SearchFields", "FilterFields", "Description", "CustomHandler"))

    Set colModeName = GetOrAddListColumn(tblConfig, "ModeName")
    Set colSearchFields = GetOrAddListColumn(tblConfig, "SearchFields")
    Set colFilterFields = GetOrAddListColumn(tblConfig, "FilterFields")
    Set colDescription = GetOrAddListColumn(tblConfig, "Description")
    Set colCustomHandler = GetOrAddListColumn(tblConfig, "CustomHandler")

    ' === Search for existing entry ===
    For Each r In tblConfig.ListRows
        If Trim$(CStr(r.Range(colModeName.Index).Value)) = MODE_NAME Then
            Set foundRow = r
            Exit For
        End If
    Next r

    ' === If found, verify and update fields ===
    If Not foundRow Is Nothing Then
        With foundRow
            If Trim$(CStr(.Range(colSearchFields.Index).Value)) <> SEARCH_FIELDS Then
                .Range(colSearchFields.Index).Value = SEARCH_FIELDS
            End If
            If Trim$(CStr(.Range(colFilterFields.Index).Value)) <> FILTER_FIELDS Then
                .Range(colFilterFields.Index).Value = FILTER_FIELDS
            End If
            If Trim$(CStr(.Range(colDescription.Index).Value)) <> DESCRIPTION Then
                .Range(colDescription.Index).Value = DESCRIPTION
            End If
            If Trim$(CStr(.Range(colCustomHandler.Index).Value)) <> "Init_SootblowerLocator" Then
                .Range(colCustomHandler.Index).Value = "Init_SootblowerLocator"
            End If
        End With
    Else
        ' === If not found, add new row ===
        Set foundRow = tblConfig.ListRows.Add
        With foundRow
            .Range(colModeName.Index).Value = MODE_NAME
            .Range(colSearchFields.Index).Value = SEARCH_FIELDS
            .Range(colFilterFields.Index).Value = FILTER_FIELDS
            .Range(colDescription.Index).Value = DESCRIPTION
            .Range(colCustomHandler.Index).Value = "Init_SootblowerLocator"
        End With
    End If
End Sub

Public Sub Ensure_ConfigKeys_Sootblower()
    ' Ensure required keys exist in ConfigTable for Sootblower Locator
    Dim ws As Worksheet, loCfg As ListObject
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set loCfg = ws.ListObjects("ConfigTable")
    If loCfg Is Nothing Then Exit Sub

    ' Helper to upsert key/value
    Dim pairs As Variant
    pairs = Array( _
        Array("DataTable_FunctionalSystemCategory", "Functional System Category"), _
        Array("DataTable_FunctionalSystem", "Functional System"), _
        Array("DataTable_TagID", "Tag ID"), _
        Array("DataTable_EquipDescription", "Equipment Description"), _
        Array("SSB_FunctionalSystemCategoryValue", "SOOT BLOWING"), _
        Array("SSB_TagPrefix", "(SSB)"), _
        Array("SSB_TagRegex", "^\(SSB\)\s*(\d{1,3})\s+([A-Za-z0-9_\-]+)"), _
        Array("SSB_FS_Retracts", "RETRACTS"), _
        Array("SSB_FS_WallBlower", "WALL BLOWER"), _
        Array("SSB_Group_Retracts_Types", "SBEL,SBIK"), _
        Array("SSB_Group_Wall_Types", "SBIR,SBWB"), _
        Array("SSB_ParsedPrefixCol", "SSB Prefix"), _
        Array("SSB_ParsedNumberCol", "SSB Number"), _
        Array("SSB_ParsedTypeCol", "SSB Type"), _
        Array("SSB_AutoParseColumns", "Yes"), _
        Array("SSB_Assoc_ButtonLabel", "Show all associated equipment"), _
        Array("SSB_Assoc_Mode", "InlineBelow"), _
        Array("SSB_Assoc_MaxRows", "500"), _
        Array("SSB_Assoc_FilterCategory", "SOOT BLOWING"), _
        Array("SSB_AssocHelper_NumberCol", "Assoc SSB Number"), _
        Array("SSB_AssocHelper_CategoryCol", "Assoc SSB Category"), _
        Array("SSB_AssocKeywords_Retracts", "IK,EL,RETRACT"), _
        Array("SSB_AssocKeywords_Wall", "IR,WB,WALL,WATER") _
    )

    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        UpsertConfig loCfg, CStr(pairs(i)(0)), CStr(pairs(i)(1))
    Next i
End Sub

Private Sub UpsertConfig(ByVal loCfg As ListObject, ByVal key As String, ByVal val As String)
    Dim r As Range
    For Each r In loCfg.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            If Trim$(CStr(r.Cells(1, 2).Value)) = "" Then r.Cells(1, 2).Value = val
            Exit Sub
        End If
    Next r
    ' Add new
    Dim lr As ListRow: Set lr = loCfg.ListRows.Add
    lr.Range.Cells(1, 1).Value = key
    lr.Range.Cells(1, 2).Value = val
End Sub

' =====================
' Helpers: Sheet/Table/Column
' =====================
Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        On Error Resume Next
        GetOrCreateWorksheet.Name = sheetName
        On Error GoTo 0
    End If
End Function

Private Function GetOrCreateListObject(ByVal ws As Worksheet, ByVal tblName As String, ByVal headerNames As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    If lo Is Nothing Then
        ' Create headers starting at A1
        Dim i As Long
        For i = LBound(headerNames) To UBound(headerNames)
            ws.Cells(1, i + 1).Value = CStr(headerNames(i))
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headerNames) + 1)), , xlYes)
        lo.Name = tblName
    Else
        ' Ensure any missing columns are appended
        Dim nameMap As Object
        Set nameMap = CreateObject("Scripting.Dictionary")
        Dim lc As ListColumn
        For Each lc In lo.ListColumns
            nameMap(lc.Name) = True
        Next lc
        Dim nm As Variant
        For Each nm In headerNames
            If Not nameMap.Exists(CStr(nm)) Then
                With lo.ListColumns.Add
                    .Name = CStr(nm)
                End With
            End If
        Next nm
    End If
    Set GetOrCreateListObject = lo
End Function

Private Function GetOrAddListColumn(ByVal lo As ListObject, ByVal colName As String) As ListColumn
    On Error Resume Next
    Set GetOrAddListColumn = lo.ListColumns(colName)
    On Error GoTo 0
    If GetOrAddListColumn Is Nothing Then
        Set GetOrAddListColumn = lo.ListColumns.Add
        GetOrAddListColumn.Name = colName
    End If
End Function

