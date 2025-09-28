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

    Set wsConfig = ThisWorkbook.Sheets("ModeConfig")
    Set tblConfig = wsConfig.ListObjects("ModeConfigTable")

    Set colModeName = tblConfig.ListColumns("ModeName")
    Set colSearchFields = tblConfig.ListColumns("SearchFields")
    Set colFilterFields = tblConfig.ListColumns("FilterFields")
    Set colDescription = tblConfig.ListColumns("Description")
    ' Ensure CustomHandler column exists
    On Error Resume Next
    Set colCustomHandler = tblConfig.ListColumns("CustomHandler")
    On Error GoTo 0
    If colCustomHandler Is Nothing Then
        Set colCustomHandler = tblConfig.ListColumns.Add
        colCustomHandler.name = "CustomHandler"
    End If

    ' === Search for existing entry ===
    For Each r In tblConfig.ListRows
        If Trim(r.Range(colModeName.Index).value) = MODE_NAME Then
            Set foundRow = r
            Exit For
        End If
    Next r

    ' === If found, verify and update fields ===
    If Not foundRow Is Nothing Then
        With foundRow
            If Trim(.Range(colSearchFields.Index).value) <> SEARCH_FIELDS Then
                .Range(colSearchFields.Index).value = SEARCH_FIELDS
            End If
            If Trim(.Range(colFilterFields.Index).value) <> FILTER_FIELDS Then
                .Range(colFilterFields.Index).value = FILTER_FIELDS
            End If
            If Trim(.Range(colDescription.Index).value) <> DESCRIPTION Then
                .Range(colDescription.Index).value = DESCRIPTION
            End If
            If Trim(.Range(colCustomHandler.Index).value) <> "Init_SootblowerLocator" Then
                .Range(colCustomHandler.Index).value = "Init_SootblowerLocator"
            End If
        End With
    Else
        ' === If not found, add new row ===
        Set foundRow = tblConfig.ListRows.Add
        With foundRow
            .Range(colModeName.Index).value = MODE_NAME
            .Range(colSearchFields.Index).value = SEARCH_FIELDS
            .Range(colFilterFields.Index).value = FILTER_FIELDS
            .Range(colDescription.Index).value = DESCRIPTION
            .Range(colCustomHandler.Index).value = "Init_SootblowerLocator"
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
        If StrComp(CStr(r.Cells(1, 1).value), key, vbTextCompare) = 0 Then
            If Trim$(CStr(r.Cells(1, 2).value)) = "" Then r.Cells(1, 2).value = val
            Exit Sub
        End If
    Next r
    ' Add new
    Dim lr As ListRow: Set lr = loCfg.ListRows.Add
    lr.Range.Cells(1, 1).value = key
    lr.Range.Cells(1, 2).value = val
End Sub

