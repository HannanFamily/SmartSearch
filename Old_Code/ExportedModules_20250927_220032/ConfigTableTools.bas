Attribute VB_Name = "ConfigTableTools"
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

    Set wsConfig = ThisWorkbook.Sheets("ModeConfig")
    Set tblConfig = wsConfig.ListObjects("ModeConfigTable")

    Set colModeName = tblConfig.ListColumns("ModeName")
    Set colSearchFields = tblConfig.ListColumns("SearchFields")
    Set colFilterFields = tblConfig.ListColumns("FilterFields")
    Set colDescription = tblConfig.ListColumns("Description")

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
        End With
    Else
        ' === If not found, add new row ===
        Set foundRow = tblConfig.ListRows.Add
        With foundRow
            .Range(colModeName.Index).value = MODE_NAME
            .Range(colSearchFields.Index).value = SEARCH_FIELDS
            .Range(colFilterFields.Index).value = FILTER_FIELDS
            .Range(colDescription.Index).value = DESCRIPTION
        End With
    End If
End Sub

