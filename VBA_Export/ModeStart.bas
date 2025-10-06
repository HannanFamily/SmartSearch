Attribute VB_Name = "ModeStart"
' =====================================================================================
' Handler: Search_SootblowerLocation
' Purpose: Filters equipment records by search term and sootblower location
' Triggered by: ModeDrivenSearch when SearchMode = "Sootblower Location"
' =====================================================================================

Public Sub Search_SootblowerLocation()
    Dim wsData As Worksheet, wsResults As Worksheet
    Dim tblData As ListObject
    Dim searchTerm As String, locationFilter As String
    Dim r As ListRow, matchFound As Boolean
    Dim resultRow As Long

    ' === Setup ===
    Set wsData = ThisWorkbook.Sheets("EquipmentData")
    Set wsResults = ThisWorkbook.Sheets("SearchResults")
    Set tblData = wsData.ListObjects("tbl_Equipment")

    searchTerm = Trim(wsResults.Range("SearchInput").Value)
    locationFilter = Trim(wsResults.Range("LocationFilter").Value)

    wsResults.Range("ResultsTable").ClearContents
    resultRow = wsResults.Range("ResultsTable").Row

    ' === Loop through data table ===
    For Each r In tblData.ListRows
        matchFound = False

        ' Match search term in Tag or Description
        If InStr(1, r.Range(tblData.ListColumns("Tag").Index).Value, searchTerm, vbTextCompare) > 0 _
        Or InStr(1, r.Range(tblData.ListColumns("Description").Index).Value, searchTerm, vbTextCompare) > 0 Then
            matchFound = True
        End If

        ' Match location filter
        If matchFound Then
            If locationFilter = "" Or _
               StrComp(r.Range(tblData.ListColumns("Location").Index).Value, locationFilter, vbTextCompare) = 0 Then
                ' Copy matching row to results
                r.Range.Copy Destination:=wsResults.Cells(resultRow, 1)
                resultRow = resultRow + 1
            End If
        End If
    Next r

    ' === Finalize ===
    If resultRow = wsResults.Range("ResultsTable").Row Then
        MsgBox "No matching records found for Sootblower Location mode.", vbInformation
    End If
End Sub


