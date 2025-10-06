Attribute VB_Name = "mod_MetadataSnapshot_Enhanced"
'============================================================
' METADATA SNAPSHOT & COLUMN VALIDATION
'============================================================
' Purpose: Export workbook structure and verify DataTableColumn config
'============================================================
Option Explicit

' Main entry point
Public Sub ExportMetadataAndValidateColumns()
    Debug.Print "========================================="
    Debug.Print "WORKBOOK METADATA SNAPSHOT & COLUMN VALIDATION"
    Debug.Print "========================================="
    Call ExportSheetNames
    Call ExportNamedRanges
    Call ExportTableNames
    Call ValidateConfigDataTableColumns
    Debug.Print "========================================="
    Debug.Print "SNAPSHOT & VALIDATION COMPLETE"
    Debug.Print "========================================="
End Sub

' Export all worksheet names
Private Sub ExportSheetNames()
    Dim ws As Worksheet
    Debug.Print "-- Worksheets --"
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "WS: " & ws.Index & ": " & ws.Name & " Visible=" & IIf(ws.Visible = xlSheetVisible, "Visible", "Hidden")
    Next ws
End Sub

' Export all named ranges
Private Sub ExportNamedRanges()
    Dim nm As Name
    Debug.Print "-- Named Ranges (Workbook Scope Only) --"
    For Each nm In ThisWorkbook.Names
        Debug.Print "Name: " & nm.Name & " -> " & nm.RefersTo & " | Addr=" & GetRangeAddressSafe(nm)
    Next nm
End Sub

Private Function GetRangeAddressSafe(nm As Name) As String
    On Error Resume Next
    GetRangeAddressSafe = nm.RefersToRange.Address
    If Err.Number <> 0 Then GetRangeAddressSafe = "(invalid)"
    On Error GoTo 0
End Function

' Export all table names and columns
Private Sub ExportTableNames()
    Dim ws As Worksheet, lo As ListObject, i As Long
    Debug.Print "-- Tables --"
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            Debug.Print "Table: " & lo.Name & " Rows=" & lo.ListRows.Count & " Cols=" & lo.ListColumns.Count & " DataR1C1=" & lo.DataBodyRange.Address
            Dim colNames As String: colNames = "Columns: "
            For i = 1 To lo.ListColumns.Count
                colNames = colNames & lo.ListColumns(i).Name
                If i < lo.ListColumns.Count Then colNames = colNames & ", "
            Next i
            Debug.Print colNames
        Next lo
    Next ws
End Sub

' Validate all config DataTableColumn entries against DataTable
Private Sub ValidateConfigDataTableColumns()
    Debug.Print "-- DataTableColumn Validation --"
    Dim configWs As Worksheet, configLo As ListObject
    Set configWs = Nothing: Set configLo = Nothing
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets("ConfigSheet")
    Set configLo = configWs.ListObjects("ConfigTable")
    On Error GoTo 0
    If configLo Is Nothing Then
        Debug.Print "ConfigTable not found. Skipping column validation."
        Exit Sub
    End If
    
    ' Get DataTable reference
    Dim dataTableName As String
    dataTableName = GetConfigValue("DATA_TABLE_NAME")
    If Len(dataTableName) = 0 Then
        Debug.Print "DATA_TABLE_NAME not set in config."
        Exit Sub
    End If
    Dim dataLo As ListObject
    Set dataLo = FindTableByName(dataTableName)
    If dataLo Is Nothing Then
        Debug.Print "DataTable '" & dataTableName & "' not found."
        Exit Sub
    End If
    
    ' Build set of DataTable columns
    Dim colSet As Object: Set colSet = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To dataLo.ListColumns.Count
        colSet(dataLo.ListColumns(i).Name) = True
    Next i
    
    ' Check all config entries with Type = DataTableColumn
    Dim r As Range, configColName As String, found As Boolean
    For Each r In configLo.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 3).Value), "DataTableColumn", vbTextCompare) = 0 Then
            configColName = CStr(r.Cells(1, 2).Value)
            found = colSet.Exists(configColName)
            If found Then
                Debug.Print "✓ Config DataTableColumn found: " & configColName
            Else
                Debug.Print "❌ Config DataTableColumn missing: " & configColName
            End If
        End If
    Next r
End Sub

' Helper: Find table by name
Private Function FindTableByName(tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindTableByName = Nothing
End Function

' Helper: Get config value
Private Function GetConfigValue(configKey As String) As String
    Dim ws As Worksheet, lo As ListObject, r As Range
    Set ws = Nothing: Set lo = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set lo = ws.ListObjects("ConfigTable")
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    For Each r In lo.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), configKey, vbTextCompare) = 0 Then
            GetConfigValue = CStr(r.Cells(1, 2).Value)
            Exit Function
        End If
    Next r
    GetConfigValue = ""
End Function
