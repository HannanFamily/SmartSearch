Attribute VB_Name = "ModeConfigBootstrap"
Option Explicit

Public Sub Bootstrap_ModeConfig()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = GetOrCreateWorksheet("ModeConfig")
    Set lo = GetOrCreateListObject(ws, "ModeConfigTable", Array("ModeName", "SearchFields", "FilterFields", "Description", "CustomHandler"))

    Dim r As ListRow, found As ListRow
    Dim colMode As Long, colSearch As Long, colFilter As Long, colDesc As Long, colHandler As Long
    colMode = lo.ListColumns("ModeName").Index
    colSearch = lo.ListColumns("SearchFields").Index
    colFilter = lo.ListColumns("FilterFields").Index
    colDesc = lo.ListColumns("Description").Index
    colHandler = lo.ListColumns("CustomHandler").Index

    For Each r In lo.ListRows
        If Trim$(CStr(r.Range(colMode).Value)) = "Sootblower Location" Then Set found = r: Exit For
    Next r

    If found Is Nothing Then Set found = lo.ListRows.Add
    With found
        .Range(colMode).Value = "Sootblower Location"
        .Range(colSearch).Value = "Tag, Description"
        .Range(colFilter).Value = "Location, System"
        .Range(colDesc).Value = "Search by physical sootblower location"
        .Range(colHandler).Value = "Init_SootblowerLocator"
    End With
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        On Error Resume Next
        GetOrCreateWorksheet.name = sheetName
        On Error GoTo 0
    End If
End Function

Private Function GetOrCreateListObject(ByVal ws As Worksheet, ByVal tblName As String, ByVal headerNames As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    If lo Is Nothing Then
        Dim i As Long
        For i = LBound(headerNames) To UBound(headerNames)
            ws.Cells(1, i + 1).Value = CStr(headerNames(i))
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headerNames) + 1)), , xlYes)
        lo.name = tblName
    Else
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim lc As ListColumn
        For Each lc In lo.ListColumns: dict(lc.name) = True: Next lc
        Dim nm As Variant
        For Each nm In headerNames
            If Not dict.exists(CStr(nm)) Then lo.ListColumns.Add.name = CStr(nm)
        Next nm
    End If
    Set GetOrCreateListObject = lo
End Function
