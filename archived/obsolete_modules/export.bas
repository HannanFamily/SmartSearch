' filepath: /run/user/1000/gvfs/smb-share:server=tower.local,share=systemfiles/allshares/nvmeshare/Dashboard Project/ExportWorkbookMetadata.bas
' Exports worksheet names, named ranges, table names/headers, and key worksheet contents to a text file

Sub ExportWorkbookMetadata()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name
    Dim lo As ListObject
    Dim fso As Object, ts As Object
    Dim outPath As String
    Dim r As Range, c As Range
    Dim i As Long, j As Long

    Set wb = ThisWorkbook
    outPath = wb.Path & Application.PathSeparator & "Workbook_Metadata.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(outPath, True)

    ts.WriteLine "=== Worksheet Names ==="
    For Each ws In wb.Worksheets
        ts.WriteLine ws.Name
    Next ws

    ts.WriteLine vbCrLf & "=== Named Ranges ==="
    For Each nm In wb.Names
        ts.WriteLine nm.Name & " -> " & nm.RefersTo
        On Error Resume Next
        If Not nm.RefersToRange Is Nothing Then
            ts.WriteLine "    Value: " & nm.RefersToRange.Text
        End If
        On Error GoTo 0
    Next nm

    ts.WriteLine vbCrLf & "=== Tables and Headers ==="
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            ts.WriteLine "Table: " & lo.Name & " (Sheet: " & ws.Name & ")"
            ts.Write "    Headers: "
            For i = 1 To lo.HeaderRowRange.Columns.Count
                ts.Write lo.HeaderRowRange.Cells(1, i).Value
                If i < lo.HeaderRowRange.Columns.Count Then ts.Write ", "
            Next i
            ts.WriteLine ""
        Next lo
    Next ws


    ' Export contents of every worksheet
    For Each ws In wb.Worksheets
        ts.WriteLine vbCrLf & "=== Worksheet: " & ws.Name & " ==="
        On Error Resume Next
        If Not ws.UsedRange Is Nothing Then
            For Each r In ws.UsedRange.Rows
                For Each c In r.Cells
                    ts.Write c.Text & vbTab
                Next c
                ts.WriteLine ""
            Next r
        End If
        On Error GoTo 0
    Next ws

    ts.Close
    MsgBox "Workbook metadata exported to: " & outPath, vbInformation
End Sub