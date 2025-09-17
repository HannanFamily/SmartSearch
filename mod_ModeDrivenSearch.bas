Attribute VB_Name = "mod_ModeDrivenSearch"
'============================================================
' ModeDrivenSearch.bas
' Core logic for mode-driven search and output, using ModeConfig table
'============================================================
Option Explicit

' Returns the active mode name from the dashboard selector
Public Function GetActiveModeName() As String
    On Error Resume Next
    GetActiveModeName = Range("ModeSelector").Value
End Function

' Returns the ModeConfig row for the active mode as a dictionary
Public Function GetActiveModeConfig() As Object
    Dim lo As ListObject, r As ListRow, dict As Object
    Set lo = ThisWorkbook.Worksheets("ModeConfig").ListObjects("ModeConfigTable")
    Set dict = CreateObject("Scripting.Dictionary")
    Dim modeName As String: modeName = GetActiveModeName()
    For Each r In lo.ListRows
        If StrComp(CStr(r.Range.Cells(1, 1).Value), modeName, vbTextCompare) = 0 Then
            Dim i As Long
            For i = 1 To lo.ListColumns.Count
                dict(lo.HeaderRowRange.Cells(1, i).Value) = r.Range.Cells(1, i).Value
            Next i
            Exit For
        End If
    Next r
    Set GetActiveModeConfig = dict
End Function

' Applies the mode's filter to the DataTable and returns matching row indices
Public Function GetModeFilteredIndexes(dataLo As ListObject, modeDict As Object) As Variant
    Dim idxs() As Long, i As Long, n As Long
    Dim f As String: f = modeDict("FilterFormula")
    Dim evalResult As Boolean
    n = 0
    ReDim idxs(1 To dataLo.DataBodyRange.Rows.Count)
    For i = 1 To dataLo.DataBodyRange.Rows.Count
        evalResult = EvaluateModeFormula(f, dataLo, i)
        If evalResult Then
            n = n + 1
            idxs(n) = i
        End If
    Next i
    If n = 0 Then
        GetModeFilteredIndexes = Array()
    Else
        ReDim Preserve idxs(1 To n)
        GetModeFilteredIndexes = idxs
    End If
End Function

' Evaluates the filter formula for a given row (simple parser for [@ColName] tokens)
Public Function EvaluateModeFormula(formula As String, dataLo As ListObject, rowIdx As Long) As Boolean
    Dim rx As Object, m As Object, colName As String, val As String, fEval As String
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "\[@([A-Za-z0-9_]+)\]"
    fEval = formula
    Do While rx.Test(fEval)
        Set m = rx.Execute(fEval)(0)
        colName = m.SubMatches(0)
        val = CStr(dataLo.DataBodyRange.Cells(rowIdx, HeaderIndexByText(dataLo, colName)).Value)
        fEval = Replace(fEval, m.Value, Chr(34) & val & Chr(34))
    Loop
    EvaluateModeFormula = Application.Evaluate(fEval)
End Function

' Outputs results for the active mode (columns, format)
Public Sub OutputModeResults()
    Dim dataLo As ListObject, modeDict As Object, idxs As Variant
    Dim outCols As Variant, i As Long, j As Long, resultsStart As Range
    Set dataLo = lo(DATA_TABLE_NAME)
    Set modeDict = GetActiveModeConfig()
    idxs = GetModeFilteredIndexes(dataLo, modeDict)
    outCols = Split(modeDict("OutputColumns"), ",")
    Set resultsStart = NR(GetConfigValue("ResultsStartCell"))
    ' Clear old results
    resultsStart.Offset(1, 0).Resize(1000, UBound(outCols) + 1).ClearContents
    ' Write headers
    For j = 0 To UBound(outCols)
        resultsStart.Offset(0, j).Value = outCols(j)
    Next j
    ' Write data
    For i = 1 To UBound(idxs)
        For j = 0 To UBound(outCols)
            resultsStart.Offset(i, j).Value = dataLo.DataBodyRange.Cells(idxs(i), HeaderIndexByText(dataLo, Trim$(outCols(j)))).Value
        Next j
    Next i
End Sub

'============================================================
' Notes:
' - Call OutputModeResults instead of the default output routine when a mode is active.
' - Extend EvaluateModeFormula for more complex logic as needed.
'============================================================

