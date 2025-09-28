Attribute VB_Name = "PythonVBAConverter"
Option Explicit

'============================================================
' Python to VBA Code Converter
' Converts Python functions and logic to VBA equivalents
' File: PythonVBAConverter.bas
'============================================================

Public Type ConversionMapping
    PythonPattern As String
    VBAEquivalent As String
    Description As String
End Type

Public Sub ConvertPythonToVBA()
    '============================================================
    ' Main conversion routine - converts Python files to VBA
    '============================================================
    
    Application.ScreenUpdating = False
    
    ' Create or clear conversion worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Python_To_VBA")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Python_To_VBA"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    ws.Range("A1:F1").Value = Array("Python File", "Python Function", "VBA Equivalent", "Status", "Notes", "Priority")
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    Dim row As Long
    row = 2
    
    ' Scan Python files and suggest VBA conversions
    Call ScanPythonFilesForConversion(ws, row)
    
    ' Format worksheet
    ws.Columns.AutoFit
    ws.Range("A1:F" & row - 1).AutoFilter
    
    Application.ScreenUpdating = True
    
    MsgBox "Python to VBA conversion analysis complete!" & vbCrLf & _
           "Check the Python_To_VBA worksheet for conversion suggestions.", vbInformation
    
    ws.Activate
    
End Sub

Private Sub ScanPythonFilesForConversion(ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Scan Python files and identify functions for conversion
    '============================================================
    
    Dim pythonDir As String
    pythonDir = ThisWorkbook.Path & "\python\"
    
    If Dir(pythonDir, vbDirectory) = "" Then
        ws.Cells(row, 1).Value = "No python directory found"
        ws.Cells(row, 4).Value = "Setup Required"
        ws.Cells(row, 5).Value = "Create python/ subdirectory with Python files"
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = Dir(pythonDir & "*.py")
    
    Do While fileName <> ""
        Call AnalyzePythonFile(pythonDir & fileName, fileName, ws, row)
        fileName = Dir()
    Loop
    
End Sub

Private Sub AnalyzePythonFile(filePath As String, fileName As String, ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Analyze a single Python file for conversion opportunities
    '============================================================
    
    On Error Resume Next
    
    Dim fileContent As String
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    If Err.Number <> 0 Then
        ws.Cells(row, 1).Value = fileName
        ws.Cells(row, 4).Value = "Error Reading File"
        ws.Cells(row, 5).Value = "Could not read Python file"
        row = row + 1
        Err.Clear
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Extract function definitions
    Dim lines() As String
    lines = Split(fileContent, vbLf)
    
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        If Left(line, 4) = "def " And InStr(line, "(") > 0 Then
            Dim funcName As String
            funcName = ExtractPythonFunctionName(line)
            
            If funcName <> "" And Not funcName Like "__*" Then
                ws.Cells(row, 1).Value = fileName
                ws.Cells(row, 2).Value = funcName
                ws.Cells(row, 3).Value = GenerateVBAFunctionName(funcName)
                ws.Cells(row, 4).Value = "Ready for Conversion"
                ws.Cells(row, 5).Value = "Convert to VBA: " & line
                ws.Cells(row, 6).Value = DeterminePriority(funcName, line)
                row = row + 1
            End If
        End If
    Next i
    
End Sub

Private Function ExtractPythonFunctionName(line As String) As String
    '============================================================
    ' Extract function name from Python function definition
    '============================================================
    
    Dim startPos As Long, endPos As Long
    
    startPos = InStr(line, "def ") + 4
    endPos = InStr(startPos, line, "(")
    
    If endPos > startPos Then
        ExtractPythonFunctionName = Trim(Mid(line, startPos, endPos - startPos))
    Else
        ExtractPythonFunctionName = ""
    End If
    
End Function

Private Function GenerateVBAFunctionName(pythonName As String) As String
    '============================================================
    ' Convert Python function name to VBA naming convention
    '============================================================
    
    ' Convert snake_case to PascalCase
    Dim parts() As String
    parts = Split(pythonName, "_")
    
    Dim vbaName As String
    Dim i As Long
    
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            vbaName = vbaName & UCase(Left(parts(i), 1)) & LCase(Mid(parts(i), 2))
        End If
    Next i
    
    GenerateVBAFunctionName = vbaName
    
End Function

Private Function DeterminePriority(funcName As String, line As String) As String
    '============================================================
    ' Determine conversion priority based on function characteristics
    '============================================================
    
    ' High priority functions
    If InStr(LCase(funcName), "search") > 0 Or _
       InStr(LCase(funcName), "analyze") > 0 Or _
       InStr(LCase(funcName), "process") > 0 Then
        DeterminePriority = "HIGH"
    ' Medium priority functions
    ElseIf InStr(LCase(funcName), "get") > 0 Or _
           InStr(LCase(funcName), "set") > 0 Or _
           InStr(LCase(funcName), "load") > 0 Then
        DeterminePriority = "MEDIUM"
    ' Low priority functions
    Else
        DeterminePriority = "LOW"
    End If
    
End Function

Public Sub GenerateVBATemplate()
    '============================================================
    ' Generate VBA template code from Python functions
    '============================================================
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Python_To_VBA")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Please run ConvertPythonToVBA first!", vbExclamation
        Exit Sub
    End If
    
    ' Create VBA code generation worksheet
    Dim codeWs As Worksheet
    On Error Resume Next
    Set codeWs = ThisWorkbook.Worksheets("Generated_VBA")
    If codeWs Is Nothing Then
        Set codeWs = ThisWorkbook.Worksheets.Add
        codeWs.Name = "Generated_VBA"
    Else
        codeWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Generate VBA template code
    codeWs.Range("A1").Value = "Generated VBA Code Templates"
    codeWs.Range("A1").Font.Bold = True
    codeWs.Range("A1").Font.Size = 14
    
    Dim row As Long
    row = 3
    
    ' Process each function in the conversion worksheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value <> "" Then
            Dim pythonFunc As String, vbaFunc As String
            pythonFunc = ws.Cells(i, 2).Value
            vbaFunc = ws.Cells(i, 3).Value
            
            ' Generate VBA function template
            codeWs.Cells(row, 1).Value = "' Converted from Python function: " & pythonFunc
            row = row + 1
            codeWs.Cells(row, 1).Value = "Public Function " & vbaFunc & "() As Variant"
            row = row + 1
            codeWs.Cells(row, 1).Value = "    ' TODO: Implement " & pythonFunc & " functionality"
            row = row + 1
            codeWs.Cells(row, 1).Value = "    " & vbaFunc & " = ""Not implemented"""
            row = row + 1
            codeWs.Cells(row, 1).Value = "End Function"
            row = row + 2
        End If
    Next i
    
    codeWs.Columns.AutoFit
    codeWs.Activate
    
    MsgBox "VBA template code generated!" & vbCrLf & _
           "Check the Generated_VBA worksheet for template functions.", vbInformation
    
End Sub

Public Function GetConversionMappings() As ConversionMapping()
    '============================================================
    ' Get common Python to VBA conversion mappings
    '============================================================
    
    Dim mappings(20) As ConversionMapping
    
    mappings(0).PythonPattern = "len()"
    mappings(0).VBAEquivalent = "UBound() - LBound() + 1"
    mappings(0).Description = "Get array length"
    
    mappings(1).PythonPattern = "print()"
    mappings(1).VBAEquivalent = "Debug.Print"
    mappings(1).Description = "Print to console"
    
    mappings(2).PythonPattern = ".strip()"
    mappings(2).VBAEquivalent = "Trim$()"
    mappings(2).Description = "Remove whitespace"
    
    mappings(3).PythonPattern = ".lower()"
    mappings(3).VBAEquivalent = "LCase$()"
    mappings(3).Description = "Convert to lowercase"
    
    mappings(4).PythonPattern = ".upper()"
    mappings(4).VBAEquivalent = "UCase$()"
    mappings(4).Description = "Convert to uppercase"
    
    mappings(5).PythonPattern = ".split()"
    mappings(5).VBAEquivalent = "Split()"
    mappings(5).Description = "Split string into array"
    
    mappings(6).PythonPattern = ".join()"
    mappings(6).VBAEquivalent = "Join()"
    mappings(6).Description = "Join array into string"
    
    mappings(7).PythonPattern = ".append()"
    mappings(7).VBAEquivalent = "ReDim Preserve arr(UBound(arr) + 1): arr(UBound(arr)) = value"
    mappings(7).Description = "Append to array"
    
    mappings(8).PythonPattern = "range()"
    mappings(8).VBAEquivalent = "For i = 1 To n"
    mappings(8).Description = "Loop through range"
    
    mappings(9).PythonPattern = "enumerate()"
    mappings(9).VBAEquivalent = "For i = LBound() To UBound()"
    mappings(9).Description = "Loop with index"
    
    mappings(10).PythonPattern = "str()"
    mappings(10).VBAEquivalent = "CStr()"
    mappings(10).Description = "Convert to string"
    
    mappings(11).PythonPattern = "int()"
    mappings(11).VBAEquivalent = "CLng()"
    mappings(11).Description = "Convert to integer"
    
    mappings(12).PythonPattern = "float()"
    mappings(12).VBAEquivalent = "CDbl()"
    mappings(12).Description = "Convert to double"
    
    mappings(13).PythonPattern = "True"
    mappings(13).VBAEquivalent = "True"
    mappings(13).Description = "Boolean true"
    
    mappings(14).PythonPattern = "False"
    mappings(14).VBAEquivalent = "False"
    mappings(14).Description = "Boolean false"
    
    mappings(15).PythonPattern = "None"
    mappings(15).VBAEquivalent = "Nothing"
    mappings(15).Description = "Null value"
    
    mappings(16).PythonPattern = "def function_name():"
    mappings(16).VBAEquivalent = "Public Function FunctionName() As Variant"
    mappings(16).Description = "Function definition"
    
    mappings(17).PythonPattern = "if condition:"
    mappings(17).VBAEquivalent = "If condition Then"
    mappings(17).Description = "If statement"
    
    mappings(18).PythonPattern = "elif condition:"
    mappings(18).VBAEquivalent = "ElseIf condition Then"
    mappings(18).Description = "Else if statement"
    
    mappings(19).PythonPattern = "else:"
    mappings(19).VBAEquivalent = "Else"
    mappings(19).Description = "Else statement"
    
    mappings(20).PythonPattern = "for item in list:"
    mappings(20).VBAEquivalent = "For Each item In list"
    mappings(20).Description = "For each loop"
    
    GetConversionMappings = mappings
    
End Function

Public Sub ShowConversionReference()
    '============================================================
    ' Show a reference guide for Python to VBA conversions
    '============================================================
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Conversion_Reference")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Conversion_Reference"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    ws.Range("A1:C1").Value = Array("Python Pattern", "VBA Equivalent", "Description")
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Add conversion mappings
    Dim mappings() As ConversionMapping
    mappings = GetConversionMappings()
    
    Dim i As Long
    For i = 0 To UBound(mappings)
        ws.Cells(i + 2, 1).Value = mappings(i).PythonPattern
        ws.Cells(i + 2, 2).Value = mappings(i).VBAEquivalent
        ws.Cells(i + 2, 3).Value = mappings(i).Description
    Next i
    
    ws.Columns.AutoFit
    ws.Range("A1:C" & UBound(mappings) + 2).AutoFilter
    
    ws.Activate
    MsgBox "Conversion reference guide created!", vbInformation
    
End Sub