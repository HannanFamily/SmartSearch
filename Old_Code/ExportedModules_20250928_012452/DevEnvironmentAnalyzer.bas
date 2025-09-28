Attribute VB_Name = "DevEnvironmentAnalyzer"
'============================================================
' Development Environment Analyzer
' Analyzes Python and VBA files to track synchronization status
'============================================================
Option Explicit

Public Type FunctionInfo
    name As String
    Signature As String
    fileName As String
    fileType As String
    LastModified As Date
    Status As String
End Type

Public Type ProjectAnalysis
    TotalFunctions As Long
    PythonOnlyCount As Long
    VBAOnlyCount As Long
    SynchronizedCount As Long
    ConflictCount As Long
End Type

' Main analysis routine
Public Sub AnalyzeDevEnvironment()
    Application.ScreenUpdating = False
    
    ' Clear existing analysis
    Call ClearAnalysisSheets
    
    ' Create analysis worksheets if they don't exist
    Call CreateAnalysisWorksheets
    
    ' Scan all files
    Dim pythonFunctions() As FunctionInfo
    Dim vbaFunctions() As FunctionInfo
    
    pythonFunctions = ScanPythonFiles()
    vbaFunctions = ScanVBAFiles()
    
    ' Populate analysis worksheets
    Call PopulateFunctionOverview(pythonFunctions, vbaFunctions)
    Call CreateSyncStatusDashboard(pythonFunctions, vbaFunctions)
    Call CreateConversionTracker(pythonFunctions, vbaFunctions)
    
    ' Format worksheets
    Call FormatAnalysisWorksheets
    
    Application.ScreenUpdating = True
    
    MsgBox "Development environment analysis complete!" & vbCrLf & _
           "Check the new worksheets for detailed analysis.", vbInformation
End Sub

' Scan Python files for function definitions
Private Function ScanPythonFiles() As FunctionInfo()
    Dim functions() As FunctionInfo
    Dim functionCount As Long
    Dim pythonDir As String
    Dim fileName As String
    
    pythonDir = ThisWorkbook.path & "\python\"
    
    If Dir(pythonDir, vbDirectory) = "" Then
        ' Return empty array if python directory doesn't exist
        ReDim functions(0)
        ScanPythonFiles = functions
        Exit Function
    End If
    
    fileName = Dir(pythonDir & "*.py")
    
    Do While fileName <> ""
        Dim filePath As String
        filePath = pythonDir & fileName
        
        ' Read file and extract functions
        Dim fileContent As String
        fileContent = ReadTextFile(filePath)
        
        If fileContent <> "" Then
            Call ExtractPythonFunctions(fileContent, fileName, functions, functionCount)
        End If
        
        fileName = Dir()
    Loop
    
    ' Resize array to actual count
    If functionCount > 0 Then
        ReDim Preserve functions(1 To functionCount)
    Else
        ReDim functions(0)
    End If
    
    ScanPythonFiles = functions
End Function

' Scan VBA files for function definitions
Private Function ScanVBAFiles() As FunctionInfo()
    Dim functions() As FunctionInfo
    Dim functionCount As Long
    Dim projectDir As String
    Dim fileName As String
    
    projectDir = ThisWorkbook.path & "\"
    
    ' Scan .bas files
    fileName = Dir(projectDir & "*.bas")
    Do While fileName <> ""
        Dim filePath As String
        filePath = projectDir & fileName
        
        Dim fileContent As String
        fileContent = ReadTextFile(filePath)
        
        If fileContent <> "" Then
            Call ExtractVBAFunctions(fileContent, fileName, functions, functionCount)
        End If
        
        fileName = Dir()
    Loop
    
    ' Scan .cls files
    fileName = Dir(projectDir & "*.cls")
    Do While fileName <> ""
        filePath = projectDir & fileName
        fileContent = ReadTextFile(filePath)
        
        If fileContent <> "" Then
            Call ExtractVBAFunctions(fileContent, fileName, functions, functionCount)
        End If
        
        fileName = Dir()
    Loop
    
    ' Resize array to actual count
    If functionCount > 0 Then
        ReDim Preserve functions(1 To functionCount)
    Else
        ReDim functions(0)
    End If
    
    ScanVBAFiles = functions
End Function

' Extract Python function definitions
Private Sub ExtractPythonFunctions(fileContent As String, fileName As String, _
                                 ByRef functions() As FunctionInfo, ByRef functionCount As Long)
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim funcName As String
    
    lines = Split(fileContent, vbLf)
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Look for function definitions: def function_name(
        If Left(line, 4) = "def " And InStr(line, "(") > 0 Then
            ' Extract function name
            funcName = ExtractPythonFunctionName(line)
            
            If funcName <> "" And Not funcName Like "__*" Then ' Skip private/magic methods
                functionCount = functionCount + 1
                ReDim Preserve functions(1 To functionCount)
                
                With functions(functionCount)
                    .name = funcName
                    .Signature = line
                    .fileName = fileName
                    .fileType = "Python"
                    .LastModified = FileDateTime(ThisWorkbook.path & "\python\" & fileName)
                    .Status = "Needs Analysis"
                End With
            End If
        End If
    Next i
End Sub

' Extract VBA function definitions
Private Sub ExtractVBAFunctions(fileContent As String, fileName As String, _
                               ByRef functions() As FunctionInfo, ByRef functionCount As Long)
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim funcName As String
    
    lines = Split(fileContent, vbLf)
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Look for Sub or Function definitions
        If (Left(UCase(line), 4) = "SUB " Or Left(UCase(line), 9) = "FUNCTION " Or _
           Left(UCase(line), 11) = "PUBLIC SUB " Or Left(UCase(line), 16) = "PUBLIC FUNCTION " Or _
           Left(UCase(line), 12) = "PRIVATE SUB " Or Left(UCase(line), 17) = "PRIVATE FUNCTION ") _
           And InStr(line, "(") > 0 Then
            
            ' Extract function name
            funcName = ExtractVBAFunctionName(line)
            
            If funcName <> "" Then
                functionCount = functionCount + 1
                ReDim Preserve functions(1 To functionCount)
                
                With functions(functionCount)
                    .name = funcName
                    .Signature = line
                    .fileName = fileName
                    .fileType = "VBA"
                    .LastModified = FileDateTime(ThisWorkbook.path & "\" & fileName)
                    .Status = "Needs Analysis"
                End With
            End If
        End If
    Next i
End Sub

' Helper function to extract Python function name
Private Function ExtractPythonFunctionName(line As String) As String
    Dim startPos As Long, endPos As Long
    
    startPos = InStr(line, "def ") + 4
    endPos = InStr(startPos, line, "(")
    
    If endPos > startPos Then
        ExtractPythonFunctionName = Trim(Mid(line, startPos, endPos - startPos))
    Else
        ExtractPythonFunctionName = ""
    End If
End Function

' Helper function to extract VBA function name
Private Function ExtractVBAFunctionName(line As String) As String
    Dim words() As String
    Dim i As Long
    Dim funcIndex As Long
    
    words = Split(line, " ")
    
    ' Find the word before the opening parenthesis
    For i = 0 To UBound(words)
        If InStr(words(i), "(") > 0 Then
            ExtractVBAFunctionName = Replace(words(i), "(", "")
            Exit For
        ElseIf UCase(words(i)) = "SUB" Or UCase(words(i)) = "FUNCTION" Then
            If i < UBound(words) Then
                ExtractVBAFunctionName = words(i + 1)
                If InStr(ExtractVBAFunctionName, "(") > 0 Then
                    ExtractVBAFunctionName = Left(ExtractVBAFunctionName, InStr(ExtractVBAFunctionName, "(") - 1)
                End If
                Exit For
            End If
        End If
    Next i
End Function

' Create analysis worksheets
Private Sub CreateAnalysisWorksheets()
    Dim ws As Worksheet
    
    ' Function Overview worksheet
    If Not WorksheetExists("Function_Overview") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "Function_Overview"
    End If
    
    ' Sync Status Dashboard
    If Not WorksheetExists("Sync_Dashboard") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "Sync_Dashboard"
    End If
    
    ' Conversion Tracker
    If Not WorksheetExists("Conversion_Tracker") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "Conversion_Tracker"
    End If
End Sub

' Check if worksheet exists
Private Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    WorksheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function

' Clear existing analysis data
Private Sub ClearAnalysisSheets()
    Dim wsNames As Variant
    Dim i As Long
    
    wsNames = Array("Function_Overview", "Sync_Dashboard", "Conversion_Tracker")
    
    For i = 0 To UBound(wsNames)
        If WorksheetExists(CStr(wsNames(i))) Then
            ThisWorkbook.Worksheets(CStr(wsNames(i))).Cells.Clear
        End If
    Next i
End Sub

' Populate Function Overview worksheet
Private Sub PopulateFunctionOverview(pythonFunctions() As FunctionInfo, vbaFunctions() As FunctionInfo)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Function_Overview")
    
    ' Create headers
    With ws.Range("A1:H1")
        .value = Array("Function Name", "Status", "Python File", "VBA File", "Python Signature", "VBA Signature", "Priority", "Action Needed")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    Dim row As Long
    row = 2
    
    ' Create function mapping
    Dim allFunctions As Object
    Set allFunctions = CreateObject("Scripting.Dictionary")
    
    ' Add Python functions
''    Dim i As Long
  ''  If UBound(pythonFunctions) > 0 Then
    ''    For i = 1 To UBound(pythonFunctions)
      ''      If Not allFunctions.exists(pythonFunctions(i).name) Then
        ''        Set allFunctions(pythonFunctions(i).name) = CreateObject("Scripting.Dictionary")
          ''      allFunctions(pythonFunctions(i).name)("python") = pythonFunctions(i)
            ''End If
  ''      Next i
   '' End If
    
    ' Add VBA functions
 ''   If UBound(vbaFunctions) > 0 Then
   ''     For i = 1 To UBound(vbaFunctions)
     ''       If Not allFunctions.exists(vbaFunctions(i).name) Then
       ''         Set allFunctions(vbaFunctions(i).name) = CreateObject("Scripting.Dictionary")
         ''   End If
           '' allFunctions(vbaFunctions(i).name)("vba") = vbaFunctions(i)
   ''     Next i
  ''  End If
    
    ' Populate rows
    Dim funcName As Variant
    For Each funcName In allFunctions.Keys
        Dim funcData As Object
        Set funcData = allFunctions(funcName)
        
        Dim Status As String, Priority As String, action As String
        Dim pythonFile As String, vbaFile As String
        Dim pythonSig As String, vbaSig As String
        
        If funcData.exists("python") And funcData.exists("vba") Then
            Status = "✅ Synchronized"
            Priority = "Low"
            action = "Ready - No action needed"
            pythonFile = funcData("python").fileName
            vbaFile = funcData("vba").fileName
            pythonSig = funcData("python").Signature
            vbaSig = funcData("vba").Signature
        ElseIf funcData.exists("python") Then
            Status = "🔄 Python Only"
            Priority = "High"
            action = "Convert Python to VBA"
            pythonFile = funcData("python").fileName
            vbaFile = ""
            pythonSig = funcData("python").Signature
            vbaSig = ""
        Else
            Status = "🔄 VBA Only"
            Priority = "Medium"
            action = "Create Python equivalent"
            pythonFile = ""
            vbaFile = funcData("vba").fileName
            pythonSig = ""
            vbaSig = funcData("vba").Signature
        End If
        
        ws.Cells(row, 1).value = funcName
        ws.Cells(row, 2).value = Status
        ws.Cells(row, 3).value = pythonFile
        ws.Cells(row, 4).value = vbaFile
        ws.Cells(row, 5).value = pythonSig
        ws.Cells(row, 6).value = vbaSig
        ws.Cells(row, 7).value = Priority
        ws.Cells(row, 8).value = action
        
        row = row + 1
    Next funcName
    
    ' Auto-fit columns
    ws.Columns.AutoFit
End Sub

' Create Sync Status Dashboard
Private Sub CreateSyncStatusDashboard(pythonFunctions() As FunctionInfo, vbaFunctions() As FunctionInfo)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sync_Dashboard")
    
    ' Title
    ws.Range("A1").value = "Development Environment Sync Dashboard"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ' Summary statistics
    Dim pythonCount As Long, vbaCount As Long
    pythonCount = IIf(UBound(pythonFunctions) > 0, UBound(pythonFunctions), 0)
    vbaCount = IIf(UBound(vbaFunctions) > 0, UBound(vbaFunctions), 0)
    
    ws.Range("A3").value = "Summary Statistics:"
    ws.Range("A3").Font.Bold = True
    
    ws.Range("A5:B10").value = Array(Array("Python Functions:", pythonCount), _
                                   Array("VBA Functions:", vbaCount), _
                                   Array("Total Functions:", pythonCount + vbaCount), _
                                   Array("Last Analysis:", Now()), _
                                   Array("Project Status:", "In Development"), _
                                   Array("Sync Required:", "Yes"))
    
    ' Quick Actions
    ws.Range("D3").value = "Quick Actions:"
    ws.Range("D3").Font.Bold = True
    
    ' Add buttons for common actions
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("D5").Left, ws.Range("D5").top, 120, 25)
    btn.text = "Refresh Analysis"
    btn.OnAction = "AnalyzeDevEnvironment"
    
    Set btn = ws.Buttons.Add(ws.Range("D7").Left, ws.Range("D7").top, 120, 25)
    btn.text = "Export Report"
    btn.OnAction = "ExportAnalysisReport"
    
    ' Format
    ws.Range("A5:A10").Font.Bold = True
    ws.Columns.AutoFit
End Sub

' Create Conversion Tracker
Private Sub CreateConversionTracker(pythonFunctions() As FunctionInfo, vbaFunctions() As FunctionInfo)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Conversion_Tracker")
    
    ' Create headers for conversion tracking
    With ws.Range("A1:F1")
        .value = Array("Function", "From", "To", "Status", "Assigned", "Due Date")
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
    End With
    
    Dim row As Long
    row = 2
    
    ' Add Python functions that need VBA conversion
    Dim i As Long
    If UBound(pythonFunctions) > 0 Then
        For i = 1 To UBound(pythonFunctions)
            ws.Cells(row, 1).value = pythonFunctions(i).name
            ws.Cells(row, 2).value = "Python"
            ws.Cells(row, 3).value = "VBA"
            ws.Cells(row, 4).value = "Pending"
            ws.Cells(row, 5).value = ""
            ws.Cells(row, 6).value = ""
            row = row + 1
        Next i
    End If
    
    ws.Columns.AutoFit
End Sub

' Format analysis worksheets
Private Sub FormatAnalysisWorksheets()
    Dim ws As Worksheet
    
    ' Format Function Overview
    Set ws = ThisWorkbook.Worksheets("Function_Overview")
    If ws.UsedRange.Rows.count > 1 Then
        ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).name = "FunctionOverviewTable"
        ws.ListObjects("FunctionOverviewTable").TableStyle = "TableStyleMedium2"
    End If
    
    ' Format Conversion Tracker
    Set ws = ThisWorkbook.Worksheets("Conversion_Tracker")
    If ws.UsedRange.Rows.count > 1 Then
        ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).name = "ConversionTrackerTable"
        ws.ListObjects("ConversionTrackerTable").TableStyle = "TableStyleMedium4"
    End If
End Sub

' Export analysis report
Public Sub ExportAnalysisReport()
    MsgBox "Analysis report functionality can be extended here.", vbInformation
End Sub

' Read text file helper
Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ReadTextFile = fileContent
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ReadTextFile = ""
End Function
