Sub QuickImportAnalyzer()
    '============================================================
    ' Quick Import Development Environment Analyzer
    ' Run this macro to import the analyzer functionality
    '============================================================
    
    ' First, let's add the analyzer code directly
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
    vbComp.Name = "DevEnvironmentAnalyzer"
    
    ' Add the analyzer code
    Dim code As String
    code = code & "Option Explicit" & vbCrLf & vbCrLf
    code = code & "Public Sub AnalyzeDevEnvironment()" & vbCrLf
    code = code & "    ' Quick analysis of Python/VBA differences" & vbCrLf
    code = code & "    Application.ScreenUpdating = False" & vbCrLf & vbCrLf
    code = code & "    ' Create analysis worksheet" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Worksheets(""Dev_Analysis"")" & vbCrLf
    code = code & "    If ws Is Nothing Then" & vbCrLf
    code = code & "        Set ws = ThisWorkbook.Worksheets.Add" & vbCrLf
    code = code & "        ws.Name = ""Dev_Analysis""" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        ws.Cells.Clear" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf & vbCrLf
    code = code & "    ' Set up headers" & vbCrLf
    code = code & "    ws.Range(""A1:F1"").Value = Array(""File Type"", ""File Name"", ""Functions Found"", ""Status"", ""Action Needed"", ""Priority"")" & vbCrLf
    code = code & "    ws.Range(""A1:F1"").Font.Bold = True" & vbCrLf
    code = code & "    ws.Range(""A1:F1"").Interior.Color = RGB(68, 114, 196)" & vbCrLf
    code = code & "    ws.Range(""A1:F1"").Font.Color = RGB(255, 255, 255)" & vbCrLf & vbCrLf
    code = code & "    Dim row As Long" & vbCrLf
    code = code & "    row = 2" & vbCrLf & vbCrLf
    code = code & "    ' Analyze Python files" & vbCrLf
    code = code & "    Call AnalyzePythonFiles(ws, row)" & vbCrLf & vbCrLf
    code = code & "    ' Analyze VBA files" & vbCrLf
    code = code & "    Call AnalyzeVBAFiles(ws, row)" & vbCrLf & vbCrLf
    code = code & "    ' Format and finish" & vbCrLf
    code = code & "    ws.Columns.AutoFit" & vbCrLf
    code = code & "    ws.Range(""A1"").Select" & vbCrLf & vbCrLf
    code = code & "    Application.ScreenUpdating = True" & vbCrLf
    code = code & "    MsgBox ""Analysis complete! Check the Dev_Analysis worksheet."", vbInformation" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub AnalyzePythonFiles(ws As Worksheet, ByRef row As Long)" & vbCrLf
    code = code & "    Dim pythonDir As String" & vbCrLf
    code = code & "    pythonDir = ThisWorkbook.Path & ""\python\""" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If Dir(pythonDir, vbDirectory) <> """" Then" & vbCrLf
    code = code & "        Dim fileName As String" & vbCrLf
    code = code & "        fileName = Dir(pythonDir & ""*.py"")" & vbCrLf
    code = code & "        Do While fileName <> """"" & vbCrLf
    code = code & "            ws.Cells(row, 1).Value = ""Python""" & vbCrLf
    code = code & "            ws.Cells(row, 2).Value = fileName" & vbCrLf
    code = code & "            ws.Cells(row, 3).Value = ""Functions detected""" & vbCrLf
    code = code & "            ws.Cells(row, 4).Value = ""Needs VBA conversion""" & vbCrLf
    code = code & "            ws.Cells(row, 5).Value = ""Convert to VBA""" & vbCrLf
    code = code & "            ws.Cells(row, 6).Value = ""High""" & vbCrLf
    code = code & "            row = row + 1" & vbCrLf
    code = code & "            fileName = Dir()" & vbCrLf
    code = code & "        Loop" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        ws.Cells(row, 1).Value = ""Python""" & vbCrLf
    code = code & "        ws.Cells(row, 2).Value = ""No python folder found""" & vbCrLf
    code = code & "        ws.Cells(row, 4).Value = ""Setup needed""" & vbCrLf
    code = code & "        row = row + 1" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub AnalyzeVBAFiles(ws As Worksheet, ByRef row As Long)" & vbCrLf
    code = code & "    Dim projectDir As String" & vbCrLf
    code = code & "    projectDir = ThisWorkbook.Path & ""\""" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Dim fileName As String" & vbCrLf
    code = code & "    fileName = Dir(projectDir & ""*.bas"")" & vbCrLf
    code = code & "    Do While fileName <> """"" & vbCrLf
    code = code & "        ws.Cells(row, 1).Value = ""VBA""" & vbCrLf
    code = code & "        ws.Cells(row, 2).Value = fileName" & vbCrLf
    code = code & "        ws.Cells(row, 3).Value = ""Functions detected""" & vbCrLf
    code = code & "        ws.Cells(row, 4).Value = ""Needs Python equivalent""" & vbCrLf
    code = code & "        ws.Cells(row, 5).Value = ""Create Python version""" & vbCrLf
    code = code & "        ws.Cells(row, 6).Value = ""Medium""" & vbCrLf
    code = code & "        row = row + 1" & vbCrLf
    code = code & "        fileName = Dir()" & vbCrLf
    code = code & "    Loop" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    ' Add the code to the module
    vbComp.CodeModule.AddFromString code
    
    MsgBox "DevEnvironmentAnalyzer imported successfully!" & vbCrLf & _
           "You can now run 'AnalyzeDevEnvironment' macro.", vbInformation
    
End Sub