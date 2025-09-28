Attribute VB_Name = "QuickDevAnalysis"
Option Explicit

'============================================================
' Quick Development Environment Analysis
' This module can be directly imported into Excel
' File: QuickDevAnalysis.bas
'============================================================

Public Sub AnalyzeDevEnvironment()
    '============================================================
    ' Main analysis routine - scans Python and VBA files
    '============================================================
    
    Application.ScreenUpdating = False
    
    ' Create or clear analysis worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dev_Analysis")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Dev_Analysis"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    ws.Range("A1:G1").value = Array("File Type", "File Name", "Status", "Action Needed", "Priority", "Last Modified", "Notes")
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    Dim row As Long
    row = 2
    
    ' Analyze Python files
    Call AnalyzePythonFiles(ws, row)
    
    ' Analyze VBA files
    Call AnalyzeVBAFiles(ws, row)
    
    ' Add summary
    Call AddAnalysisSummary(ws, row)
    
    ' Format and finalize
    Call FormatAnalysisWorksheet(ws, row)
    
    Application.ScreenUpdating = True
    
    MsgBox "✅ Development Environment Analysis Complete!" & vbCrLf & vbCrLf & _
           "Check the 'Dev_Analysis' worksheet to see:" & vbCrLf & _
           "• All your Python and VBA files" & vbCrLf & _
           "• What needs to be converted" & vbCrLf & _
           "• Priority levels for each task" & vbCrLf & _
           "• File modification dates" & vbCrLf & vbCrLf & _
           "Use the filter buttons to sort by priority or file type!", vbInformation
    
    ' Activate the analysis worksheet
    ws.Activate
    
End Sub

Private Sub AnalyzePythonFiles(ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Scan Python files in the python/ subdirectory
    '============================================================
    
    Dim pythonDir As String
    pythonDir = ThisWorkbook.path & "\python\"
    
    If Dir(pythonDir, vbDirectory) <> "" Then
        Dim fileName As String
        fileName = Dir(pythonDir & "*.py")
        
        Do While fileName <> ""
            ws.Cells(row, 1).value = "Python"
            ws.Cells(row, 2).value = fileName
            ws.Cells(row, 3).value = "🔄 Needs VBA Conversion"
            ws.Cells(row, 4).value = "Convert Python functions to VBA"
            ws.Cells(row, 5).value = "HIGH"
            
            On Error Resume Next
            ws.Cells(row, 6).value = FileDateTime(pythonDir & fileName)
            On Error GoTo 0
            
            ws.Cells(row, 7).value = "Python file - needs VBA equivalent for Excel integration"
            
            row = row + 1
            fileName = Dir()
        Loop
    Else
        ws.Cells(row, 1).value = "Python"
        ws.Cells(row, 2).value = "No python folder found"
        ws.Cells(row, 3).value = "⚠️ Setup Issue"
        ws.Cells(row, 4).value = "Create python/ directory"
        ws.Cells(row, 5).value = "MEDIUM"
        ws.Cells(row, 7).value = "Expected: python/ subdirectory with .py files"
        row = row + 1
    End If
    
End Sub

Private Sub AnalyzeVBAFiles(ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Scan VBA files in the project directory
    '============================================================
    
    Dim projectDir As String
    projectDir = ThisWorkbook.path & "\"
    
    ' Scan .bas files
    Dim fileName As String
    fileName = Dir(projectDir & "*.bas")
    
    Do While fileName <> ""
        ' Skip this analysis file to avoid recursion
        If fileName <> "QuickDevAnalysis.bas" And fileName <> "DevEnvironmentAnalyzer.bas" Then
            ws.Cells(row, 1).value = "VBA Module"
            ws.Cells(row, 2).value = fileName
            ws.Cells(row, 3).value = "🔄 Needs Python Equivalent"
            ws.Cells(row, 4).value = "Create Python version for AI testing"
            ws.Cells(row, 5).value = "MEDIUM"
            
            On Error Resume Next
            ws.Cells(row, 6).value = FileDateTime(projectDir & fileName)
            On Error GoTo 0
            
            ws.Cells(row, 7).value = "VBA module - Python version recommended for AI development"
            
            row = row + 1
        End If
        fileName = Dir()
    Loop
    
    ' Scan .cls files
    fileName = Dir(projectDir & "*.cls")
    Do While fileName <> ""
        ws.Cells(row, 1).value = "VBA Class"
        ws.Cells(row, 2).value = fileName
        ws.Cells(row, 3).value = "🔄 Needs Python Equivalent"
        ws.Cells(row, 4).value = "Create Python class version"
        ws.Cells(row, 5).value = "MEDIUM"
        
        On Error Resume Next
        ws.Cells(row, 6).value = FileDateTime(projectDir & fileName)
        On Error GoTo 0
        
        ws.Cells(row, 7).value = "VBA class - Python equivalent recommended for testing"
        
        row = row + 1
        fileName = Dir()
    Loop
    
End Sub

Private Sub AddAnalysisSummary(ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Add summary statistics to the analysis
    '============================================================
    
    row = row + 2
    
    ws.Range("A" & row).value = "ANALYSIS SUMMARY:"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("A" & row).Font.Size = 12
    
    row = row + 1
    ws.Range("A" & row).value = "• Total files analyzed: " & (row - 4)
    
    row = row + 1
    ws.Range("A" & row).value = "• Python files found: " & Application.WorksheetFunction.CountIf(ws.Range("A:A"), "Python")
    
    row = row + 1
    ws.Range("A" & row).value = "• VBA modules found: " & Application.WorksheetFunction.CountIf(ws.Range("A:A"), "VBA Module")
    
    row = row + 1
    ws.Range("A" & row).value = "• VBA classes found: " & Application.WorksheetFunction.CountIf(ws.Range("A:A"), "VBA Class")
    
    row = row + 1
    ws.Range("A" & row).value = "• High priority items: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "HIGH")
    
    row = row + 1
    ws.Range("A" & row).value = "• Medium priority items: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "MEDIUM")
    
    row = row + 2
    ws.Range("A" & row).value = "RECOMMENDATIONS:"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("A" & row).Font.Size = 12
    
    row = row + 1
    ws.Range("A" & row).value = "1. Focus on HIGH priority Python → VBA conversions first"
    
    row = row + 1
    ws.Range("A" & row).value = "2. Create Python equivalents for VBA modules to enable AI testing"
    
    row = row + 1
    ws.Range("A" & row).value = "3. Use the existing sync scripts to import converted VBA code"
    
    row = row + 1
    ws.Range("A" & row).value = "4. Re-run this analysis after each conversion batch"
    
End Sub

Private Sub FormatAnalysisWorksheet(ws As Worksheet, row As Long)
    '============================================================
    ' Format the analysis worksheet for better readability
    '============================================================
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    ' Add borders to data area
    Dim dataRange As Range
    Dim dataLastRow As Long
    dataLastRow = row - 4
    If dataLastRow < 1 Then dataLastRow = 1
    Set dataRange = ws.Range("A1:G" & dataLastRow)
    On Error Resume Next
    dataRange.Borders.LineStyle = xlContinuous
    dataRange.Borders.Weight = xlThin
    On Error GoTo 0
    
    ' Add filter to the table
    On Error Resume Next
    ws.Range("A1:G" & dataLastRow).AutoFilter
    On Error GoTo 0
    
    ' Freeze top row
    FreezeTopRowSafe ws
    
    ' Color-code priorities
    Dim i As Long
    For i = 2 To dataLastRow
        If ws.Cells(i, 5).value = "HIGH" Then
            ws.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 230, 230) ' Light red
        ElseIf ws.Cells(i, 5).value = "MEDIUM" Then
            ws.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 255, 230) ' Light yellow
        End If
    Next i
    
    ' Select first cell
    SafeActivateCell ws, "A1"
    
End Sub

'==== Helpers to avoid Select/Activate errors ====
Private Sub FreezeTopRowSafe(ws As Worksheet)
    On Error Resume Next
    Dim prev As Worksheet
    Set prev = ActiveSheet
    ' If there's no active window or sheet is protected, skip freezing
    If Application.Windows.count = 0 Then Exit Sub
    If ws.ProtectContents Then Exit Sub
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(2, 1).Activate ' A2
    ActiveWindow.FreezePanes = True
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

Private Sub SafeActivateCell(ws As Worksheet, ByVal addr As String)
    On Error Resume Next
    Dim prev As Worksheet
    Set prev = ActiveSheet
    If Application.Windows.count = 0 Then Exit Sub
    ws.Activate
    ws.Range(addr).Activate
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

Public Sub RefreshAnalysis()
    '============================================================
    ' Quick refresh of the analysis - can be called anytime
    '============================================================
    Call AnalyzeDevEnvironment
End Sub

Public Sub ExportAnalysisToText()
    '============================================================
    ' Export analysis results to a text file for external use
    '============================================================
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dev_Analysis")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Please run AnalyzeDevEnvironment first!", vbExclamation
        Exit Sub
    End If
    
    Dim filePath As String
    filePath = ThisWorkbook.path & "\Development_Analysis_Report.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
    Print #fileNum, "DEVELOPMENT ENVIRONMENT ANALYSIS REPORT"
    Print #fileNum, "Generated: " & Now()
    Print #fileNum, String(50, "=")
    Print #fileNum, ""
    
    ' Export the table data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    For i = 1 To lastRow
        If ws.Cells(i, 1).value <> "" Then
            Print #fileNum, ws.Cells(i, 1).value & " | " & _
                          ws.Cells(i, 2).value & " | " & _
                          ws.Cells(i, 3).value & " | " & _
                          ws.Cells(i, 4).value & " | " & _
                          ws.Cells(i, 5).value
        End If
    Next i
    
    Close #fileNum
    
    MsgBox "Analysis exported to: " & filePath, vbInformation
    
End Sub
