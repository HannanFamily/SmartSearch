Attribute VB_Name = "SyncManager"
Option Explicit

'============================================================
' Synchronization Manager
' Manages synchronization between Python and VBA environments
' File: SyncManager.bas
'============================================================

Public Type SyncStatus
    ItemName As String
    ItemType As String
    PythonStatus As String
    VBAStatus As String
    LastSync As Date
    SyncRequired As Boolean
    Priority As String
End Type

Public Sub RunFullSync()
    '============================================================
    ' Run complete synchronization between Python and VBA
    '============================================================
    
    Application.ScreenUpdating = False
    
    MsgBox "Full Synchronization will:" & vbCrLf & vbCrLf & _
           "1. Analyze Python and VBA environments" & vbCrLf & _
           "2. Identify sync differences" & vbCrLf & _
           "3. Generate conversion recommendations" & vbCrLf & _
           "4. Create sync status dashboard" & vbCrLf & _
           "5. Export sync reports" & vbCrLf & vbCrLf & _
           "Continue?", vbInformation + vbYesNo
    
    ' Create sync status worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sync_Status")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Sync_Status"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    ws.Range("A1:H1").value = Array("Item Name", "Type", "Python Status", "VBA Status", "Sync Required", "Priority", "Last Sync", "Action")
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 0)
        .Font.Color = RGB(0, 0, 0)
    End With
    
    Dim row As Long
    row = 2
    
    ' Analyze and populate sync status
    Call AnalyzePythonVBASync(ws, row)
    
    ' Add sync summary
    Call AddSyncSummary(ws, row)
    
    ' Format worksheet
    ws.Columns.AutoFit
    ws.Range("A1:H" & row - 1).AutoFilter
    
    ' Add action buttons
    Call AddSyncActionButtons(ws)
    
    Application.ScreenUpdating = True
    
    MsgBox "✅ Full synchronization analysis complete!" & vbCrLf & _
           "Check the Sync_Status worksheet for detailed sync information.", vbInformation
    
    ws.Activate
    
End Sub

Private Sub AnalyzePythonVBASync(ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Analyze synchronization status between Python and VBA
    '============================================================
    
    ' Get Python functions
    Dim pythonFunctions As Object
    Set pythonFunctions = GetPythonFunctions()
    
    ' Get VBA functions
    Dim vbaFunctions As Object
    Set vbaFunctions = GetVBAFunctions()
    
    ' Create combined list of all functions
    Dim allFunctions As Object
    Set allFunctions = CreateObject("Scripting.Dictionary")
    
    ' Add Python functions to master list
    Dim key As Variant
    For Each key In pythonFunctions.Keys
        If Not allFunctions.exists(key) Then
            Set allFunctions(key) = CreateObject("Scripting.Dictionary")
        End If
        allFunctions(key)("python") = pythonFunctions(key)
    Next key
    
    ' Add VBA functions to master list
    For Each key In vbaFunctions.Keys
        If Not allFunctions.exists(key) Then
            Set allFunctions(key) = CreateObject("Scripting.Dictionary")
        End If
        allFunctions(key)("vba") = vbaFunctions(key)
    Next key
    
    ' Analyze each function
    For Each key In allFunctions.Keys
        Dim funcData As Object
        Set funcData = allFunctions(key)
        
        ws.Cells(row, 1).value = key
        ws.Cells(row, 2).value = "Function"
        
        ' Determine sync status
        If funcData.exists("python") And funcData.exists("vba") Then
            ws.Cells(row, 3).value = "✅ Present"
            ws.Cells(row, 4).value = "✅ Present"
            ws.Cells(row, 5).value = "No"
            ws.Cells(row, 6).value = "LOW"
            ws.Cells(row, 8).value = "Synchronized"
        ElseIf funcData.exists("python") Then
            ws.Cells(row, 3).value = "✅ Present"
            ws.Cells(row, 4).value = "❌ Missing"
            ws.Cells(row, 5).value = "Yes"
            ws.Cells(row, 6).value = "HIGH"
            ws.Cells(row, 8).value = "Convert Python to VBA"
        ElseIf funcData.exists("vba") Then
            ws.Cells(row, 3).value = "❌ Missing"
            ws.Cells(row, 4).value = "✅ Present"
            ws.Cells(row, 5).value = "Yes"
            ws.Cells(row, 6).value = "MEDIUM"
            ws.Cells(row, 8).value = "Create Python equivalent"
        End If
        
        ws.Cells(row, 7).value = Now()
        
        ' Color code based on priority
        If ws.Cells(row, 6).value = "HIGH" Then
            ws.Range("A" & row & ":H" & row).Interior.Color = RGB(255, 230, 230) ' Light red
        ElseIf ws.Cells(row, 6).value = "MEDIUM" Then
            ws.Range("A" & row & ":H" & row).Interior.Color = RGB(255, 255, 230) ' Light yellow
        Else
            ws.Range("A" & row & ":H" & row).Interior.Color = RGB(230, 255, 230) ' Light green
        End If
        
        row = row + 1
    Next key
    
End Sub

Private Function GetPythonFunctions() As Object
    '============================================================
    ' Get list of Python functions from python directory
    '============================================================
    
    Dim functions As Object
    Set functions = CreateObject("Scripting.Dictionary")
    
    Dim pythonDir As String
    pythonDir = ThisWorkbook.path & "\python\"
    
    If Dir(pythonDir, vbDirectory) = "" Then
        Set GetPythonFunctions = functions
        Exit Function
    End If
    
    Dim fileName As String
    fileName = Dir(pythonDir & "*.py")
    
    Do While fileName <> ""
        Call ExtractPythonFunctionsFromFile(pythonDir & fileName, fileName, functions)
        fileName = Dir()
    Loop
    
    Set GetPythonFunctions = functions
    
End Function

Private Sub ExtractPythonFunctionsFromFile(filePath As String, fileName As String, functions As Object)
    '============================================================
    ' Extract function names from a Python file
    '============================================================
    
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    On Error GoTo 0
    
    Dim lines() As String
    lines = Split(fileContent, vbLf)
    
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        If Left(line, 4) = "def " And InStr(line, "(") > 0 Then
            Dim funcName As String
            funcName = ExtractFunctionName(line, "def ")
            
            If funcName <> "" And Not funcName Like "__*" Then
                functions(funcName) = fileName
            End If
        End If
    Next i
    
End Sub

Private Function GetVBAFunctions() As Object
    '============================================================
    ' Get list of VBA functions from VBA files
    '============================================================
    
    Dim functions As Object
    Set functions = CreateObject("Scripting.Dictionary")
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path & "\"
    
    ' Scan .bas files
    Dim fileName As String
    fileName = Dir(projectPath & "*.bas")
    
    Do While fileName <> ""
        Call ExtractVBAFunctionsFromFile(projectPath & fileName, fileName, functions)
        fileName = Dir()
    Loop
    
    ' Scan .cls files
    fileName = Dir(projectPath & "*.cls")
    Do While fileName <> ""
        Call ExtractVBAFunctionsFromFile(projectPath & fileName, fileName, functions)
        fileName = Dir()
    Loop
    
    Set GetVBAFunctions = functions
    
End Function

Private Sub ExtractVBAFunctionsFromFile(filePath As String, fileName As String, functions As Object)
    '============================================================
    ' Extract function names from a VBA file
    '============================================================
    
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    On Error GoTo 0
    
    Dim lines() As String
    lines = Split(fileContent, vbLf)
    
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Check for Sub or Function definitions
        If (InStr(1, line, "Sub ", vbTextCompare) > 0 Or InStr(1, line, "Function ", vbTextCompare) > 0) And _
           InStr(line, "(") > 0 Then
            
            Dim funcName As String
            If InStr(1, line, "Sub ", vbTextCompare) > 0 Then
                funcName = ExtractFunctionName(line, "Sub ")
            Else
                funcName = ExtractFunctionName(line, "Function ")
            End If
            
            If funcName <> "" Then
                functions(funcName) = fileName
            End If
        End If
    Next i
    
End Sub

Private Function ExtractFunctionName(line As String, keyword As String) As String
    '============================================================
    ' Extract function name from a function definition line
    '============================================================
    
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, line, keyword, vbTextCompare) + Len(keyword)
    endPos = InStr(startPos, line, "(")
    
    If endPos > startPos Then
        Dim funcName As String
        funcName = Trim(Mid(line, startPos, endPos - startPos))
        
        ' Remove visibility keywords if present
        funcName = Replace(funcName, "Public ", "")
        funcName = Replace(funcName, "Private ", "")
        funcName = Trim(funcName)
        
        ExtractFunctionName = funcName
    Else
        ExtractFunctionName = ""
    End If
    
End Function

Private Sub AddSyncSummary(ws As Worksheet, ByRef row As Long)
    '============================================================
    ' Add synchronization summary
    '============================================================
    
    row = row + 2
    
    ws.Range("A" & row).value = "SYNCHRONIZATION SUMMARY:"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("A" & row).Font.Size = 12
    
    row = row + 1
    ws.Range("A" & row).value = "• Total items: " & Application.WorksheetFunction.Counta(ws.Range("A:A")) - 3
    
    row = row + 1
    ws.Range("A" & row).value = "• Synchronized: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "No")
    
    row = row + 1
    ws.Range("A" & row).value = "• Need sync: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "Yes")
    
    row = row + 1
    ws.Range("A" & row).value = "• High priority: " & Application.WorksheetFunction.CountIf(ws.Range("F:F"), "HIGH")
    
    row = row + 1
    ws.Range("A" & row).value = "• Medium priority: " & Application.WorksheetFunction.CountIf(ws.Range("F:F"), "MEDIUM")
    
    row = row + 1
    ws.Range("A" & row).value = "• Last analysis: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    
End Sub

Private Sub AddSyncActionButtons(ws As Worksheet)
    '============================================================
    ' Add action buttons to the sync worksheet
    '============================================================
    
    Dim btn As Button
    
    ' Refresh sync status button
    Set btn = ws.Buttons.Add(ws.Range("J2").Left, ws.Range("J2").top, 120, 25)
    btn.text = "Refresh Sync Status"
    btn.OnAction = "RunFullSync"
    
    ' Export sync report button
    Set btn = ws.Buttons.Add(ws.Range("J4").Left, ws.Range("J4").top, 120, 25)
    btn.text = "Export Sync Report"
    btn.OnAction = "ExportSyncReport"
    
    ' Quick fix high priority button
    Set btn = ws.Buttons.Add(ws.Range("J6").Left, ws.Range("J6").top, 120, 25)
    btn.text = "Show High Priority"
    btn.OnAction = "ShowHighPriorityItems"
    
End Sub

Public Sub ExportSyncReport()
    '============================================================
    ' Export synchronization report to text file
    '============================================================
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sync_Status")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Please run sync analysis first!", vbExclamation
        Exit Sub
    End If
    
    Dim filePath As String
    filePath = ThisWorkbook.path & "\Sync_Report_" & Format(Now, "yyyy-mm-dd_hh-nn-ss") & ".txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
    Print #fileNum, "PYTHON/VBA SYNCHRONIZATION REPORT"
    Print #fileNum, "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNum, String(50, "=")
    Print #fileNum, ""
    
    ' Export data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    For i = 1 To lastRow
        If ws.Cells(i, 1).value <> "" And i <= lastRow - 10 Then ' Exclude summary section
            Print #fileNum, ws.Cells(i, 1).value & " | " & _
                          ws.Cells(i, 2).value & " | " & _
                          ws.Cells(i, 3).value & " | " & _
                          ws.Cells(i, 4).value & " | " & _
                          ws.Cells(i, 5).value & " | " & _
                          ws.Cells(i, 6).value & " | " & _
                          ws.Cells(i, 8).value
        End If
    Next i
    
    Close #fileNum
    
    MsgBox "Sync report exported to: " & filePath, vbInformation
    
End Sub

Public Sub ShowHighPriorityItems()
    '============================================================
    ' Filter to show only high priority sync items
    '============================================================
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Sync_Status")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Please run sync analysis first!", vbExclamation
        Exit Sub
    End If
    
    ' Clear existing filter
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' Apply filter for HIGH priority items
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ws.Range("A1:H" & lastRow).AutoFilter Field:=6, Criteria1:="HIGH"
    
    ws.Activate
    MsgBox "Filtered to show HIGH priority items only." & vbCrLf & _
           "Use Data > Filter > Clear to show all items.", vbInformation
    
End Sub

Public Sub QuickSync()
    '============================================================
    ' Quick sync operation for immediate needs
    '============================================================
    
    MsgBox "Quick Sync will perform a rapid analysis of sync status." & vbCrLf & _
           "For detailed analysis, use 'Run Full Sync'.", vbInformation
    
    ' Run abbreviated sync
    Call RunFullSync
    
End Sub
