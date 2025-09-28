Attribute VB_Name = "FileSystemManager"
Option Explicit

'============================================================
' File System Manager
' Handles file operations and project management
' File: FileSystemManager.bas
'============================================================

Public Type FileInfo
    fileName As String
    filePath As String
    FileSize As Long
    LastModified As Date
    fileType As String
    Status As String
End Type

Public Sub ScanProjectFiles()
    '============================================================
    ' Scan and catalog all project files
    '============================================================
    
    Application.ScreenUpdating = False
    
    ' Create or clear file system worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("File_System")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "File_System"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    ws.Range("A1:G1").Value = Array("File Name", "File Path", "File Type", "Size (KB)", "Last Modified", "Status", "Description")
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    Dim row As Long
    row = 2
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    ' Scan different file types
    Call ScanFilesByExtension(ws, row, projectPath, "*.py", "Python", "Python source file")
    Call ScanFilesByExtension(ws, row, projectPath, "*.bas", "VBA Module", "VBA module file")
    Call ScanFilesByExtension(ws, row, projectPath, "*.cls", "VBA Class", "VBA class module")
    Call ScanFilesByExtension(ws, row, projectPath, "*.xlsm", "Excel", "Excel macro-enabled workbook")
    Call ScanFilesByExtension(ws, row, projectPath, "*.md", "Documentation", "Markdown documentation")
    Call ScanFilesByExtension(ws, row, projectPath, "*.txt", "Text", "Text file")
    Call ScanFilesByExtension(ws, row, projectPath, "*.json", "JSON", "JSON configuration file")
    Call ScanFilesByExtension(ws, row, projectPath, "*.ps1", "PowerShell", "PowerShell script")
    Call ScanFilesByExtension(ws, row, projectPath, "*.sh", "Shell", "Shell script")
    
    ' Scan subdirectories
    Call ScanSubdirectories(ws, row, projectPath)
    
    ' Add summary
    Call AddFileSummary(ws, row)
    
    ' Format worksheet
    ws.Columns.AutoFit
    ws.Range("A1:G" & row - 1).AutoFilter
    
    Application.ScreenUpdating = True
    
    MsgBox "Project file scan complete!" & vbCrLf & _
           "Check the File_System worksheet for detailed file information.", vbInformation
    
    ws.Activate
    
End Sub

Private Sub ScanFilesByExtension(ws As Worksheet, ByRef row As Long, basePath As String, _
                                pattern As String, fileType As String, DESCRIPTION As String)
    '============================================================
    ' Scan files by extension pattern
    '============================================================
    
    Dim fileName As String
    fileName = Dir(basePath & "\" & pattern)
    
    Do While fileName <> ""
        Dim fullPath As String
        fullPath = basePath & "\" & fileName
        
        ws.Cells(row, 1).Value = fileName
        ws.Cells(row, 2).Value = fullPath
        ws.Cells(row, 3).Value = fileType
        
        ' Get file size
        On Error Resume Next
        ws.Cells(row, 4).Value = Round(FileLen(fullPath) / 1024, 2)
        ws.Cells(row, 5).Value = FileDateTime(fullPath)
        On Error GoTo 0
        
        ws.Cells(row, 6).Value = "Available"
        ws.Cells(row, 7).Value = DESCRIPTION
        
        row = row + 1
        fileName = Dir()
    Loop
    
End Sub

Private Sub ScanSubdirectories(ws As Worksheet, ByRef row As Long, basePath As String)
    '============================================================
    ' Scan important subdirectories
    '============================================================
    
    Dim subDirs As Variant
    subDirs = Array("python", "shared", "Old_Code", "ai_project_template")
    
    Dim i As Long
    For i = 0 To UBound(subDirs)
        Dim subDirPath As String
        subDirPath = basePath & "\" & subDirs(i)
        
        If Dir(subDirPath, vbDirectory) <> "" Then
            ' Scan Python files in python directory
            If subDirs(i) = "python" Then
                Call ScanFilesByExtension(ws, row, subDirPath, "*.py", "Python (subdir)", "Python file in python/ subdirectory")
            End If
            
            ' Add directory entry
            ws.Cells(row, 1).Value = subDirs(i) & "/"
            ws.Cells(row, 2).Value = subDirPath
            ws.Cells(row, 3).Value = "Directory"
            ws.Cells(row, 6).Value = "Available"
            ws.Cells(row, 7).Value = "Project subdirectory"
            row = row + 1
        End If
    Next i
    
End Sub

Private Sub AddFileSummary(ws As Worksheet, row As Long)
    '============================================================
    ' Add file summary statistics
    '============================================================
    
    row = row + 2
    
    ws.Range("A" & row).Value = "FILE SUMMARY:"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("A" & row).Font.Size = 12
    
    row = row + 1
    ws.Range("A" & row).Value = "• Python files: " & Application.WorksheetFunction.CountIfs(ws.Range("C:C"), "Python", ws.Range("C:C"), "Python (subdir)")
    
    row = row + 1
    ws.Range("A" & row).Value = "• VBA modules: " & Application.WorksheetFunction.CountIf(ws.Range("C:C"), "VBA Module")
    
    row = row + 1
    ws.Range("A" & row).Value = "• VBA classes: " & Application.WorksheetFunction.CountIf(ws.Range("C:C"), "VBA Class")
    
    row = row + 1
    ws.Range("A" & row).Value = "• Excel files: " & Application.WorksheetFunction.CountIf(ws.Range("C:C"), "Excel")
    
    row = row + 1
    ws.Range("A" & row).Value = "• Documentation: " & Application.WorksheetFunction.CountIf(ws.Range("C:C"), "Documentation")
    
    row = row + 1
    ws.Range("A" & row).Value = "• Total files: " & (Application.WorksheetFunction.Counta(ws.Range("A:A")) - 3) ' Subtract headers and summary
    
End Sub

Public Function GetFileInfo(filePath As String) As FileInfo
    '============================================================
    ' Get detailed information about a specific file
    '============================================================
    
    Dim info As FileInfo
    
    On Error GoTo ErrorHandler
    
    info.fileName = Dir(filePath)
    info.filePath = filePath
    info.FileSize = FileLen(filePath)
    info.LastModified = FileDateTime(filePath)
    info.fileType = GetFileTypeFromExtension(filePath)
    info.Status = "Available"
    
    GetFileInfo = info
    Exit Function
    
ErrorHandler:
    info.Status = "Error: " & Err.DESCRIPTION
    GetFileInfo = info
    
End Function

Private Function GetFileTypeFromExtension(filePath As String) As String
    '============================================================
    ' Determine file type from file extension
    '============================================================
    
    Dim extension As String
    extension = LCase(Right(filePath, Len(filePath) - InStrRev(filePath, ".")))
    
    Select Case extension
        Case "py"
            GetFileTypeFromExtension = "Python"
        Case "bas"
            GetFileTypeFromExtension = "VBA Module"
        Case "cls"
            GetFileTypeFromExtension = "VBA Class"
        Case "xlsm"
            GetFileTypeFromExtension = "Excel Macro"
        Case "xlsx"
            GetFileTypeFromExtension = "Excel"
        Case "md"
            GetFileTypeFromExtension = "Markdown"
        Case "txt"
            GetFileTypeFromExtension = "Text"
        Case "json"
            GetFileTypeFromExtension = "JSON"
        Case "ps1"
            GetFileTypeFromExtension = "PowerShell"
        Case "sh"
            GetFileTypeFromExtension = "Shell Script"
        Case Else
            GetFileTypeFromExtension = "Unknown"
    End Select
    
End Function

Public Sub BackupProjectFiles()
    '============================================================
    ' Create backup of important project files
    '============================================================
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    Dim backupDir As String
    backupDir = projectPath & "\Backup_" & Format(Now, "yyyy-mm-dd_hh-nn-ss")
    
    ' Create backup directory
    On Error Resume Next
    MkDir backupDir
    On Error GoTo 0
    
    If Dir(backupDir, vbDirectory) = "" Then
        MsgBox "Could not create backup directory: " & backupDir, vbCritical
        Exit Sub
    End If
    
    Dim fileCount As Long
    fileCount = 0
    
    ' Backup VBA files
    fileCount = fileCount + BackupFilesByPattern(projectPath, backupDir, "*.bas")
    fileCount = fileCount + BackupFilesByPattern(projectPath, backupDir, "*.cls")
    
    ' Backup Excel files
    fileCount = fileCount + BackupFilesByPattern(projectPath, backupDir, "*.xlsm")
    
    ' Backup Python files
    If Dir(projectPath & "\python", vbDirectory) <> "" Then
        fileCount = fileCount + BackupFilesByPattern(projectPath & "\python", backupDir, "*.py")
    End If
    
    ' Backup documentation
    fileCount = fileCount + BackupFilesByPattern(projectPath, backupDir, "*.md")
    
    MsgBox "Backup complete!" & vbCrLf & _
           "Files backed up: " & fileCount & vbCrLf & _
           "Backup location: " & backupDir, vbInformation
    
End Sub

Private Function BackupFilesByPattern(sourcePath As String, backupPath As String, pattern As String) As Long
    '============================================================
    ' Backup files matching a specific pattern
    '============================================================
    
    Dim fileName As String
    Dim count As Long
    
    fileName = Dir(sourcePath & "\" & pattern)
    
    Do While fileName <> ""
        On Error Resume Next
        FileCopy sourcePath & "\" & fileName, backupPath & "\" & fileName
        If Err.Number = 0 Then
            count = count + 1
        End If
        On Error GoTo 0
        fileName = Dir()
    Loop
    
    BackupFilesByPattern = count
    
End Function

Public Sub CleanupTempFiles()
    '============================================================
    ' Clean up temporary and backup files
    '============================================================
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    Dim cleanupCount As Long
    cleanupCount = 0
    
    ' Clean up Excel backup files
    Dim fileName As String
    fileName = Dir(projectPath & "\*.xlsm.backup*")
    
    Do While fileName <> ""
        On Error Resume Next
        Kill projectPath & "\" & fileName
        If Err.Number = 0 Then
            cleanupCount = cleanupCount + 1
        End If
        On Error GoTo 0
        fileName = Dir()
    Loop
    
    ' Clean up temp files
    fileName = Dir(projectPath & "\*.tmp")
    Do While fileName <> ""
        On Error Resume Next
        Kill projectPath & "\" & fileName
        If Err.Number = 0 Then
            cleanupCount = cleanupCount + 1
        End If
        On Error GoTo 0
        fileName = Dir()
    Loop
    
    MsgBox "Cleanup complete!" & vbCrLf & _
           "Files cleaned up: " & cleanupCount, vbInformation
    
End Sub

Public Function CheckProjectIntegrity() As Boolean
    '============================================================
    ' Check project integrity and report missing files
    '============================================================
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    Dim issues() As String
    Dim issueCount As Long
    ReDim issues(0)
    
    ' Check for essential directories
    If Dir(projectPath & "\python", vbDirectory) = "" Then
        issueCount = issueCount + 1
        ReDim Preserve issues(issueCount)
        issues(issueCount) = "Missing: python/ subdirectory"
    End If
    
    ' Check for essential VBA files
    Dim essentialVBA As Variant
    essentialVBA = Array("mod_ModeDrivenSearch.bas", "Dashboard.cls", "ThisWorkbook.cls")
    
    Dim i As Long
    For i = 0 To UBound(essentialVBA)
        If Dir(projectPath & "\" & essentialVBA(i)) = "" Then
            issueCount = issueCount + 1
            ReDim Preserve issues(issueCount)
            issues(issueCount) = "Missing: " & essentialVBA(i)
        End If
    Next i
    
    ' Report results
    If issueCount = 0 Then
        MsgBox "✅ Project integrity check passed!" & vbCrLf & _
               "All essential files are present.", vbInformation
        CheckProjectIntegrity = True
    Else
        Dim report As String
        report = "⚠️ Project integrity issues found:" & vbCrLf & vbCrLf
        For i = 1 To issueCount
            report = report & "• " & issues(i) & vbCrLf
        Next i
        MsgBox report, vbExclamation, "Project Integrity Check"
        CheckProjectIntegrity = False
    End If
    
End Function
