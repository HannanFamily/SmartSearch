Attribute VB_Name = "SimpleModuleImporter"
Option Explicit

' Simple, bulletproof module importer
' No fancy features - just basic import with proper error handling

Public Sub ImportFromActiveModules()
    On Error GoTo ErrHandler
    
    ' Check VBA access
    If Not HasVBAAccess() Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Excel Options > Trust Center.", vbExclamation
        Exit Sub
    End If
    
    ' Get folder path
    Dim srcFolder As String
    srcFolder = ThisWorkbook.Path & Application.PathSeparator & "ActiveModules"
    
    ' Check folder exists
    If Len(Dir(srcFolder, vbDirectory)) = 0 Then
        MsgBox "ActiveModules folder not found at: " & srcFolder, vbExclamation
        Exit Sub
    End If
    
    ' Import .bas files
    Dim fileName As String
    Dim filePath As String
    Dim importCount As Long
    importCount = 0
    
    fileName = Dir(srcFolder & Application.PathSeparator & "*.bas")
    Do While Len(fileName) > 0
        filePath = srcFolder & Application.PathSeparator & fileName
        
        ' Skip ourselves
        If UCase$(Left$(fileName, Len(fileName) - 4)) <> "SIMPLEMODULEIMPORTER" Then
            If ImportSingleFile(filePath, fileName) Then
                importCount = importCount + 1
            End If
        End If
        
        fileName = Dir()
    Loop
    
    ' Import .cls files
    fileName = Dir(srcFolder & Application.PathSeparator & "*.cls")
    Do While Len(fileName) > 0
        filePath = srcFolder & Application.PathSeparator & fileName
        
        If ImportSingleFile(filePath, fileName) Then
            importCount = importCount + 1
        End If
        
        fileName = Dir()
    Loop
    
    MsgBox "Import complete. Imported " & importCount & " modules.", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

Private Function ImportSingleFile(filePath As String, fileName As String) As Boolean
    On Error GoTo FileError
    
    ' Check file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        ImportSingleFile = False
        Exit Function
    End If
    
    ' Get module name
    Dim moduleName As String
    moduleName = Left$(fileName, InStrRev(fileName, ".") - 1)
    
    ' Remove existing module if it exists
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo FileError
    
    ' Import the file
    ThisWorkbook.VBProject.VBComponents.Import filePath
    ImportSingleFile = True
    Exit Function
    
FileError:
    MsgBox "Failed to import: " & fileName & vbCrLf & Err.Description, vbExclamation
    ImportSingleFile = False
End Function

Private Function HasVBAAccess() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    HasVBAAccess = (Err.Number = 0)
    On Error GoTo 0
End Function