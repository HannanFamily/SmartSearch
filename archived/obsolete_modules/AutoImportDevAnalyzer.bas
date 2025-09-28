Attribute VB_Name = "AutoImportDevAnalyzer"
'============================================================
' Auto-Import Development Environment Analyzer
' Automatically imports the DevEnvironmentAnalyzer module
'============================================================
Option Explicit

Sub ImportAndRunDevAnalyzer()
    On Error GoTo ErrorHandler
    
    ' Import the DevEnvironmentAnalyzer module if it doesn't exist
    If Not ModuleExists("DevEnvironmentAnalyzer") Then
        Dim modulePath As String
        modulePath = ThisWorkbook.Path & "\DevEnvironmentAnalyzer.bas"
        
        If Dir(modulePath) <> "" Then
            ThisWorkbook.VBProject.VBComponents.Import modulePath
            MsgBox "DevEnvironmentAnalyzer module imported successfully!", vbInformation
        Else
            MsgBox "DevEnvironmentAnalyzer.bas file not found in project directory.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Run the analysis
    Application.Run "DevEnvironmentAnalyzer.AnalyzeDevEnvironment"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & _
           "Make sure macros are enabled and you have permission to import VBA modules.", vbCritical
End Sub

' Check if a VBA module exists
Private Function ModuleExists(moduleName As String) As Boolean
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    ModuleExists = (Not vbc Is Nothing)
    On Error GoTo 0
End Function

' Quick setup routine
Sub QuickSetupDevEnvironment()
    MsgBox "Development Environment Setup" & vbCrLf & vbCrLf & _
           "This will:" & vbCrLf & _
           "1. Import the DevEnvironmentAnalyzer module" & vbCrLf & _
           "2. Scan your Python and VBA files" & vbCrLf & _
           "3. Create analysis worksheets" & vbCrLf & _
           "4. Build sync dashboard" & vbCrLf & vbCrLf & _
           "Click OK to proceed...", vbInformation
           
    Call ImportAndRunDevAnalyzer
End Sub