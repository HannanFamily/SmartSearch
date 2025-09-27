Attribute VB_Name = "MasterVBAImporter"
Option Explicit

'============================================================
' Master VBA Module Importer
' This module imports all other VBA modules automatically
' File: MasterVBAImporter.bas
'============================================================

Public Sub ImportAllVBAModules()
    '============================================================
    ' Import all VBA modules from the project directory
    '============================================================
    
    On Error GoTo ErrorHandler
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    Dim modulesToImport As Variant
    modulesToImport = Array( _
        "QuickDevAnalysis.bas", _
        "DevEnvironmentAnalyzer.bas", _
        "PythonVBAConverter.bas", _
        "FileSystemManager.bas", _
        "SyncManager.bas" _
    )
    
    Dim i As Long
    Dim importedCount As Long
    Dim failedCount As Long
    Dim results As String
    
    results = "VBA MODULE IMPORT RESULTS:" & vbCrLf & String(40, "=") & vbCrLf
    
    Application.ScreenUpdating = False
    
    For i = 0 To UBound(modulesToImport)
        Dim modulePath As String
        modulePath = projectPath & "\" & modulesToImport(i)
        
        If Dir(modulePath) <> "" Then
            If ImportSingleModule(modulePath, CStr(modulesToImport(i))) Then
                importedCount = importedCount + 1
                results = results & "✅ " & modulesToImport(i) & " - SUCCESS" & vbCrLf
            Else
                failedCount = failedCount + 1
                results = results & "❌ " & modulesToImport(i) & " - FAILED" & vbCrLf
            End If
        Else
            results = results & "⚠️ " & modulesToImport(i) & " - FILE NOT FOUND" & vbCrLf
        End If
    Next i
    
    results = results & vbCrLf & "SUMMARY:" & vbCrLf
    results = results & "• Imported: " & importedCount & vbCrLf
    results = results & "• Failed: " & failedCount & vbCrLf
    results = results & "• Total: " & (UBound(modulesToImport) + 1) & vbCrLf
    
    Application.ScreenUpdating = True
    
    MsgBox results, vbInformation, "VBA Import Complete"
    
    ' Run the development analysis if QuickDevAnalysis was imported
    If ModuleExists("QuickDevAnalysis") Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Would you like to run the Development Environment Analysis now?", _
                         vbYesNo + vbQuestion, "Run Analysis?")
        If response = vbYes Then
            Application.Run "QuickDevAnalysis.AnalyzeDevEnvironment"
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during import: " & Err.DESCRIPTION, vbCritical
End Sub

Private Function ImportSingleModule(filePath As String, fileName As String) As Boolean
    '============================================================
    ' Import a single VBA module
    '============================================================
    
    On Error GoTo ErrorHandler
    
    ' Get module name without extension
    Dim moduleName As String
    moduleName = Replace(fileName, ".bas", "")
    moduleName = Replace(moduleName, ".cls", "")
    
    ' Check if module already exists
    If ModuleExists(moduleName) Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Module '" & moduleName & "' already exists. Replace it?", _
                         vbYesNoCancel + vbQuestion, "Module Exists")
        
        If response = vbCancel Then
            ImportSingleModule = False
            Exit Function
        ElseIf response = vbYes Then
            ' Remove existing module
            ThisWorkbook.VBProject.VBComponents.Remove _
                ThisWorkbook.VBProject.VBComponents(moduleName)
        Else
            ImportSingleModule = False
            Exit Function
        End If
    End If
    
    ' Import the module
    ThisWorkbook.VBProject.VBComponents.Import filePath
    ImportSingleModule = True
    Exit Function
    
ErrorHandler:
    ImportSingleModule = False
End Function

Private Function ModuleExists(moduleName As String) As Boolean
    '============================================================
    ' Check if a VBA module exists in the workbook
    '============================================================
    
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    ModuleExists = (Not vbc Is Nothing)
    On Error GoTo 0
End Function

Public Sub ListAvailableModules()
    '============================================================
    ' List all available VBA modules in the project directory
    '============================================================
    
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    Dim moduleList As String
    moduleList = "AVAILABLE VBA MODULES:" & vbCrLf & String(30, "=") & vbCrLf
    
    ' Check for .bas files
    Dim fileName As String
    fileName = Dir(projectPath & "\*.bas")
    Do While fileName <> ""
        If fileName <> "MasterVBAImporter.bas" Then ' Don't list self
            moduleList = moduleList & "📋 " & fileName
            If Dir(projectPath & "\" & fileName) <> "" Then
                moduleList = moduleList & " ✅" & vbCrLf
            Else
                moduleList = moduleList & " ❌" & vbCrLf
            End If
        End If
        fileName = Dir()
    Loop
    
    ' Check for .cls files
    fileName = Dir(projectPath & "\*.cls")
    Do While fileName <> ""
        moduleList = moduleList & "📋 " & fileName
        If Dir(projectPath & "\" & fileName) <> "" Then
            moduleList = moduleList & " ✅" & vbCrLf
        Else
            moduleList = moduleList & " ❌" & vbCrLf
        End If
        fileName = Dir()
    Loop
    
    MsgBox moduleList, vbInformation, "Available Modules"
End Sub

Public Sub CreateModuleImportWorksheet()
    '============================================================
    ' Create a worksheet for managing VBA module imports
    '============================================================
    
    Application.ScreenUpdating = False
    
    ' Create or clear worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Module_Manager")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Module_Manager"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    ws.Range("A1:F1").Value = Array("Module Name", "File Type", "Status", "Last Modified", "Description", "Action")
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    Dim row As Long
    row = 2
    
    ' Add module information
    Dim projectPath As String
    projectPath = ThisWorkbook.path
    
    ' Add .bas files
    Dim fileName As String
    fileName = Dir(projectPath & "\*.bas")
    Do While fileName <> ""
        If fileName <> "MasterVBAImporter.bas" Then
            ws.Cells(row, 1).Value = fileName
            ws.Cells(row, 2).Value = "Module"
            ws.Cells(row, 3).Value = IIf(ModuleExists(Replace(fileName, ".bas", "")), "✅ Imported", "⏳ Available")
            On Error Resume Next
            ws.Cells(row, 4).Value = FileDateTime(projectPath & "\" & fileName)
            On Error GoTo 0
            ws.Cells(row, 5).Value = GetModuleDescription(fileName)
            ws.Cells(row, 6).Value = "Ready for Import"
            row = row + 1
        End If
        fileName = Dir()
    Loop
    
    ' Add .cls files
    fileName = Dir(projectPath & "\*.cls")
    Do While fileName <> ""
        ws.Cells(row, 1).Value = fileName
        ws.Cells(row, 2).Value = "Class"
        ws.Cells(row, 3).Value = IIf(ModuleExists(Replace(fileName, ".cls", "")), "✅ Imported", "⏳ Available")
        On Error Resume Next
        ws.Cells(row, 4).Value = FileDateTime(projectPath & "\" & fileName)
        On Error GoTo 0
        ws.Cells(row, 5).Value = GetModuleDescription(fileName)
        ws.Cells(row, 6).Value = "Ready for Import"
        row = row + 1
        fileName = Dir()
    Loop
    
    ' Format worksheet
    ws.Columns.AutoFit
    ws.Range("A1:F" & row - 1).AutoFilter
    
    ' Add action buttons
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top, 120, 25)
    btn.text = "Import All Modules"
    btn.OnAction = "ImportAllVBAModules"
    
    Set btn = ws.Buttons.Add(ws.Range("H4").Left, ws.Range("H4").Top, 120, 25)
    btn.text = "Refresh List"
    btn.OnAction = "CreateModuleImportWorksheet"
    
    Set btn = ws.Buttons.Add(ws.Range("H6").Left, ws.Range("H6").Top, 120, 25)
    btn.text = "Run Dev Analysis"
    btn.OnAction = "QuickDevAnalysis.AnalyzeDevEnvironment"
    
    Application.ScreenUpdating = True
    
    ws.Activate
    MsgBox "Module Manager worksheet created! Use the buttons to manage your VBA modules.", vbInformation
    
End Sub

Private Function GetModuleDescription(fileName As String) As String
    '============================================================
    ' Get description for each module based on filename
    '============================================================
    
    Select Case fileName
        Case "QuickDevAnalysis.bas"
            GetModuleDescription = "Development environment analysis and Python/VBA comparison"
        Case "DevEnvironmentAnalyzer.bas"
            GetModuleDescription = "Advanced development environment analyzer with detailed reporting"
        Case "PythonVBAConverter.bas"
            GetModuleDescription = "Tools for converting between Python and VBA code"
        Case "FileSystemManager.bas"
            GetModuleDescription = "File system operations and project management"
        Case "SyncManager.bas"
            GetModuleDescription = "Synchronization between Python and VBA environments"
        Case "Dashboard.cls"
            GetModuleDescription = "Main dashboard class for worksheet events"
        Case "ThisWorkbook.cls"
            GetModuleDescription = "Workbook-level events and initialization"
        Case Else
            GetModuleDescription = "VBA module for project functionality"
    End Select
    
End Function

Public Sub QuickSetup()
    '============================================================
    ' Quick setup routine - imports modules and runs analysis
    '============================================================
    
    MsgBox "Quick Setup will:" & vbCrLf & vbCrLf & _
           "1. Import all available VBA modules" & vbCrLf & _
           "2. Create Module Manager worksheet" & vbCrLf & _
           "3. Run Development Environment Analysis" & vbCrLf & vbCrLf & _
           "This will give you complete visibility into your Python/VBA project!", vbInformation
    
    ' Import all modules
    Call ImportAllVBAModules
    
    ' Create management worksheet
    Call CreateModuleImportWorksheet
    
    MsgBox "✅ Quick Setup Complete!" & vbCrLf & vbCrLf & _
           "Check the new worksheets:" & vbCrLf & _
           "• Module_Manager - Manage VBA imports" & vbCrLf & _
           "• Dev_Analysis - Python/VBA comparison" & vbCrLf & vbCrLf & _
           "You now have a complete development environment!", vbInformation
    
End Sub
