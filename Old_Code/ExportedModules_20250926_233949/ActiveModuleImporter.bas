Attribute VB_Name = "ActiveModuleImporter"
Option Explicit

' Active Module Importer
' ----------------------
' Purpose: Maintain an "ActiveModules" folder next to the workbook as the source of truth.
' This module provides:
'   - ReplaceAllModules_FromActiveFolder: Purge non-document modules and import all .bas/.cls from ActiveModules.
'   - ExportModulesToActiveFolder: Export current modules into ActiveModules to seed/sync.
'
' Safe behaviors:
'   - Never removes document modules (ThisWorkbook, worksheet code-behind).
'   - If a .cls matches a document module name (e.g., "ThisWorkbook" or a sheet CodeName),
'     its code is written into that document module instead of importing a new component.
'   - Creates a timestamped backup export before destructive changes.
'
' Requirements:
'   - Excel Trust Center: Enable "Trust access to the VBA project object model".
'   - Place .bas/.cls files in: <WorkbookFolder>\ActiveModules
'
' Notes:
'   - Document modules cannot be removed/imported; we replace their code via CodeModule.
'   - We strip Attribute lines from files before inserting into document modules.

Private Const CT_StdModule As Long = 1
Private Const CT_ClassModule As Long = 2
Private Const CT_MSForm As Long = 3
Private Const CT_Document As Long = 100

Public Sub ReplaceAllModules_FromActiveFolder()
    On Error GoTo ErrHandler

    If Not HasVBATrustAccess() Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Trust Center and try again.", vbExclamation, "VBA Access Required"
        Exit Sub
    End If

    Dim srcFolder As String
    srcFolder = GetActiveModulesFolder()
    If Len(Dir(srcFolder, vbDirectory)) = 0 Then
        MkDir srcFolder
        MsgBox "Created ActiveModules folder here:" & vbCrLf & srcFolder & vbCrLf & _
               "Put your .bas and .cls files in this folder, then run this again.", vbInformation, "ActiveModules Created"
        Exit Sub
    End If

    ' Backup current modules to Old_Code/export
    Dim backupRoot As String
    backupRoot = ThisWorkbook.path & Application.PathSeparator & "Old_Code"
    If Len(Dir(backupRoot, vbDirectory)) = 0 Then MkDir backupRoot
    Dim exportPath As String
    exportPath = backupRoot & Application.PathSeparator & "ExportedModules_" & Format(Now, "yyyymmdd_hhnnss")
    MkDir exportPath
    ExportCurrentModulesToFolder exportPath

    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler

    ' Remove all non-document modules
    Dim vbc As Object
    Dim comps As Object
    Set comps = ThisWorkbook.VBProject.VBComponents

    Dim toRemove As Collection
    Set toRemove = New Collection

    For Each vbc In comps
        If vbc.Type <> CT_Document Then
            ' queue for removal (cannot remove while iterating)
            toRemove.Add vbc
        End If
    Next vbc

    Dim item As Variant
    For Each item In toRemove
        On Error Resume Next
        comps.Remove item
        On Error GoTo 0
    Next item

    ' Import all files from ActiveModules
    Dim f As String
    Dim filePath As String

    ' First import .bas (modules), then .cls (classes) for stability
    f = Dir(srcFolder & Application.PathSeparator & "*.bas")
    Do While Len(f) > 0
        filePath = srcFolder & Application.PathSeparator & f
        comps.Import filePath
        f = Dir()
    Loop

    f = Dir(srcFolder & Application.PathSeparator & "*.cls")
    Do While Len(f) > 0
        filePath = srcFolder & Application.PathSeparator & f

        Dim modName As String
        modName = Left$(f, InStrRev(f, ".") - 1)

        If DocumentModuleExists(modName) Then
            ' Replace code inside the document module
            Dim content As String
            content = ReadFileStrippingAttributes(filePath)
            ReplaceDocumentModuleCode modName, content
        Else
            ' Import as a regular class module
            comps.Import filePath
        End If

        f = Dir()
    Loop

    Application.ScreenUpdating = True
    MsgBox "Modules replaced from ActiveModules successfully." & vbCrLf & _
           "Backup exported to:" & vbCrLf & exportPath, vbInformation, "Import Complete"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error replacing modules: " & Err.DESCRIPTION, vbCritical, "Import Failed"
End Sub

Public Sub ExportModulesToActiveFolder()
    On Error GoTo ErrHandler

    If Not HasVBATrustAccess() Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Trust Center and try again.", vbExclamation, "VBA Access Required"
        Exit Sub
    End If

    Dim dstFolder As String
    dstFolder = GetActiveModulesFolder()
    If Len(Dir(dstFolder, vbDirectory)) = 0 Then MkDir dstFolder

    ExportCurrentModulesToFolder dstFolder
    MsgBox "Exported current modules to ActiveModules folder:" & vbCrLf & dstFolder, vbInformation, "Export Complete"
    Exit Sub

ErrHandler:
    MsgBox "Error exporting modules: " & Err.DESCRIPTION, vbCritical, "Export Failed"
End Sub

Public Sub OpenActiveModulesFolder()
    Dim folder As String
    folder = GetActiveModulesFolder()
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    Shell "explorer """ & folder & """", vbNormalFocus
End Sub

Private Function GetActiveModulesFolder() As String
    GetActiveModulesFolder = ThisWorkbook.path & Application.PathSeparator & "ActiveModules"
End Function

Private Function HasVBATrustAccess() As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = ThisWorkbook.VBProject.VBComponents
    HasVBATrustAccess = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub ExportCurrentModulesToFolder(ByVal targetFolder As String)
    Dim comps As Object
    Set comps = ThisWorkbook.VBProject.VBComponents

    Dim vbc As Object
    For Each vbc In comps
        Select Case vbc.Type
            Case CT_StdModule
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.name & ".bas"
                On Error GoTo 0
            Case CT_ClassModule
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.name & ".cls"
                On Error GoTo 0
            Case CT_MSForm
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.name & ".frm"
                On Error GoTo 0
            Case CT_Document
                ' Export document modules as .cls for editing/reference
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.name & ".cls"
                On Error GoTo 0
        End Select
    Next vbc
End Sub

Private Function DocumentModuleExists(ByVal moduleName As String) As Boolean
    On Error Resume Next
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not vbc Is Nothing Then
        DocumentModuleExists = (vbc.Type = CT_Document)
    Else
        DocumentModuleExists = False
    End If
    On Error GoTo 0
End Function

Private Function ReadFileStrippingAttributes(ByVal filePath As String) As String
    Dim f As Integer
    f = FreeFile(0)
    Dim line As String
    Dim buf As String
    Open filePath For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        If Left$(LTrim$(line), 9) <> "Attribute" Then
            buf = buf & line & vbCrLf
        End If
    Loop
    Close #f
    ReadFileStrippingAttributes = buf
End Function

Private Sub ReplaceDocumentModuleCode(ByVal moduleName As String, ByVal newContent As String)
    On Error GoTo ErrHandler
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    If vbc Is Nothing Then Exit Sub

    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(newContent) > 0 Then .AddFromString newContent
    End With
    Exit Sub

ErrHandler:
    MsgBox "Failed to replace code in document module '" & moduleName & "': " & Err.DESCRIPTION, vbExclamation, "Replace Failed"
End Sub


