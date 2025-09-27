Attribute VB_Name = "ActiveModuleImporter"
Option Explicit

' Active Module Importer
' ----------------------
' Purpose: Maintain an "ActiveModules" folder next to the workbook as the source of truth.
' This module provides:
'   - ReplaceAllModules_FromActiveFolder: Purge non-document modules and import all .bas/.cls from ActiveModules.
'   - SyncModules_FromActiveFolder: Non-destructive sync; imports/updates modules from ActiveModules without purging.
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
    backupRoot = ThisWorkbook.Path & Application.PathSeparator & "Old_Code"
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
            If UCase$(vbc.Name) <> "ACTIVEMODULEIMPORTER" Then
                ' queue for removal (cannot remove while iterating) — but never remove the running importer
                toRemove.Add vbc
            End If
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
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' First import .bas (modules), then .cls (classes) for stability
    f = Dir(srcFolder & Application.PathSeparator & "*.bas")
    Do While Len(f) > 0
        filePath = srcFolder & Application.PathSeparator & f
        If Not ShouldSkipFile(f) And Not IsImporterFile(f) Then
            If ShouldRemoveTarget(f) Then
                RemoveExistingByFileName f
            Else
                If fso.FileExists(filePath) Then
                    On Error GoTo ImportError
                    comps.Import filePath
                    On Error GoTo ErrHandler
                Else
                    MsgBox "File not found: " & filePath, vbExclamation, "Import Error"
                End If
            End If
        End If
        f = Dir()
    Loop

    f = Dir(srcFolder & Application.PathSeparator & "*.cls")
    Do While Len(f) > 0
        filePath = srcFolder & Application.PathSeparator & f

        Dim modName As String
        modName = Left$(f, InStrRev(f, ".") - 1)

        If ShouldSkipFile(f) Or IsImporterFile(f) Then
            ' skip
        ElseIf ShouldRemoveTarget(f) Then
            RemoveExistingByFileName f ' do not import delete markers
        ElseIf DocumentModuleExists(modName) Then
            ' Replace code inside the document module
            Dim content As String
            content = ReadFileStrippingAttributes(filePath)
            ReplaceDocumentModuleCode modName, content
        Else
            ' Import as a regular class module
            If fso.FileExists(filePath) Then
                On Error GoTo ImportError
                comps.Import filePath
                On Error GoTo ErrHandler
            Else
                MsgBox "File not found: " & filePath, vbExclamation, "Import Error"
            End If
        End If

        f = Dir()
    Loop

    Application.ScreenUpdating = True
    MsgBox "Modules replaced from ActiveModules successfully." & vbCrLf & _
           "Backup exported to:" & vbCrLf & exportPath, vbInformation, "Import Complete"
    Exit Sub

ImportError:
    Application.ScreenUpdating = True
    MsgBox "Import failed for file: " & f & vbCrLf & _
           "Path: " & filePath & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Check if:" & vbCrLf & _
           "- File exists and is readable" & vbCrLf & _
           "- VBA Trust access is enabled" & vbCrLf & _
           "- File is not corrupted", vbCritical, "Import Failed"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error replacing modules: " & Err.Description, vbCritical, "Import Failed"
End Sub

Public Sub SyncModules_FromActiveFolder()
    ' Non-destructive: import or update modules that exist in ActiveModules
    On Error GoTo ErrHandler

    If Not HasVBATrustAccess() Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Trust Center and try again.", vbExclamation, "VBA Access Required"
        Exit Sub
    End If

    Dim srcFolder As String
    srcFolder = GetActiveModulesFolder()
    If Len(Dir(srcFolder, vbDirectory)) = 0 Then
        MkDir srcFolder
        MsgBox "Created ActiveModules folder here:" & vbCrLf & srcFolder, vbInformation, "ActiveModules Created"
        Exit Sub
    End If

    Dim comps As Object
    Set comps = ThisWorkbook.VBProject.VBComponents

    Dim f As String, filePath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Import .bas
    f = Dir(srcFolder & Application.PathSeparator & "*.bas")
    Do While Len(f) > 0
        filePath = srcFolder & Application.PathSeparator & f
        If Not ShouldSkipFile(f) And Not IsImporterFile(f) Then
            If ShouldRemoveTarget(f) Then
                RemoveExistingByFileName f
            End If
            Dim baseName As String
            baseName = Left$(f, InStrRev(f, ".") - 1)
            If ModuleExists(baseName) Then RemoveExistingModule baseName
            If fso.FileExists(filePath) Then
                On Error GoTo SyncError
                comps.Import filePath
                On Error GoTo ErrHandler
            End If
        End If
        f = Dir()
    Loop

    ' Import .cls
    f = Dir(srcFolder & Application.PathSeparator & "*.cls")
    Do While Len(f) > 0
        filePath = srcFolder & Application.PathSeparator & f
        Dim modName As String
        modName = Left$(f, InStrRev(f, ".") - 1)

        If Not ShouldSkipFile(f) And Not IsImporterFile(f) Then
            If ShouldRemoveTarget(f) Then RemoveExistingByFileName f
            If DocumentModuleExists(modName) Then
                Dim content As String
                content = ReadFileStrippingAttributes(filePath)
                ReplaceDocumentModuleCode modName, content
            Else
                If ModuleExists(modName) Then RemoveExistingModule modName
                If fso.FileExists(filePath) Then
                    On Error GoTo SyncError
                    comps.Import filePath
                    On Error GoTo ErrHandler
                End If
            End If
        End If
        f = Dir()
    Loop

    MsgBox "Sync complete from ActiveModules.", vbInformation
    Exit Sub

SyncError:
    MsgBox "Sync failed for file: " & f & vbCrLf & _
           "Path: " & filePath & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Sync Failed"
    Exit Sub

ErrHandler:
    MsgBox "Error during sync: " & Err.Description, vbCritical
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
    MsgBox "Error exporting modules: " & Err.Description, vbCritical, "Export Failed"
End Sub

Public Sub OpenActiveModulesFolder()
    Dim folder As String
    folder = GetActiveModulesFolder()
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    Shell "explorer """ & folder & """", vbNormalFocus
End Sub

Private Function GetActiveModulesFolder() As String
    GetActiveModulesFolder = ThisWorkbook.Path & Application.PathSeparator & "ActiveModules"
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
                vbc.Export targetFolder & Application.PathSeparator & vbc.Name & ".bas"
                On Error GoTo 0
            Case CT_ClassModule
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.Name & ".cls"
                On Error GoTo 0
            Case CT_MSForm
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.Name & ".frm"
                On Error GoTo 0
            Case CT_Document
                ' Export document modules as .cls for editing/reference
                On Error Resume Next
                vbc.Export targetFolder & Application.PathSeparator & vbc.Name & ".cls"
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

Private Function ModuleExists(moduleName As String) As Boolean
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
    ModuleExists = Not (vbc Is Nothing)
    On Error GoTo 0
End Function

Private Sub RemoveExistingModule(moduleName As String)
    On Error Resume Next
    If ThisWorkbook.VBProject.VBComponents(moduleName).Type <> CT_Document Then
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)
    End If
    On Error GoTo 0
End Sub

Private Sub RemoveExistingByFileName(fileName As String)
    ' If a file is named with markers like [REMOVE]ModuleName.bas, remove ModuleName if present
    Dim base As String
    base = Left$(fileName, InStrRev(fileName, ".") - 1)
    base = NormalizeName(base)
    If ModuleExists(base) Then RemoveExistingModule base
End Sub

Private Function ShouldSkipFile(fileName As String) As Boolean
    Dim n As String: n = UCase$(fileName)
    ShouldSkipFile = (InStr(n, "[SKIP]") > 0)
End Function

Private Function ShouldRemoveTarget(fileName As String) As Boolean
    Dim n As String: n = UCase$(fileName)
    ShouldRemoveTarget = (InStr(n, "[REMOVE]") > 0) Or (InStr(n, "[OBSOLETE]") > 0)
End Function

Private Function NormalizeName(ByVal raw As String) As String
    Dim s As String: s = raw
    s = Replace$(s, "[REMOVE]", "")
    s = Replace$(s, "[OBSOLETE]", "")
    s = Replace$(s, "[SKIP]", "")
    s = Trim$(s)
    NormalizeName = s
End Function

Private Function IsImporterFile(ByVal fileName As String) As Boolean
    Dim base As String
    base = NormalizeName(Left$(fileName, InStrRev(fileName, ".") - 1))
    IsImporterFile = (UCase$(base) = "ACTIVEMODULEIMPORTER")
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
    MsgBox "Failed to replace code in document module '" & moduleName & "': " & Err.Description, vbExclamation, "Replace Failed"
End Sub
