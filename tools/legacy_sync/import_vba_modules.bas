Sub ImportVBAModules()
    ' Auto-generated VBA import script
    ' Run this macro in Excel to import updated modules
    
    Dim fso As Object
    Dim projectPath As String
    Dim vbcomp As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    projectPath = ThisWorkbook.Path
    
    ' Remove existing modules first (optional - comment out if you want to keep old versions)
    
    ' Remove and re-import export
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("export")
    On Error GoTo 0
    
    If fso.FileExists(projectPath & "\export.bas") Then
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(projectPath & "\export.bas")
        Debug.Print "Imported: export.bas"
    Else
        Debug.Print "File not found: export.bas"
    End If
    
    ' Remove and re-import mod_ModeDrivenSearch
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("mod_ModeDrivenSearch")
    On Error GoTo 0
    
    If fso.FileExists(projectPath & "\mod_ModeDrivenSearch.bas") Then
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(projectPath & "\mod_ModeDrivenSearch.bas")
        Debug.Print "Imported: mod_ModeDrivenSearch.bas"
    Else
        Debug.Print "File not found: mod_ModeDrivenSearch.bas"
    End If
    
    ' Remove and re-import temp_mod_ConfigTableTools
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("temp_mod_ConfigTableTools")
    On Error GoTo 0
    
    If fso.FileExists(projectPath & "\temp_mod_ConfigTableTools.bas") Then
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(projectPath & "\temp_mod_ConfigTableTools.bas")
        Debug.Print "Imported: temp_mod_ConfigTableTools.bas"
    Else
        Debug.Print "File not found: temp_mod_ConfigTableTools.bas"
    End If
    
    ' Remove and re-import mod_PrimaryConsolidatedModule
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("mod_PrimaryConsolidatedModule")
    On Error GoTo 0
    
    If fso.FileExists(projectPath & "\mod_PrimaryConsolidatedModule.bas") Then
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(projectPath & "\mod_PrimaryConsolidatedModule.bas")
        Debug.Print "Imported: mod_PrimaryConsolidatedModule.bas"
    Else
        Debug.Print "File not found: mod_PrimaryConsolidatedModule.bas"
    End If

    
    MsgBox "VBA module import complete! Check Immediate Window for details.", vbInformation
End Sub
