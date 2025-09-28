Attribute VB_Name = "Dev_ModuleCatalog"
Option Explicit

' Generate a sheet catalog of modules and their purposes (lightweight)

Public Sub RUN_GenerateModuleCatalog()
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = SheetByName("ModuleCatalog")
    If Not ws Is Nothing Then Application.DisplayAlerts = False: ws.Delete: Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "ModuleCatalog"

    Dim r As Long: r = 1
    ws.Cells(r, 1).Value = "Module"
    ws.Cells(r, 2).Value = "Type"
    ws.Cells(r, 3).Value = "Purpose (summary)"
    ws.Rows(r).Font.Bold = True: r = r + 1

    Dim comps As Object, vbc As Object
    Set comps = ThisWorkbook.VBProject.VBComponents
    For Each vbc In comps
        Dim t As String
        Select Case vbc.Type
            Case 1: t = "StdModule"
            Case 2: t = "Class"
            Case 3: t = "UserForm"
            Case Else: t = "Document"
        End Select
        ws.Cells(r, 1).Value = vbc.Name
        ws.Cells(r, 2).Value = t
        ws.Cells(r, 3).Value = InferPurpose(vbc.Name)
        r = r + 1
    Next vbc

    ws.Columns("A:C").AutoFit
    MsgBox "Module catalog generated.", vbInformation
    Exit Sub
EH:
    MsgBox "Catalog failed: " & Err.Description, vbCritical
End Sub

Private Function InferPurpose(ByVal name As String) As String
    Dim n As String: n = UCase$(name)
    If n = "MOD_PRIMARYCONSOLIDATEDMODULE" Then InferPurpose = "Core search/config helpers"
    If n = "MOD_MODEDRIVENSEARCH" Then InferPurpose = "Config-driven mode search/output"
    If n = "MOD_SOOTBLOWERLOCATOR" Then InferPurpose = "Sootblower mode UI + logic"
    If n = "SOOTBLOWERFORMCREATOR" Then InferPurpose = "Dynamic SSB form builder"
    If n = "C_SSB_FORMEVENTS" Then InferPurpose = "WithEvents handler for SSB form"
    If n = "DATATABLEUPDATER" Then InferPurpose = "Apply cleaned data into DataTable"
    If n = "DEV_CONTROLCENTER" Then InferPurpose = "Convenience macros (sync/export/test)"
    If n = "DEV_SMOKETESTS" Then InferPurpose = "Workbook smoke test"
    If n = "DEV_EXPORTS" Then InferPurpose = "Project snapshot (tables/modules/env)"
    If Len(InferPurpose) = 0 Then InferPurpose = "(unspecified)"
End Function
