VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
'Attribute VB_Name = "ThisWorkbook"  ' commented for export
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Public Sub Workbook_Open()
    On Error GoTo Done
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
    ' Hide the ribbon only while this workbook is open
   '' Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"  'TODO:Create a button to show ribbon

Application.DisplayFullScreen = True
'Application.DisplayFormulaBar = False
'ActiveWindow.DisplayHeadings = False
'CommandBars.ExecuteMso "HideRibbon"
    ' 1) Clear search inputs
   ' ClearNameIfExists "SearchBox"
    'ClearNameIfExists "ValveNumSearchBox"
Call Clearfilters

    ' 2) Clear filters (table + slicers)
    ClearDataTableFilters "DataTable"
    ClearSlicersForTable "DataTable"

    ' 3) Ensure slicer pulse and refresh results (if the module is present)
    On Error Resume Next
    EnsurePulseCell False
    RefreshResults
    On Error GoTo Done

    ' 4) Show Dashboard, top-left, 100% zoom
    ShowDashboardTopLeft GetConfigSafe("DASHBOARD_SHEET", "Dashboard"), 100

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Unified close behavior: discard changes silently (no prompt) and restore UI.
    ' We do NOT cancel unless UI restoration fails; we mark Saved=True so Excel skips the save prompt.
    On Error GoTo FailClose
    Call WBInternal_RestoreUI
    ThisWorkbook.Saved = True          ' Discard edits silently
    Exit Sub                           ' Allow Excel to proceed with close
FailClose:
    ' If something unexpected happened, allow close but still suppress prompt.
    ThisWorkbook.Saved = True
End Sub

Public Sub Button_Exit_NoPrompt()
    ' Can be assigned to a shape/button: performs same discard+close logic intentionally.
    On Error GoTo Done
    Application.EnableEvents = False   ' Prevent double firing of Workbook_BeforeClose logic
    Call WBInternal_RestoreUI
    ThisWorkbook.Saved = True          ' Discard changes
    ThisWorkbook.Close SaveChanges:=False
Done:
    On Error Resume Next
    Application.EnableEvents = True
End Sub

Private Sub WBInternal_RestoreUI()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(\"Ribbon\", True)"
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayHeadings = True
End Sub


' ---------- helpers (local to ThisWorkbook) ----------

Public Sub ClearNameIfExists(ByVal nm As String)
    On Error Resume Next
    Dim r As Range
    Set r = ThisWorkbook.Names(nm).RefersToRange
    If Not r Is Nothing Then r.Cells(1, 1).Value = ""
End Sub

Public Sub ClearDataTableFilters(ByVal tableName As String)
    Dim lo As ListObject: Set lo = GetListObjectByName(tableName)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    If Not lo.AutoFilter Is Nothing Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    On Error GoTo 0
End Sub

Public Sub ClearSlicersForTable(ByVal tableName As String)
    Dim sc As SlicerCache
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        ' SourceName returns the ListObject name for table slicers
        If Len(sc.SourceName) > 0 Then
            If StrComp(sc.SourceName, tableName, vbTextCompare) = 0 Then
                sc.ClearManualFilter
            End If
        End If
    Next sc
    On Error GoTo 0
End Sub


Public Sub ShowDashboardTopLeft(ByVal dashName As String, ByVal zoomPct As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(dashName)
    On Error GoTo 0
    If ws Is Nothing Then
        If ThisWorkbook.Worksheets.Count > 0 Then
            Set ws = ThisWorkbook.Worksheets(1)
        Else
            Exit Sub
        End If
    End If

    ws.Activate
    With ActiveWindow
        .ScrollRow = 1
        .ScrollColumn = 1
        .Zoom = zoomPct
    End With
    ws.Range("A1").Select
End Sub

Public Function GetListObjectByName(ByVal name As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, name, vbTextCompare) = 0 Then
                Set GetListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function GetConfigSafe(ByVal key As String, ByVal defaultVal As String) As String
    On Error GoTo fallback
    Dim ws As Worksheet, lo As ListObject, r As Range
    Set ws = ThisWorkbook.Worksheets("ConfigSheet")
    Set lo = ws.ListObjects("ConfigTable")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo fallback
    For Each r In lo.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).Value), key, vbTextCompare) = 0 Then
            GetConfigSafe = CStr(r.Cells(1, 2).Value)
            Exit Function
        End If
    Next r
fallback:
    If Len(GetConfigSafe) = 0 Then GetConfigSafe = defaultVal
End Function


Private Sub Workbook_WindowResize(ByVal Wn As Window)
'CommandBars.ExecuteMso "HideRibbon"
End Sub
