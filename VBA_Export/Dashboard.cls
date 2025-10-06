''VERSION 1.0 CLASS
''BEGIN
''  MultiUse = -1  'True
''END
''Attribute VB_Name = "Dashboard"
''Attribute VB_GlobalNameSpace = False
''Attribute VB_Creatable = False
''Attribute VB_PredeclaredId = True
''Attribute VB_Exposed = True
' -------------------------------------------------------------------------
'Dashboard SHEET CLASS members
' -------------------------------------------------------------------------

Option Explicit
Private lastPulse As Variant  'lives for the lifetime of the sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Done
    Application.EnableEvents = False

    Dim rngPulse As Variant
    Set rngPulse = Nothing
    On Error Resume Next
    Set rngPulse = ThisWorkbook.Names("SlicerPulseCell").RefersToRange
    On Error GoTo 0

    Dim needRefresh As Boolean

    ' 1) Any dynamic input named range?
    Dim coll As Collection: Set coll = mod_SearchEngine_Enhanced.AllInputNamedRanges_Enhanced()
    Dim i As Long, r As Range
    For i = 1 To coll.Count
        Set r = coll(i)
        If Not Application.Intersect(Target, r) Is Nothing Then
            needRefresh = True
            Exit For
        End If
    Next i

    ' 2) Pulse changed directly on sheet
    If Not needRefresh Then
        If Not rngPulse Is Nothing Then
            If Not Application.Intersect(Target, rngPulse) Is Nothing Then needRefresh = True
        End If
    End If

    If needRefresh Then mod_SearchEngine_Enhanced.RefreshResults_Enhanced

Done:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    EnsurePulseCell True
    mod_SearchEngine_Enhanced.RefreshResults_Enhanced
End Sub

Private Sub Worksheet_Calculate()
    Dim c As Range
    On Error Resume Next
    Set c = ThisWorkbook.Names("SlicerPulseCell").RefersToRange
    On Error GoTo 0
    If c Is Nothing Then Exit Sub

    If IsEmpty(lastPulse) Then lastPulse = c.Value
    If c.Value <> lastPulse Then
        lastPulse = c.Value
    mod_SearchEngine_Enhanced.RefreshResults_Enhanced
    End If
End Sub


' ---------- local helpers ----------

Private Function ResolveNameToRange(ByVal nm As String) As Range
    On Error Resume Next
    Dim r As Range
    If Len(Trim$(nm)) = 0 Then Exit Function
    Set r = ThisWorkbook.Names(nm).RefersToRange
    Set ResolveNameToRange = r
End Function

Private Function IntersectSafe(ByVal a As Range, ByVal b As Range) As Boolean
    On Error Resume Next
    If a Is Nothing Or b Is Nothing Then Exit Function
    IntersectSafe = Not Application.Intersect(a, b) Is Nothing
End Function

