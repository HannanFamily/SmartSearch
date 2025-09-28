VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private lastPulse As Variant

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Done
    Application.EnableEvents = False

    Dim rngDesc As Range, rngValve As Range, rngPulse As Range
    Set rngDesc = ResolveNameToRange(GetConfigValue("InputCell_DescripSearch"))
    Set rngValve = ResolveNameToRange(GetConfigValue("InputCell_ValveNumSearch"))
    Set rngPulse = ResolveNameToRange("SlicerPulse")

    If IntersectSafe(Target, rngDesc) Or IntersectSafe(Target, rngValve) Or IntersectSafe(Target, rngPulse) Then
        RefreshResults
    End If

Done:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    EnsurePulseCell True    ' create the pulse and force one calc
    RefreshResults          ' initial paint
End Sub

Private Sub Worksheet_Calculate()
    Dim c As Range
    On Error Resume Next
    Set c = ResolveNameToRange("SlicerPulseCell")
    On Error GoTo 0
    If c Is Nothing Then Exit Sub

    If IsEmpty(lastPulse) Then lastPulse = c.value
    If c.value <> lastPulse Then
        lastPulse = c.value
        RefreshResults
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

