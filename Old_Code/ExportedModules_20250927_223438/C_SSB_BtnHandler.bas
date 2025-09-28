Attribute VB_Name = "C_SSB_BtnHandler"
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Public ParentForm As Object
Public Role As String  ' "search" | "showall" | "assoc" | "close"

Private Sub Btn_Click()
    On Error Resume Next
    Dim num As String, grp As String
    grp = GetSelectedGroup_Runtime(ParentForm)
    num = GetNumberText_Runtime(ParentForm)

    Select Case LCase$(Role)
        Case "search"
            mod_SootblowerLocator.SB_ExecuteSearch num, grp
            SetStatus RuntimeLabel(ParentForm, "lblStatus"), "Search completed"
        Case "showall"
            mod_SootblowerLocator.SB_DisplayAll grp
            SetStatus RuntimeLabel(ParentForm, "lblStatus"), "Showing all"
        Case "assoc"
            mod_SootblowerLocator.SB_ShowAssociated num, grp
            SetStatus RuntimeLabel(ParentForm, "lblStatus"), "Associated updated"
        Case "close"
            Unload ParentForm
    End Select
End Sub

Private Sub SetStatus(ByVal lbl As Object, ByVal msg As String)
    On Error Resume Next
    If Not lbl Is Nothing Then lbl.caption = msg
End Sub

Private Function RuntimeLabel(ByVal frm As Object, ByVal name As String) As Object
    On Error Resume Next
    Set RuntimeLabel = frm.Controls(name)
End Function

Private Function GetNumberText_Runtime(ByVal frm As Object) As String
    On Error Resume Next
    Dim ctl As Object, firstText As Object
    ' Prefer a control tagged as role:number
    For Each ctl In frm.Controls
        If typeName(ctl) = "TextBox" Then
            If LCase$(ctl.tag) = "role:number" Then GetNumberText_Runtime = CStr(ctl.text): Exit Function
            If firstText Is Nothing Then Set firstText = ctl
        End If
    Next ctl
    If Not firstText Is Nothing Then GetNumberText_Runtime = CStr(firstText.text)
End Function

Private Function GetSelectedGroup_Runtime(ByVal frm As Object) As String
    On Error Resume Next
    Dim ctl As Object, firstAll As Object, firstRetr As Object, firstWall As Object
    ' Prefer tagged options
    For Each ctl In frm.Controls
        If typeName(ctl) = "OptionButton" Then
            Select Case LCase$(ctl.tag)
                Case "role:opt_all": If ctl.value Then GetSelectedGroup_Runtime = "": Exit Function
                Case "role:opt_retracts": If ctl.value Then GetSelectedGroup_Runtime = "Retracts": Exit Function
                Case "role:opt_wall": If ctl.value Then GetSelectedGroup_Runtime = "Wall": Exit Function
            End Select
            ' remember by order
            If firstAll Is Nothing Then Set firstAll = ctl ElseIf firstRetr Is Nothing Then Set firstRetr = ctl ElseIf firstWall Is Nothing Then Set firstWall = ctl
        End If
    Next ctl
    ' Fallback by order: 1=All, 2=Retracts, 3=Wall
    If Not firstAll Is Nothing And firstAll.value Then Exit Function ' returns ""
    If Not firstRetr Is Nothing And firstRetr.value Then GetSelectedGroup_Runtime = "Retracts": Exit Function
    If Not firstWall Is Nothing And firstWall.value Then GetSelectedGroup_Runtime = "Wall": Exit Function
    GetSelectedGroup_Runtime = ""
End Function
