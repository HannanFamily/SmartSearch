Attribute VB_Name = "SootblowerFormFactory"
Option Explicit

Public Function CreateSootblowerUserForm() As Boolean
    On Error GoTo EH
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim frm As Object
    Set vbProj = ThisWorkbook.VBProject
    On Error Resume Next
    Set vbComp = vbProj.VBComponents("frmSootblowerLocator")
    On Error GoTo EH
    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
        vbComp.Name = "frmSootblowerLocator"
        Dim d As MSForms.UserForm: Set d = vbComp.Designer
        d.Caption = "Sootblower Locator": d.Width = 360: d.Height = 190
        Dim lbl As MSForms.Label
        Set lbl = d.Controls.Add("Forms.Label.1", "lblTitle", True)
        lbl.Caption = "Enter Sootblower Number (1-3 digits)": lbl.Left = 12: lbl.Top = 12: lbl.Width = 300
        Dim txt As MSForms.TextBox
        Set txt = d.Controls.Add("Forms.TextBox.1", "txtNumber", True)
        txt.Left = 12: txt.Top = 32: txt.Width = 120
        Dim optRetr As MSForms.ToggleButton, optWall As MSForms.ToggleButton
        Set optRetr = d.Controls.Add("Forms.ToggleButton.1", "tglRetracts", True)
        optRetr.Caption = "IK/EL (Retracts)": optRetr.Left = 12: optRetr.Top = 64: optRetr.Width = 150
        Set optWall = d.Controls.Add("Forms.ToggleButton.1", "tglWall", True)
        optWall.Caption = "IR/WB (Wall Blower)": optWall.Left = 180: optWall.Top = 64: optWall.Width = 150
        Dim btnSearch As MSForms.CommandButton, btnAll As MSForms.CommandButton, btnClose As MSForms.CommandButton, btnAssoc As MSForms.CommandButton
        Set btnSearch = d.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
        btnSearch.Caption = "Search": btnSearch.Left = 12: btnSearch.Top = 110: btnSearch.Width = 80
        Set btnAll = d.Controls.Add("Forms.CommandButton.1", "btnShowAll", True)
        btnAll.Caption = "Show All": btnAll.Left = 104: btnAll.Top = 110: btnAll.Width = 80
        Set btnAssoc = d.Controls.Add("Forms.CommandButton.1", "btnAssoc", True)
        btnAssoc.Caption = "Show all associated equipment": btnAssoc.Left = 12: btnAssoc.Top = 140: btnAssoc.Width = 264
        Set btnClose = d.Controls.Add("Forms.CommandButton.1", "btnClose", True)
        btnClose.Caption = "Close": btnClose.Left = 196: btnClose.Top = 110: btnClose.Width = 80
        Dim code As String
        code = "Option Explicit" & vbCrLf & _
               "Private Sub btnSearch_Click()" & vbCrLf & _
               "    Dim grp As String: grp = GetGroup()" & vbCrLf & _
               "    mod_SootblowerLocator.SB_ExecuteSearch Me.txtNumber.Text, grp" & vbCrLf & _
               "End Sub" & vbCrLf & _
               "Private Sub btnShowAll_Click()" & vbCrLf & _
               "    Dim grp As String: grp = GetGroup()" & vbCrLf & _
               "    mod_SootblowerLocator.SB_DisplayAll grp" & vbCrLf & _
               "End Sub" & vbCrLf & _
               "Private Sub btnClose_Click()" & vbCrLf & _
               "    Unload Me" & vbCrLf & _
               "End Sub" & vbCrLf & _
               "Private Sub btnAssoc_Click()" & vbCrLf & _
               "    Dim grp As String: grp = GetGroup()" & vbCrLf & _
               "    mod_SootblowerLocator.SB_ShowAssociated Me.txtNumber.Text, grp" & vbCrLf & _
               "End Sub" & vbCrLf & _
               "Private Sub txtNumber_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf & _
               "    If KeyCode = 13 Then" & vbCrLf & _
               "        KeyCode = 0" & vbCrLf & _
               "        btnSearch_Click" & vbCrLf & _
               "    End If" & vbCrLf & _
               "End Sub" & vbCrLf & _
               "Private Function GetGroup() As String" & vbCrLf & _
               "    If Me.tglRetracts.Value And Not Me.tglWall.Value Then" & vbCrLf & _
               "        GetGroup = ""Retracts""" & vbCrLf & _
               "    ElseIf Me.tglWall.Value And Not Me.tglRetracts.Value Then" & vbCrLf & _
               "        GetGroup = ""Wall""" & vbCrLf & _
               "    Else" & vbCrLf & _
               "        GetGroup = """"" & vbCrLf & _
               "    End If" & vbCrLf & _
               "End Function" & vbCrLf
        vbComp.CodeModule.InsertLines 1, code
    End If
    CreateSootblowerUserForm = True
    Exit Function
EH:
    CreateSootblowerUserForm = False
End Function
