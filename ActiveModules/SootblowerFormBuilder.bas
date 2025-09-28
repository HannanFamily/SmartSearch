Attribute VB_Name = "SootblowerFormBuilder"
Option Explicit

' Builds a design-time UserForm for the Sootblower Locator UI using the VBIDE.
' - Requires Trust Center: enable "Trust access to the VBA project object model"
' - Avoids code injection; pair with mod_SSB_RuntimeBinder to wire events at runtime

Private Const CT_MSFORM As Long = 3 ' vbext_ct_MSForm without requiring VBIDE reference at compile-time

Public Function Ensure_SootblowerForm_Built(Optional ByVal forceRebuild As Boolean = False) As Object
    On Error GoTo EH
    Dim vbProj As Object, vbComps As Object, vbComp As Object, designer As Object
    Dim exists As Boolean

    Set vbProj = ThisWorkbook.VBProject ' Requires Trust Access
    Set vbComps = vbProj.VBComponents

    exists = ComponentExists(vbComps, "frmSootblowerLocator")
    If exists And forceRebuild Then
        On Error Resume Next
        vbComps.Remove vbComps("frmSootblowerLocator")
        On Error GoTo EH
    End If

    If Not ComponentExists(vbComps, "frmSootblowerLocator") Then
        Set vbComp = vbComps.Add(CT_MSFORM)
        vbComp.Name = "frmSootblowerLocator"
    Else
        Set vbComp = vbComps("frmSootblowerLocator")
    End If

    Set designer = vbComp.Designer

    ' Build UI if empty or forced
    If forceRebuild Or designer.Controls.Count = 0 Then
        ClearControls designer
        BuildSootblowerForm designer
    End If

    Set Ensure_SootblowerForm_Built = vbComp
    Exit Function

EH:
    MsgBox "Ensure_SootblowerForm_Built failed: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Ensure 'Trust access to the VBA project object model' is enabled.", vbCritical
End Function

Public Sub Dev_BuildAndShow_SootblowerForm(Optional ByVal forceRebuild As Boolean = False)
    Dim comp As Object
    Set comp = Ensure_SootblowerForm_Built(forceRebuild)

    On Error Resume Next
    ' Prefer runtime binder to wire events without code injection
    Application.Run "SSB_BindAndShow", "frmSootblowerLocator"
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback: show directly (buttons won't be wired without binder)
        VBA.UserForms.Add("frmSootblowerLocator").Show
    End If
End Sub

Private Function ComponentExists(ByVal vbComps As Object, ByVal name As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = vbComps(name)
    ComponentExists = Not tmp Is Nothing
End Function

Private Sub ClearControls(ByVal designer As Object)
    On Error Resume Next
    Dim i As Long
    For i = designer.Controls.Count - 1 To 0 Step -1
        designer.Controls.Remove designer.Controls(i).Name
    Next i
End Sub

Private Sub BuildSootblowerForm(ByVal frm As Object)
    ' Form properties
    With frm
        .Caption = "Sootblower Locator"
        .Width = 420
        .Height = 240
    End With

    ' Layout constants
    Dim m As Single: m = 12
    Dim col1 As Single: col1 = m
    Dim top As Single: top = m

    ' Label: Number
    Dim lbl As Object
    Set lbl = frm.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "Sootblower Number:"
        .Left = col1
        .Top = top
        .Width = 130
        .Height = 14
        .Tag = "role:lbl_number"
    End With

    ' TextBox: Number
    Dim txt As Object
    Set txt = frm.Controls.Add("Forms.TextBox.1")
    With txt
        .Name = "txtNumber"
        .Left = col1 + 140
        .Top = top - 2
        .Width = 80
        .Height = 20
        .Tag = "role:sb_number"
    End With

    top = top + 26

    ' Label: Type
    Dim lblGroup As Object
    Set lblGroup = frm.Controls.Add("Forms.Label.1")
    With lblGroup
        .Caption = "Type:"
        .Left = col1
        .Top = top
        .Width = 40
        .Height = 14
    End With

    ' OptionButtons: All / Retracts / Wall
    Dim optAll As Object, optRetr As Object, optWall As Object
    Set optAll = frm.Controls.Add("Forms.OptionButton.1")
    With optAll
        .Name = "optAll"
        .Caption = "All"
        .Left = col1 + 50
        .Top = top - 2
        .Width = 50
        .Tag = "role:opt_all"
        .Value = True
    End With

    Set optRetr = frm.Controls.Add("Forms.OptionButton.1")
    With optRetr
        .Name = "optRetracts"
        .Caption = "Retracts (IK/EL)"
        .Left = col1 + 110
        .Top = top - 2
        .Width = 110
        .Tag = "role:opt_retracts"
    End With

    Set optWall = frm.Controls.Add("Forms.OptionButton.1")
    With optWall
        .Name = "optWall"
        .Caption = "Wall (IR/WB)"
        .Left = col1 + 230
        .Top = top - 2
        .Width = 100
        .Tag = "role:opt_wall"
    End With

    top = top + 32

    ' Buttons: Search / Show All / Associated / Close
    Dim btnSearch As Object, btnAll As Object, btnAssoc As Object, btnClose As Object
    Set btnSearch = frm.Controls.Add("Forms.CommandButton.1")
    With btnSearch
        .Name = "btnSearch"
        .Caption = "Search"
        .Left = col1
        .Top = top
        .Width = 80
        .Height = 22
        .Tag = "role:btn_search"
    End With

    Set btnAll = frm.Controls.Add("Forms.CommandButton.1")
    With btnAll
        .Name = "btnShowAll"
        .Caption = "Show All"
        .Left = col1 + 90
        .Top = top
        .Width = 80
        .Height = 22
        .Tag = "role:btn_showall"
    End With

    Set btnAssoc = frm.Controls.Add("Forms.CommandButton.1")
    With btnAssoc
        .Name = "btnAssoc"
        .Caption = "Associated"
        .Left = col1 + 180
        .Top = top
        .Width = 90
        .Height = 22
        .Tag = "role:btn_assoc"
    End With

    Set btnClose = frm.Controls.Add("Forms.CommandButton.1")
    With btnClose
        .Name = "btnClose"
        .Caption = "Close"
        .Left = col1 + 280
        .Top = top
        .Width = 70
        .Height = 22
        .Tag = "role:btn_close"
    End With

    top = top + 36

    ' Optional result/count/status labels
    Dim lblRes As Object, lblCnt As Object, lblStat As Object
    Set lblRes = frm.Controls.Add("Forms.Label.1")
    With lblRes
        .Name = "lblResults"
        .Caption = "Results:"
        .Left = col1
        .Top = top
        .Width = 60
        .Height = 14
        .Tag = "role:lbl_results"
    End With

    Set lblCnt = frm.Controls.Add("Forms.Label.1")
    With lblCnt
        .Name = "lblCount"
        .Caption = "0 items"
        .Left = col1 + 70
        .Top = top
        .Width = 90
        .Height = 14
        .Tag = "role:lbl_count"
    End With

    top = top + 20

    Set lblStat = frm.Controls.Add("Forms.Label.1")
    With lblStat
        .Name = "lblStatus"
        .Caption = "Ready"
        .Left = col1
        .Top = top
        .Width = 350
        .Height = 14
        .Tag = "role:lbl_status"
    End With
End Sub
