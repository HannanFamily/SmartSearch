Attribute VB_Name = "mod_SSB_RuntimeBinder"
Option Explicit

' Runtime binder for dropped-in UserForms without VBIDE renaming/code-injection.
' Usage:
'   1) Drop a UserForm (any name, e.g., UserForm1) with:
'      - 1 TextBox (number input)  [optional Tag: role:number]
'      - 3 OptionButtons (All, Retracts, Wall) [optional Tags: role:opt_all, role:opt_retracts, role:opt_wall]
'      - 4 CommandButtons (Search, Show All, Show Associated, Close) [optional Tags: role:btn_search, role:btn_showall, role:btn_assoc, role:btn_close]
'      - 3 Labels (Results, Count, Status) [optional Tags: role:lbl_results, role:lbl_count, role:lbl_status]
'   2) Run: SSB_BindAndShow "UserForm1"
'
' No Trust Center VBIDE access is required; we do not manipulate code or component names.

Private Type TBinding
    frm As Object
    hSearch As C_SSB_BtnHandler
    hShowAll As C_SSB_BtnHandler
    hAssoc As C_SSB_BtnHandler
    hClose As C_SSB_BtnHandler
End Type

Private mBindings As Collection ' holds TBinding items to keep WithEvents alive

Public Sub SSB_BindAndShow(Optional ByVal formName As String = "UserForm1")
    On Error GoTo EH
    Dim frm As Object
    Set frm = VBA.UserForms.Add(formName)
    If frm Is Nothing Then
        MsgBox "Could not create form '" & formName & "'", vbExclamation
        Exit Sub
    End If

    ' Initialize captions/defaults
    SafeSetCaption frm, FindLabel(frm, "role:lbl_results", 1), "Enter search criteria and click Search"
    SafeSetCaption frm, FindLabel(frm, "role:lbl_count", 2), "Results: 0"
    SafeSetCaption frm, FindLabel(frm, "role:lbl_status", 3), "Ready"

    ' Prefer tagged options, else first three by order. Set All=True by default.
    Dim optAll As Object, optRetr As Object, optWall As Object
    Set optAll = FindOption(frm, "role:opt_all", 1)
    Set optRetr = FindOption(frm, "role:opt_retracts", 2)
    Set optWall = FindOption(frm, "role:opt_wall", 3)
    SafeSetValue optAll, True

    ' Bind buttons via WithEvents handler instances
    Dim searchBtn As Object, showAllBtn As Object, assocBtn As Object, closeBtn As Object
    Set searchBtn = FindButton(frm, "role:btn_search", 1)
    Set showAllBtn = FindButton(frm, "role:btn_showall", 2)
    Set assocBtn = FindButton(frm, "role:btn_assoc", 3)
    Set closeBtn = FindButton(frm, "role:btn_close", 4)

    Dim b As TBinding
    Set b.frm = frm
    Set b.hSearch = New C_SSB_BtnHandler: Set b.hSearch.Btn = searchBtn: Set b.hSearch.ParentForm = frm: b.hSearch.Role = "search"
    Set b.hShowAll = New C_SSB_BtnHandler: Set b.hShowAll.Btn = showAllBtn: Set b.hShowAll.ParentForm = frm: b.hShowAll.Role = "showall"
    Set b.hAssoc = New C_SSB_BtnHandler: Set b.hAssoc.Btn = assocBtn: Set b.hAssoc.ParentForm = frm: b.hAssoc.Role = "assoc"
    Set b.hClose = New C_SSB_BtnHandler: Set b.hClose.Btn = closeBtn: Set b.hClose.ParentForm = frm: b.hClose.Role = "close"

    If mBindings Is Nothing Then Set mBindings = New Collection
    mBindings.Add b  ' keep handlers alive

    frm.Caption = "Sootblower Locator"
    frm.Show vbModeless
    Exit Sub
EH:
    MsgBox "SSB_BindAndShow failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ----- Helpers to find controls (prefer Tag, else Nth by type order) -----

Private Function FindByTag(ByVal frm As Object, ByVal typeName As String, ByVal tag As String) As Object
    Dim ctl As Object
    For Each ctl In frm.Controls
        If TypeName(ctl) = typeName Then
            If LCase$(Trim$(ctl.Tag)) = LCase$(tag) Then Set FindByTag = ctl: Exit Function
        End If
    Next ctl
End Function

Private Function NthByType(ByVal frm As Object, ByVal typeName As String, ByVal n As Long) As Object
    Dim ctl As Object, k As Long
    For Each ctl In frm.Controls
        If TypeName(ctl) = typeName Then
            k = k + 1
            If k = n Then Set NthByType = ctl: Exit Function
        End If
    Next ctl
End Function

Private Function FindButton(ByVal frm As Object, ByVal tag As String, ByVal n As Long) As Object
    Set FindButton = FindByTag(frm, "CommandButton", tag)
    If FindButton Is Nothing Then Set FindButton = NthByType(frm, "CommandButton", n)
End Function

Private Function FindOption(ByVal frm As Object, ByVal tag As String, ByVal n As Long) As Object
    Set FindOption = FindByTag(frm, "OptionButton", tag)
    If FindOption Is Nothing Then Set FindOption = NthByType(frm, "OptionButton", n)
End Function

Private Function FindLabel(ByVal frm As Object, ByVal tag As String, ByVal n As Long) As Object
    Set FindLabel = FindByTag(frm, "Label", tag)
    If FindLabel Is Nothing Then Set FindLabel = NthByType(frm, "Label", n)
End Function

Private Sub SafeSetCaption(ByVal frm As Object, ByVal ctl As Object, ByVal cap As String)
    On Error Resume Next
    If Not ctl Is Nothing Then ctl.Caption = cap
End Sub

Private Sub SafeSetValue(ByVal ctl As Object, ByVal v As Variant)
    On Error Resume Next
    If Not ctl Is Nothing Then ctl.Value = v
End Sub
