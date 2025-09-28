Attribute VB_Name = "temp_mod_SSB_FormStandardizer"
Option Explicit

' Temporary helper to standardize a dropped-in UserForm (default controls)
' - Renames controls to the names expected by mod_SootblowerLocator
' - Sets captions and basic layout
' - Injects event handlers that call existing module routines
' Requirements:
'   • Excel Trust Center: Enable "Trust access to the VBA project object model"
'   • No compile-time VBIDE reference needed (late-bound)
'
Private Const CT_MSForm As Long = 3

Public Sub Standardize_SSB_Form(Optional ByVal formName As String = "")
    'On Error GoTo EH
    Dim vbProj As Object, vbc As Object
    Set vbProj = ThisWorkbook.VBProject

    ' Resolve target UserForm VBComponent
    Set vbc = ResolveTargetForm(vbProj, formName)
    If vbc Is Nothing Then
        MsgBox "No UserForm found. Create a form (e.g. UserForm1) and try again.", vbExclamation, "Standardize SSB Form"
        Exit Sub
    End If

    Dim d As Object ' Designer (form surface)
    Set d = vbc.Designer
    If d Is Nothing Then
        MsgBox "Unable to access form designer. Ensure VBIDE access is allowed.", vbCritical, "Standardize SSB Form"
        Exit Sub
    End If

    Debug.Print "[SSB] Standardizing form: " & vbc.name

    ' 1) Rename controls by type and top-to-bottom order
    RenameControls d

    ' 2) Apply captions and layout
    ApplyLayoutAndCaptions d

    ' 3) Set final form name if not already
    If StrComp(vbc.name, "frmSootblowerLocator", vbTextCompare) <> 0 Then
        On Error Resume Next
        vbc.name = "frmSootblowerLocator"
        If Err.Number <> 0 Then
            Debug.Print "[SSB] Could not rename form to frmSootblowerLocator: " & Err.Number & " - " & Err.DESCRIPTION
            Err.Clear
        End If
        On Error GoTo EH
    End If

    ' 4) Inject event handlers into the form code
    InjectHandlers vbc

    MsgBox "Sootblower form standardized successfully." & vbCrLf & _
           "Form: " & vbc.name, vbInformation, "Standardize SSB Form"
    Exit Sub
EH:
    MsgBox "Error in Standardize_SSB_Form: " & Err.Number & " - " & Err.DESCRIPTION, vbCritical
End Sub

Public Sub Show_SSB_Form()
    On Error GoTo EH
    VBA.UserForms.Add("frmSootblowerLocator").Show vbModeless
    Exit Sub
EH:
    MsgBox "Show_SSB_Form failed: " & Err.Number & " - " & Err.DESCRIPTION, vbExclamation
End Sub

Private Function ResolveTargetForm(ByVal vbProj As Object, ByVal desiredName As String) As Object
    Dim vbc As Object
    If Len(Trim$(desiredName)) > 0 Then
        On Error Resume Next
        Set vbc = vbProj.VBComponents(desiredName)
        If Not vbc Is Nothing Then If vbc.Type = CT_MSForm Then Set ResolveTargetForm = vbc
        On Error GoTo 0
        If Not ResolveTargetForm Is Nothing Then Exit Function
    End If

    ' Fallback: pick first UserForm component
    For Each vbc In vbProj.VBComponents
        If vbc.Type = CT_MSForm Then Set ResolveTargetForm = vbc: Exit Function
    Next vbc
End Function

Private Sub RenameControls(ByVal d As Object)
    On Error GoTo EH
    Dim txt As Collection, opts As Collection, cmds As Collection, lbls As Collection
    Set txt = ControlsOfType(d, "TextBox")
    Set opts = ControlsOfType(d, "OptionButton")
    Set cmds = ControlsOfType(d, "CommandButton")
    Set lbls = ControlsOfType(d, "Label")

    ' Text: first textbox → txtNumber
    If txt.count >= 1 Then SafeRename txt(1), "txtNumber"

    ' Options: top-to-bottom → optAll, optRetracts, optWall
    If opts.count >= 1 Then SafeRename opts(1), "optAll"
    If opts.count >= 2 Then SafeRename opts(2), "optRetracts"
    If opts.count >= 3 Then SafeRename opts(3), "optWall"

    ' Commands: top-to-bottom → cmdSearch, cmdShowAll, cmdAssociated, cmdClose
    If cmds.count >= 1 Then SafeRename cmds(1), "cmdSearch"
    If cmds.count >= 2 Then SafeRename cmds(2), "cmdShowAll"
    If cmds.count >= 3 Then SafeRename cmds(3), "cmdAssociated"
    If cmds.count >= 4 Then SafeRename cmds(4), "cmdClose"

    ' Labels: top-to-bottom → lblResults, lblCount, lblStatus
    If lbls.count >= 1 Then SafeRename lbls(1), "lblResults"
    If lbls.count >= 2 Then SafeRename lbls(2), "lblCount"
    If lbls.count >= 3 Then SafeRename lbls(3), "lblStatus"
    Exit Sub
EH:
    Debug.Print "[SSB] RenameControls error: " & Err.Number & " - " & Err.DESCRIPTION
End Sub

Private Sub ApplyLayoutAndCaptions(ByVal d As Object)
    On Error GoTo EH
    ' Form sizing and caption
    d.caption = "Sootblower Locator"
    d.Width = 400: d.Height = 440

    ' TextBox & label defaults
    SafeSetCaption d, "lblResults", "Enter search criteria and click Search"
    SafeSetCaption d, "lblCount", "Results: 0"
    SafeSetCaption d, "lblStatus", "Ready"

    ' Options
    SafeSetCaption d, "optAll", "All Types"
    SafeSetCaption d, "optRetracts", "Retracts (IK/EL)"
    SafeSetCaption d, "optWall", "Wall (IR/WB)"
    SafeSetGroup d, Array("optAll", "optRetracts", "optWall"), "SBGroup"
    SafeSetValue d, "optAll", True

    ' Buttons
    SafeSetCaption d, "cmdSearch", "Search"
    SafeSetCaption d, "cmdShowAll", "Show All"
    SafeSetCaption d, "cmdAssociated", "Show Associated"
    SafeSetCaption d, "cmdClose", "Close"

    ' Basic arrangement (skip if any control missing)
    SafePlace d, "txtNumber", 140, 70, 80, 20
    SafePlace d, "optAll", 20, 140, 100, 20
    SafePlace d, "optRetracts", 130, 140, 120, 20
    SafePlace d, "optWall", 260, 140, 110, 20

    SafePlace d, "cmdSearch", 10, 210, 180, 30
    SafePlace d, "cmdShowAll", 210, 210, 180, 30
    SafePlace d, "cmdAssociated", 10, 250, 180, 30
    SafePlace d, "cmdClose", 210, 250, 180, 30

    SafePlace d, "lblResults", 20, 310, 360, 20
    SafePlace d, "lblCount", 20, 335, 150, 20
    SafePlace d, "lblStatus", 10, d.Height - 28, 360, 20
    Exit Sub
EH:
    Debug.Print "[SSB] ApplyLayoutAndCaptions error: " & Err.Number & " - " & Err.DESCRIPTION
End Sub

Private Sub InjectHandlers(ByVal vbc As Object)
    On Error GoTo EH
    Dim cm As Object
    Set cm = vbc.CodeModule

    ' Only add if not present
    AddProcIfMissing cm, "UserForm_Initialize", Handler_UserForm_Initialize()
    AddProcIfMissing cm, "cmdSearch_Click", Handler_cmdSearch_Click()
    AddProcIfMissing cm, "cmdShowAll_Click", Handler_cmdShowAll_Click()
    AddProcIfMissing cm, "cmdAssociated_Click", Handler_cmdAssociated_Click()
    AddProcIfMissing cm, "cmdClose_Click", Handler_cmdClose_Click()
    AddProcIfMissing cm, "SelectedGroupName", Handler_SelectedGroupName()
    Exit Sub
EH:
    MsgBox "InjectHandlers failed: " & Err.Number & " - " & Err.DESCRIPTION, vbExclamation
End Sub

Private Sub AddProcIfMissing(ByVal cm As Object, ByVal procName As String, ByVal codeText As String)
    On Error GoTo EH
    If Not ContainsText(cm, "Sub " & procName) And Not ContainsText(cm, "Function " & procName) Then
        cm.AddFromString codeText
    End If
    Exit Sub
EH:
    Debug.Print "[SSB] AddProcIfMissing(" & procName & ") error: " & Err.Number & " - " & Err.DESCRIPTION
End Sub

Private Function ContainsText(ByVal cm As Object, ByVal needle As String) As Boolean
    On Error GoTo EH
    Dim t As String
    t = cm.lines(1, cm.CountOfLines)
    ContainsText = (InStr(1, t, needle, vbTextCompare) > 0)
    Exit Function
EH:
    ContainsText = False
End Function

Private Function ControlsOfType(ByVal d As Object, ByVal typeName As String) As Collection
    Dim col As New Collection, ctl As Object
    For Each ctl In d.Controls
        If StrComp(typeName(ctl), typeName, vbTextCompare) = 0 Then col.Add ctl
    Next ctl
    ' Sort by Top ascending (simple selection sort over collection)
    Dim i As Long, j As Long
    For i = 1 To col.count - 1
        For j = i + 1 To col.count
            If col(j).Top < col(i).Top Then SwapControls col, i, j
        Next j
    Next i
    Set ControlsOfType = col
End Function

Private Sub SwapControls(ByRef col As Collection, ByVal i As Long, ByVal j As Long)
    Dim tmp As Object
    Set tmp = col(i)
    Set col(i) = col(j)
    Set col(j) = tmp
End Sub

Private Sub SafeRename(ByVal ctl As Object, ByVal newName As String)
    On Error Resume Next
    If StrComp(ctl.name, newName, vbTextCompare) <> 0 Then ctl.name = newName
    If Err.Number <> 0 Then Debug.Print "[SSB] Rename '" & ctl.name & "' → '" & newName & "' failed: " & Err.DESCRIPTION: Err.Clear
End Sub

Private Sub SafeSetCaption(ByVal d As Object, ByVal name As String, ByVal caption As String)
    On Error Resume Next
    d.Controls(name).caption = caption
    If Err.Number <> 0 Then Debug.Print "[SSB] SetCaption '" & name & "' failed: " & Err.DESCRIPTION: Err.Clear
End Sub

Private Sub SafeSetValue(ByVal d As Object, ByVal name As String, ByVal v As Variant)
    On Error Resume Next
    d.Controls(name).value = v
    If Err.Number <> 0 Then Debug.Print "[SSB] SetValue '" & name & "' failed: " & Err.DESCRIPTION: Err.Clear
End Sub

Private Sub SafeSetGroup(ByVal d As Object, ByVal names As Variant, ByVal groupName As String)
    On Error Resume Next
    Dim i As Long
    For i = LBound(names) To UBound(names)
        d.Controls(CStr(names(i))).groupName = groupName
    Next i
    If Err.Number <> 0 Then Debug.Print "[SSB] SetGroup failed: " & Err.DESCRIPTION: Err.Clear
End Sub

Private Sub SafePlace(ByVal d As Object, ByVal name As String, ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long)
    On Error Resume Next
    With d.Controls(name)
        .Left = l: .Top = t: .Width = w: .Height = h
    End With
    If Err.Number <> 0 Then Debug.Print "[SSB] Place '" & name & "' failed: " & Err.DESCRIPTION: Err.Clear
End Sub

' ===== Handlers to inject into the form code =====

Private Function Handler_UserForm_Initialize() As String
    Handler_UserForm_Initialize = _
    "Private Sub UserForm_Initialize()" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    Me.optAll.Value = True" & vbCrLf & _
    "    Me.lblStatus.Caption = ""Ready""" & vbCrLf & _
    "    Me.lblResults.Caption = ""Enter search criteria and click Search""" & vbCrLf & _
    "    Me.lblCount.Caption = ""Results: 0""" & vbCrLf & _
    "End Sub" & vbCrLf
End Function

Private Function Handler_cmdSearch_Click() As String
    Handler_cmdSearch_Click = _
    "Private Sub cmdSearch_Click()" & vbCrLf & _
    "    On Error GoTo EH" & vbCrLf & _
    "    mod_SootblowerLocator.SB_ExecuteSearch Me.txtNumber.Text, SelectedGroupName()" & vbCrLf & _
    "    Me.lblStatus.Caption = ""Search completed""" & vbCrLf & _
    "    Exit Sub" & vbCrLf & _
    "EH:" & vbCrLf & _
    "    Me.lblStatus.Caption = ""Search error: "" & Err.Number" & vbCrLf & _
    "End Sub" & vbCrLf
End Function

Private Function Handler_cmdShowAll_Click() As String
    Handler_cmdShowAll_Click = _
    "Private Sub cmdShowAll_Click()" & vbCrLf & _
    "    On Error GoTo EH" & vbCrLf & _
    "    mod_SootblowerLocator.SB_DisplayAll SelectedGroupName()" & vbCrLf & _
    "    Me.lblStatus.Caption = ""Showing all""" & vbCrLf & _
    "    Exit Sub" & vbCrLf & _
    "EH:" & vbCrLf & _
    "    Me.lblStatus.Caption = ""ShowAll error: "" & Err.Number" & vbCrLf & _
    "End Sub" & vbCrLf
End Function

Private Function Handler_cmdAssociated_Click() As String
    Handler_cmdAssociated_Click = _
    "Private Sub cmdAssociated_Click()" & vbCrLf & _
    "    On Error GoTo EH" & vbCrLf & _
    "    mod_SootblowerLocator.SB_ShowAssociated Me.txtNumber.Text, SelectedGroupName()" & vbCrLf & _
    "    Me.lblStatus.Caption = ""Associated updated""" & vbCrLf & _
    "    Exit Sub" & vbCrLf & _
    "EH:" & vbCrLf & _
    "    Me.lblStatus.Caption = ""Associated error: "" & Err.Number" & vbCrLf & _
    "End Sub" & vbCrLf
End Function

Private Function Handler_cmdClose_Click() As String
    Handler_cmdClose_Click = _
    "Private Sub cmdClose_Click()" & vbCrLf & _
    "    Unload Me" & vbCrLf & _
    "End Sub" & vbCrLf
End Function

Private Function Handler_SelectedGroupName() As String
    Handler_SelectedGroupName = _
    "Private Function SelectedGroupName() As String" & vbCrLf & _
    "    If Me.optRetracts.Value Then" & vbCrLf & _
    "        SelectedGroupName = ""Retracts""" & vbCrLf & _
    "    ElseIf Me.optWall.Value Then" & vbCrLf & _
    "        SelectedGroupName = ""Wall""" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        SelectedGroupName = """"" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Function" & vbCrLf
End Function
