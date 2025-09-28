VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SSB_FormEvents"
Option Explicit

Private WithEvents btnSearch As MSForms.CommandButton
Private WithEvents btnShowAll As MSForms.CommandButton
Private WithEvents btnAssociated As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private mForm As Object

Public Sub Bind(ByVal frm As Object)
    Set mForm = frm
    On Error Resume Next
    Set btnSearch = frm.Controls("cmdSearch")
    Set btnShowAll = frm.Controls("cmdShowAll")
    Set btnAssociated = frm.Controls("cmdAssociated")
    Set btnClose = frm.Controls("cmdClose")
    On Error GoTo 0
End Sub

Private Sub btnSearch_Click()
    On Error Resume Next
    Dim numTxt As String
    numTxt = mForm.Controls("txtNumber").Value
    Call mod_SootblowerLocator.SB_ExecuteSearch(numTxt, SSB_GetSelectedGroup(mForm))
End Sub

Private Sub btnShowAll_Click()
    On Error Resume Next
    Call mod_SootblowerLocator.SB_DisplayAll(SSB_GetSelectedGroup(mForm))
End Sub

Private Sub btnAssociated_Click()
    On Error Resume Next
    Call mod_SootblowerLocator.SB_ShowAssociated
End Sub

Private Sub btnClose_Click()
    On Error Resume Next
    Unload mForm
End Sub

Public Function SSB_GetSelectedGroup(ByVal frm As Object) As String
    On Error Resume Next
    If frm.Controls("optRetracts").Value Then
        SSB_GetSelectedGroup = "Retracts"
    ElseIf frm.Controls("optWall").Value Then
        SSB_GetSelectedGroup = "Wall"
    Else
        SSB_GetSelectedGroup = ""
    End If
End Function
