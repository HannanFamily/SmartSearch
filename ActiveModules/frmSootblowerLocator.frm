VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSootblowerLocator 
   Caption         =   "Sootblower Locator"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "frmSootblowerLocator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSootblowerLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Sootblower Locator Form
' ---------------------------------------------------------
' This UserForm provides a GUI for locating and filtering sootblowers
' within the equipment database.
'
' Controls:
'   - txtNumber: Input for sootblower number
'   - optAll/optRetracts/optWall: Radio buttons for type filtering
'   - cmdSearch: Execute search with current parameters
'   - cmdShowAll: Show all matching the selected filter
'   - cmdAssociated: Show equipment associated with current results
'
' Integration:
'   - Called from mod_SootblowerLocator.Init_SootblowerLocator()
'   - Results displayed on Dashboard worksheet
'   - Calls back to mod_SootblowerLocator for search logic
'
' Error Handling:
'   - Input validation on sootblower number (digits only)
'   - Status updates during operations
'   - Graceful error recovery

Private Sub UserForm_Initialize()
    ' Initialize form with default values
    lblStatus.Caption = "Ready"
    optAll.Value = True
    txtNumber.SetFocus
    
    ' Start with Associated button disabled until results are found
    cmdAssociated.Enabled = False
    
    ' Ensure borders and styling are consistent
    Me.BackColor = RGB(245, 245, 245)  ' Light background
End Sub

Private Sub cmdSearch_Click()
    ' Execute search with current parameters
    Dim numberText As String, groupSel As String
    
    numberText = Trim(txtNumber.Value)
    
    ' Get selected group filter
    If optAll.Value Then
        groupSel = ""
    ElseIf optRetracts.Value Then
        groupSel = "Retracts"
    ElseIf optWall.Value Then
        groupSel = "Wall"
    End If
    
    ' Update status
    lblStatus.Caption = "Searching..."
    lblResults.Caption = "Searching for sootblowers..."
    Me.Repaint
    
    ' Call search function in main module
    On Error Resume Next
    Call mod_SootblowerLocator.SB_ExecuteSearch(numberText, groupSel)
    
    If Err.Number <> 0 Then
        lblStatus.Caption = "Error: " & Err.Description
        MsgBox "Search failed: " & Err.Description, vbExclamation, "Search Error"
        Err.Clear
    End If
End Sub

Private Sub cmdShowAll_Click()
    ' Show all sootblowers matching the selected filter
    Dim groupSel As String
    
    ' Get selected group filter
    If optAll.Value Then
        groupSel = ""
    ElseIf optRetracts.Value Then
        groupSel = "Retracts"
    ElseIf optWall.Value Then
        groupSel = "Wall"
    End If
    
    ' Update status
    lblStatus.Caption = "Loading all sootblowers..."
    lblResults.Caption = "Retrieving all matching sootblowers..."
    Me.Repaint
    
    ' Call show all function
    On Error Resume Next
    Call mod_SootblowerLocator.SB_DisplayAll(groupSel)
    
    If Err.Number <> 0 Then
        lblStatus.Caption = "Error: " & Err.Description
        MsgBox "Show All failed: " & Err.Description, vbExclamation, "Search Error"
        Err.Clear
    End If
End Sub

Private Sub cmdAssociated_Click()
    ' Show equipment associated with current results
    lblStatus.Caption = "Showing associated sootblowers..."
    lblResults.Caption = "Finding associated equipment..."
    Me.Repaint
    
    ' Call associated function
    On Error Resume Next
    Call mod_SootblowerLocator.SB_ShowAssociated
    
    If Err.Number <> 0 Then
        lblStatus.Caption = "Error: " & Err.Description
        MsgBox "Show Associated failed: " & Err.Description, vbExclamation, "Search Error"
        Err.Clear
    End If
End Sub

Private Sub cmdClose_Click()
    ' Close form
    Unload Me
End Sub

Private Sub txtNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow only digits and control characters (backspace, enter)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Public Sub UpdateResultsCount(ByVal count As Long)
    ' Update the results count and status
    lblCount.Caption = "Results: " & count
    
    If count > 0 Then
        cmdAssociated.Enabled = True
        lblResults.Caption = "Found " & count & " matching sootblowers"
    Else
        cmdAssociated.Enabled = False
        lblResults.Caption = "No matching sootblowers found"
    End If
    
    lblStatus.Caption = "Ready"
End Sub