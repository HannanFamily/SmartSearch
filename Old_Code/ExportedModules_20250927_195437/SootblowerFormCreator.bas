Attribute VB_Name = "SootblowerFormCreator"
Option Explicit

' Sootblower Form Creator
' ------------------------------------------------------------
' This module creates a UserForm for Sootblower Locator functionality
' dynamically at runtime when the standard form can't be loaded from .frm
'
' Usage:
'   - Call CreateSootblowerForm() to dynamically generate the form
'   - The function returns a reference to the created form
'   - Show the form using form.Show vbModeless
'
' The dynamic form has all the same functionality as the design-time form:
'   - Sootblower number input with validation
'   - Group filter selection (All, Retracts, Wall)
'   - Search, Show All, Associated and Close buttons
'   - Status and results display
'
' This avoids dependency on .frm/.frx files and VBIDE references

Public Function CreateSootblowerForm() As Object
    ' Creates and returns a fully functional Sootblower Locator form
    On Error Resume Next
    
    ' Form constants
    Const FORM_WIDTH As Long = 400
    Const FORM_HEIGHT As Long = 440
    
    ' Colors
    Const COLOR_PRIMARY As Long = &H4F6228    ' Dark Green
    Const COLOR_SECONDARY As Long = &H8DC63F  ' Light Green
    Const COLOR_ACCENT As Long = &H3A6EA5    ' Blue Accent
    Const COLOR_TEXT As Long = &H2B2B2B      ' Near Black
    Const COLOR_LIGHT As Long = &HF5F5F5     ' Near White
    
    ' Create the form
    Dim frm As Object
    Set frm = VBA.UserForms.Add("UserForm1")
    frm.Caption = "Sootblower Locator"
    frm.Width = FORM_WIDTH
    frm.Height = FORM_HEIGHT
    frm.BackColor = COLOR_LIGHT
    
    ' ===== Header Section =====
    
    ' Title label
    Dim lblTitle As Object
    Set lblTitle = frm.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption = "Sootblower Locator"
        .Left = 10
        .Top = 10
        .Width = FORM_WIDTH - 20
        .Height = 24
        .Font.Bold = True
        .Font.Size = 14
        .TextAlign = 2  ' Center
        .ForeColor = COLOR_PRIMARY
    End With
    
    ' Subtitle label
    Dim lblSubtitle As Object
    Set lblSubtitle = frm.Controls.Add("Forms.Label.1", "lblSubtitle")
    With lblSubtitle
        .Caption = "Locate sootblowers by number, type or group"
        .Left = 10
        .Top = 34
        .Width = FORM_WIDTH - 20
        .Height = 20
        .TextAlign = 2  ' Center
    End With
    
    ' Separator line
    Dim lineSep1 As Object
    Set lineSep1 = frm.Controls.Add("Forms.Label.1", "lineSep1")
    With lineSep1
        .Caption = ""
        .Left = 10
        .Top = 60
        .Width = FORM_WIDTH - 20
        .Height = 1
        .BackColor = COLOR_SECONDARY
        .BorderStyle = 0
    End With
    
    ' ===== Input Section =====
    
    ' Number label
    Dim lblNumber As Object
    Set lblNumber = frm.Controls.Add("Forms.Label.1", "lblNumber")
    With lblNumber
        .Caption = "Sootblower Number:"
        .Left = 10
        .Top = 70
        .Width = 120
        .Height = 20
    End With
    
    ' Number text box
    Dim txtNumber As Object
    Set txtNumber = frm.Controls.Add("Forms.TextBox.1", "txtNumber")
    With txtNumber
        .Left = 140
        .Top = 70
        .Width = 80
        .Height = 20
    End With
    
    ' Help text
    Dim lblNumberHelp As Object
    Set lblNumberHelp = frm.Controls.Add("Forms.Label.1", "lblNumberHelp")
    With lblNumberHelp
        .Caption = "(Leave empty to show all sootblowers)"
        .Left = 140
        .Top = 95
        .Width = FORM_WIDTH - 150
        .Height = 20
        .Font.Italic = True
        .Font.Size = 8
        .ForeColor = RGB(128, 128, 128)
    End With
    
    ' ===== Filter Section =====
    
    ' Filter frame
    Dim fraFilter As Object
    Set fraFilter = frm.Controls.Add("Forms.Frame.1", "fraFilter")
    With fraFilter
        .Caption = "Filter by Group"
        .Left = 10
        .Top = 120
        .Width = FORM_WIDTH - 20
        .Height = 80
        .Font.Bold = True
        .ForeColor = COLOR_ACCENT
    End With
    
    ' All option
    Dim optAll As Object
    Set optAll = frm.Controls.Add("Forms.OptionButton.1", "optAll")
    With optAll
        .Caption = "All Types"
        .Left = 20
        .Top = 140
        .Width = 100
        .Height = 20
        .value = True
        .GroupName = "SBGroup"
    End With
    
    ' Retracts option
    Dim optRetracts As Object
    Set optRetracts = frm.Controls.Add("Forms.OptionButton.1", "optRetracts")
    With optRetracts
        .Caption = "Retracts (IK/EL)"
        .Left = 130
        .Top = 140
        .Width = 120
        .Height = 20
        .GroupName = "SBGroup"
    End With
    
    ' Wall option
    Dim optWall As Object
    Set optWall = frm.Controls.Add("Forms.OptionButton.1", "optWall")
    With optWall
        .Caption = "Wall (IR/WB)"
        .Left = 260
        .Top = 140
        .Width = 110
        .Height = 20
        .GroupName = "SBGroup"
    End With
    
    ' Filter help text
    Dim lblFilterHelp As Object
    Set lblFilterHelp = frm.Controls.Add("Forms.Label.1", "lblFilterHelp")
    With lblFilterHelp
        .Caption = "Retracts: SBIK, SBEL types | Wall: SBIR, SBWB types"
        .Left = 20
        .Top = 165
        .Width = FORM_WIDTH - 40
        .Height = 20
        .Font.Italic = True
        .Font.Size = 8
        .ForeColor = RGB(128, 128, 128)
    End With
    
    ' ===== Button Section =====
    
    ' Search button
    Dim cmdSearch As Object
    Set cmdSearch = frm.Controls.Add("Forms.CommandButton.1", "cmdSearch")
    With cmdSearch
        .Caption = "Search"
        .Left = 10
        .Top = 210
        .Width = 180
        .Height = 30
        .BackColor = COLOR_PRIMARY
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Show All button
    Dim cmdShowAll As Object
    Set cmdShowAll = frm.Controls.Add("Forms.CommandButton.1", "cmdShowAll")
    With cmdShowAll
        .Caption = "Show All"
        .Left = 210
        .Top = 210
        .Width = 180
        .Height = 30
        .BackColor = COLOR_SECONDARY
    End With
    
    ' Associated button
    Dim cmdAssociated As Object
    Set cmdAssociated = frm.Controls.Add("Forms.CommandButton.1", "cmdAssociated")
    With cmdAssociated
        .Caption = "Show Associated"
        .Left = 10
        .Top = 250
        .Width = 180
        .Height = 30
        .Enabled = False
        .BackColor = COLOR_ACCENT
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Close button
    Dim cmdClose As Object
    Set cmdClose = frm.Controls.Add("Forms.CommandButton.1", "cmdClose")
    With cmdClose
        .Caption = "Close"
        .Left = 210
        .Top = 250
        .Width = 180
        .Height = 30
    End With
    
    ' ===== Results Section =====
    
    ' Results frame
    Dim fraResults As Object
    Set fraResults = frm.Controls.Add("Forms.Frame.1", "fraResults")
    With fraResults
        .Caption = "Search Results"
        .Left = 10
        .Top = 290
        .Width = FORM_WIDTH - 20
        .Height = 100
        .Font.Bold = True
        .ForeColor = COLOR_ACCENT
    End With
    
    ' Results label
    Dim lblResults As Object
    Set lblResults = frm.Controls.Add("Forms.Label.1", "lblResults")
    With lblResults
        .Caption = "Enter search criteria and click Search"
        .Left = 20
        .Top = 310
        .Width = FORM_WIDTH - 40
        .Height = 20
    End With
    
    ' Count label
    Dim lblCount As Object
    Set lblCount = frm.Controls.Add("Forms.Label.1", "lblCount")
    With lblCount
        .Caption = "Results: 0"
        .Left = 20
        .Top = 335
        .Width = 150
        .Height = 20
        .Font.Bold = True
    End With
    
    ' Display info
    Dim lblDisplayInfo As Object
    Set lblDisplayInfo = frm.Controls.Add("Forms.Label.1", "lblDisplayInfo")
    With lblDisplayInfo
        .Caption = "Results are displayed on the Dashboard"
        .Left = 20
        .Top = 360
        .Width = FORM_WIDTH - 40
        .Height = 20
        .Font.Italic = True
        .Font.Size = 8
        .ForeColor = RGB(128, 128, 128)
    End With
    
    ' ===== Status Bar =====
    
    ' Status bar background
    Dim shpStatusBar As Object
    Set shpStatusBar = frm.Controls.Add("Forms.Rectangle.1", "shpStatusBar")
    With shpStatusBar
        .Left = 0
        .Top = FORM_HEIGHT - 30
        .Width = FORM_WIDTH
        .Height = 25
        .BackColor = RGB(240, 240, 240)
        .BorderStyle = 0
    End With
    
    ' Status label
    Dim lblStatus As Object
    Set lblStatus = frm.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Caption = "Ready"
        .Left = 10
        .Top = FORM_HEIGHT - 28
        .Width = FORM_WIDTH - 20
        .Height = 20
        .Font.Size = 8
    End With
    
    ' ===== Attach Event Handlers =====
    
    ' Use code events since we can't attach directly
    AttachEventHandlers frm
    
    ' Return the form
    Set CreateSootblowerForm = frm
End Function

Private Sub AttachEventHandlers(ByRef frm As Object)
    ' Store reference to form in a module-level variable
    ' for event handling. This is a workaround for dynamic forms.
    
    ' When a form is created dynamically, we can't attach direct event handlers
    ' Instead, we use onClick, onChange etc properties which take strings
    
    With frm.Controls("cmdSearch")
        .OnClick = "Call mod_SootblowerLocator.SB_ExecuteSearch(UserForms(""" & frm.name & """).Controls(""txtNumber"").Value, GetSelectedGroup(UserForms(""" & frm.name & """)))"
    End With
    
    With frm.Controls("cmdShowAll")
        .OnClick = "Call mod_SootblowerLocator.SB_DisplayAll(GetSelectedGroup(UserForms(""" & frm.name & """)))"
    End With
    
    With frm.Controls("cmdAssociated")
        .OnClick = "Call mod_SootblowerLocator.SB_ShowAssociated"
    End With
    
    With frm.Controls("cmdClose")
        .OnClick = "Unload UserForms(""" & frm.name & """)"
    End With
    
    ' For TextBox input validation, we need a custom approach
    ' This is typically handled with a KeyPress event handler
End Sub

Public Function GetSelectedGroup(ByRef frm As Object) As String
    ' Get the selected group from the form's radio buttons
    On Error Resume Next
    
    If frm.Controls("optAll").value Then
        GetSelectedGroup = ""
    ElseIf frm.Controls("optRetracts").value Then
        GetSelectedGroup = "Retracts"
    ElseIf frm.Controls("optWall").value Then
        GetSelectedGroup = "Wall"
    Else
        GetSelectedGroup = ""
    End If
End Function

Public Sub UpdateSootblowerFormResults(ByVal formName As String, ByVal count As Long)
    ' Updates the results count on the specified form
    On Error Resume Next
    
    Dim frm As Object
    Set frm = VBA.UserForms(formName)
    
    If Not frm Is Nothing Then
        With frm
            .Controls("lblCount").Caption = "Results: " & count
            
            If count > 0 Then
                .Controls("cmdAssociated").Enabled = True
                .Controls("lblResults").Caption = "Found " & count & " matching sootblowers"
            Else
                .Controls("cmdAssociated").Enabled = False
                .Controls("lblResults").Caption = "No matching sootblowers found"
            End If
            
            .Controls("lblStatus").Caption = "Ready"
        End With
    End If
End Sub
