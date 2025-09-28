Attribute VB_Name = "Module1"
Option Explicit

' Required reference: Microsoft Visual Basic for Applications Extensibility 5.3
' In the VBE: Tools -> References -> Check the box for this library.

' ---
' This is the main procedure you would run from Excel.
' It builds a UserForm from scratch, adds controls, injects code, and shows it.
' ---
Public Sub CreateAndShowDynamicForm()
    ' --- PART 1: Variable Declaration ---
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    ' --- DIAGNOSTIC CHANGE: Using 'As Object' for late binding ---
    ' Original line was: Dim formDesigner As VBIDE.Designer
    Dim formDesigner As Object
    
    Dim newButton As Object ' MSForms.CommandButton
    Dim newLabel As Object  ' MSForms.Label
    Dim newTextBox As Object ' MSForms.TextBox
    Dim codeMod As VBIDE.CodeModule
    Dim lineNum As Long
    
    ' Define the name for our new form
    Const NEW_FORM_NAME As String = "MyDynamicForm"

    ' --- PART 2: Setup and Cleanup ---
    ' Set a reference to the current VBA project
    Set vbProj = ThisWorkbook.VBProject
    
    ' Check if a form with this name already exists and delete it.
    ' This makes the builder macro re-runnable.
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(NEW_FORM_NAME)
    On Error GoTo 0

    ' --- PART 3: Create the Form ---
    ' Add a new UserForm component to the project
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    
    ' Set the form's properties
    With vbComp
        .name = NEW_FORM_NAME
        .Properties("Caption") = "Generated Form"
        .Properties("Width") = 240
        .Properties("Height") = 180
    End With
    
    Set formDesigner = vbComp.Designer

    ' --- PART 4: Add Controls to the Form ---
    ' Add a Label
    Set newLabel = formDesigner.Controls.Add("Forms.Label.1")
    With newLabel
        .caption = "Please enter your name:"
        .Left = 12
        .Top = 12
        .Width = 150
        .Height = 12
    End With
    
    ' Add a TextBox
    Set newTextBox = formDesigner.Controls.Add("Forms.TextBox.1")
    With newTextBox
        .name = "NameTextBox"
        .Left = 12
        .Top = 30
        .Width = 200
        .Height = 20
        .text = ""
    End With

    ' Add a CommandButton
    Set newButton = formDesigner.Controls.Add("Forms.CommandButton.1")
    With newButton
        .name = "SubmitButton"
        .caption = "Submit"
        .Left = 132
        .Top = 60
        .Width = 80
        .Height = 25
    End With

    ' --- PART 5: Inject Event Handler Code ---
    ' Get a reference to the form's code module
    Set codeMod = vbComp.CodeModule
    
    ' Insert the code for the button's click event
    With codeMod
        ' Find the last line in the declarations section to start writing our sub
        lineNum = .CountOfDeclarationLines + 1
        
        ' Write the subroutine
        .InsertLines lineNum, "Private Sub SubmitButton_Click()"
        .InsertLines lineNum + 1, "    ' This code was injected programmatically!"
        .InsertLines lineNum + 2, "    MsgBox ""Hello, "" & Me.NameTextBox.Text & ""!"""
        .InsertLines lineNum + 3, "    Unload Me"
        .InsertLines lineNum + 4, "End Sub"
    End With

    ' --- PART 6: Show the Form ---
    VBA.UserForms.Add(NEW_FORM_NAME).Show

    ' --- PART 7: Cleanup (Optional) ---
    ' You might want to remove the form after it's used if it's temporary.
    ' For this example, we'll leave it in the project.
    ' vbProj.VBComponents.Remove vbComp

    ' Clean up object variables
    Set newButton = Nothing
    Set newLabel = Nothing
    Set newTextBox = Nothing
    Set codeMod = Nothing
    Set vbComp = Nothing
    Set vbProj = Nothing
    
    MsgBox "Form generation complete."

End Sub


