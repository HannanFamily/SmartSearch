Attribute VB_Name = "SootblowerFormFactory"
Option Explicit

' DISABLED: This module requires VBIDE references and causes import errors
' Use manual form creation or skip this feature

Public Function CreateSootblowerUserForm() As Boolean
    MsgBox "SootblowerFormFactory is disabled to prevent import errors." & vbCrLf & _
           "Please create the UserForm manually or skip this feature.", vbInformation
    CreateSootblowerUserForm = False
End Function
