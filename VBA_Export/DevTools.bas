'Attribute VB_Name = "DevTools"  ' commented for copy/paste

'============================================================
' UTILITIES (misc dev helpers)
'============================================================
Public Sub InsertStaticDateTime()
    With ActiveCell
        .Value = Now
        .NumberFormat = "mm/dd/yyyy hh:mm"
    End With
End Sub

