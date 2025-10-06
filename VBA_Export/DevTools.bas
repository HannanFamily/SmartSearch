Attribute VB_Name = "DevTools"

'============================================================
' UTILITIES (misc dev helpers)
'============================================================
Public Sub InsertStaticDateTime()
    With ActiveCell
        .Value = Now
        .NumberFormat = "mm/dd/yyyy hh:mm"
    End With
End Sub

