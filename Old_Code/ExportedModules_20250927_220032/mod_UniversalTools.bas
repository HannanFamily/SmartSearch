Attribute VB_Name = "mod_UniversalTools"
Public Sub InsertStaticDateTime()
    ' Inserts current date and time as a static value into the active cell
    With ActiveCell
        .value = Now
        .NumberFormat = "mm/dd/yyyy hh:mm"
    End With
End Sub

