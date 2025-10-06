Option Explicit
'
' Compatibility shim module.
' This forwards legacy procedure names expected by existing sheet/workbook code
' to the new implementations in mod_SearchEngineCore. Once all references are
' updated to call mod_SearchEngineCore directly, this module can be deleted.
'
Public Sub RefreshResults_Enhanced()
	On Error Resume Next
	mod_SearchEngineCore.RefreshResults_Enhanced
End Sub

Public Sub RefreshResults()
	' Legacy name (unqualified calls like RefreshResults)
	On Error Resume Next
	mod_SearchEngineCore.RefreshResults_Enhanced
End Sub

Public Sub EnsurePulseCell(Optional ByVal resetValue As Boolean = False)
	On Error Resume Next
	mod_SearchEngineCore.EnsurePulseCell resetValue
End Sub

Public Sub ClearTempSearchFilter()
	On Error Resume Next
	mod_SearchEngineCore.ClearTempSearchFilter
End Sub

Public Function AllInputNamedRanges_Enhanced() As Collection
	On Error Resume Next
	Set AllInputNamedRanges_Enhanced = mod_SearchEngineCore.AllInputNamedRanges_Enhanced
End Function

Public Function IsAnySearchInputActive() As Boolean
	On Error Resume Next
	IsAnySearchInputActive = mod_SearchEngineCore.IsAnySearchInputActive
End Function
