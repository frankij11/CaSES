Attribute VB_Name = "M_DeleteActive_Array"
Option Explicit
Sub M_DeleteActiveArray()
    On Error Resume Next
    Selection.CurrentArray.Select
    On Error Resume Next
    Selection.ClearContents
    On Error GoTo 0
End Sub


