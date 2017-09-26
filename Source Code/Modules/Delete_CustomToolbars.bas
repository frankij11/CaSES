Attribute VB_Name = "Delete_CustomToolbars"
Option Explicit
Sub deletecustomtoolbars()
Dim Bar
For Each Bar In Application.CommandBars
    If Not Bar.BuiltIn And Not Bar.Visible Then Bar.Delete
Next

End Sub

