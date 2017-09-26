Attribute VB_Name = "M_DeleteCustomRibbon"
Option Explicit
Sub CleanCommandBar()

Const cCommandBar = "CaSES"
Const cCommandBar2 = "CMR Tools 2"
Const cCommandBar3 = "CMR Tools 3"

Application.CommandBars(cCommandBar).Delete
Application.CommandBars(cCommandBar2).Delete
Application.CommandBars(cCommandBar3).Delete

End Sub
