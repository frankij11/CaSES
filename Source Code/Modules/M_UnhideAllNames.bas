Attribute VB_Name = "M_UnhideAllNames"
Option Explicit

Sub M_Unhide_AllNames()
 Dim N As Name
 For Each N In ThisWorkbook.Names
 If N.Visible = False Then N.Visible = True
 Next N
End Sub
