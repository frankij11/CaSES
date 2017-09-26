Attribute VB_Name = "M_Break_Links"
Option Explicit
Sub M_BreakLinks()
Dim link
Dim wb As Workbook
Set wb = Application.ActiveWorkbook
If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
    For Each link In wb.LinkSources(xlExcelLinks)
        wb.BreakLink link, xlLinkTypeExcelLinks
    Next link
End If
End Sub

