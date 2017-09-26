Attribute VB_Name = "M_1_ModelTemplate"
Option Explicit
Sub OpenModel()
'get Add in path location
Dim pth As Variant
pth = ThisWorkbook.Path
'add File Name Cost Model to path
pth = pth & "\Supporting Files\ModelTemplate_2017.xltm"

'Open workbook
Workbooks.Open pth

End Sub

Sub OpenUncertainty()
Dim pth As Variant
Dim tmp As Variant
Application.ScreenUpdating = False
'Application.DisplayAlerts = False
'get Add in path location
pth = ThisWorkbook.Path
'add File Name Cost Model to path
pth = pth & "\Supporting Files\UncertaintyTemplate.xltm"

'Open workbook

Debug.Print pth
Set tmp = Workbooks.Open(pth)
'tmp.Worksheets.Copy
'tmp.Close

'Application.DisplayAlerts = True
Application.ScreenUpdating = False

End Sub

Sub Open_JACSRUH()
Dim pth As Variant
Application.ScreenUpdating = False
pth = ThisWorkbook.Path
pth = pth & "\Supporting Files\JA CSRUH_Example File_ModelTemplate.xltm"
Workbooks.Open pth
End Sub
