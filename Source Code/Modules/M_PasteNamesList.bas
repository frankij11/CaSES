Attribute VB_Name = "M_PasteNamesList"
Option Explicit

Sub M_Paste_NamesList()

Application.DisplayAlerts = False
On Error Resume Next
ActiveWorkbook.Sheets("NameList").Delete
On Error Resume Next


ActiveWorkbook.Sheets.Add.Name = "NameList"
ActiveWorkbook.Sheets("NameList").Move before:=Sheets(1)

ActiveWorkbook.Sheets("Namelist").Activate

ActiveSheet.Range("a2").Value = "Formula Name"
ActiveSheet.Range("b2").Value = "Reference"
Range("a2:b2").Select
Selection.Font.Bold = True

Columns("A:A").ColumnWidth = 21
Columns("b:B").ColumnWidth = 21
ActiveSheet.Range("a3").Select
ActiveWindow.DisplayGridlines = False
Selection.ListNames
ActiveWindow.DisplayGridlines = False
End Sub

