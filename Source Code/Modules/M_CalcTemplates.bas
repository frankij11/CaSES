Attribute VB_Name = "M_CalcTemplates"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'This procedure copies inflation indices to workbook
Public Sub addGenericCalc(Optional wbsName)
If IsMissing(wbsName) Then wbsName = "Calculation Template"
Dim location As Range
Dim i As Integer
Application.ScreenUpdating = False

Set location = Selection
'On Error GoTo ERR:
Dim newCalc As Worksheet
Set newCalc = ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.count))
ThisWorkbook.Worksheets("Calculation Template").Cells.Copy

newCalc.Range("A1").PasteSpecial (xlPasteAll) ' after:=ActiveWorkbook.ActiveSheet
newCalc.Range("A1").Select
On Error Resume Next
newCalc.Name = wbsName
For i = 1 To 20
    If Len(Replace(newCalc.Name, wbsName, "")) - Len(newCalc.Name) <> 0 Then Exit For
    newCalc.Name = wbsName & " (" & i & ")"
Next

With newCalc.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
End With

With ActiveWindow
    .DisplayGridlines = False
    
    .SplitColumn = 0
    .SplitRow = 14
    .FreezePanes = True
End With
Debug.Print location.Address
'Location.Select
Application.ScreenUpdating = True

Exit Sub
ERR:
'Location.Select
Application.ScreenUpdating = True

End Sub
