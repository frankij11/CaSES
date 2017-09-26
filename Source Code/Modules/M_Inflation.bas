Attribute VB_Name = "M_Inflation"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'This procedure copies inflation indices to workbook
Sub copyInflation()
Dim addInflation
Dim inflation As Worksheet
Dim location
Application.ScreenUpdating = False
Set location = ActiveWorkbook.ActiveSheet
On Error GoTo ERR:
Set inflation = ActiveWorkbook.Worksheets(Range("Inflation_Raw").Worksheet.Name)
    
    addInflation = MsgBox("Inflation Sheet Already Exists." & vbNewLine & "Would you like to overwrite?", vbYesNo, "Inflation Add-In")
    Select Case addInflation
        Case vbYes
    
        inflation.Cells.Clear
        ThisWorkbook.Worksheets("Inflation").Cells.Copy
        inflation.Range("A1").PasteSpecial (xlPasteAll)
        Case vbNo
        addInflation = MsgBox("Inflation Sheet Already Exists." & vbNewLine & "Would you like to add an additional inflation tab?", vbYesNo, "Inflation Add-In")
            Select Case addInflation
                Case vbYes
                ThisWorkbook.Worksheets("Inflation").Copy before:=ActiveWorkbook.Worksheets(1)
            End Select
        'do nothing
    End Select
'Debug.Print location.Address
location.Activate
Application.ScreenUpdating = True

Exit Sub
'Error handler fix number one
ERR:
On Error GoTo lastTry
ThisWorkbook.Worksheets("Inflation").Copy before:=ActiveWorkbook.Worksheets(1)
location.Activate
Application.ScreenUpdating = True
Exit Sub
'Error handler fix number two

lastTry:
On Error Resume Next
Dim newWS As Worksheet
Set newWS = ActiveWorkbook.Worksheets.Add
newWS.Name = "Inflation"
ThisWorkbook.Worksheets("Inflation").UsedRange.Copy (newWS.Range("B2"))
location.Activate
Application.ScreenUpdating = True

End Sub
