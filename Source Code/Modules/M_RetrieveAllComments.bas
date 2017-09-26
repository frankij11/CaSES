Attribute VB_Name = "M_RetrieveAllComments"
Option Explicit

Sub M_Retrieve_AllComments()

'Referenced from https://tduhameau.wordpress.com/
'get all comments from a workbook and put them in a new worksheet
 Dim rgCmt As Range, rgComments As Range, lRowLoop As Long, shtLoop As Worksheet, shtComments As Worksheet

Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual

Sheets.Add before:=Sheets(1)
Application.DisplayAlerts = False

On Error Resume Next
ActiveWorkbook.Sheets("CellComments").Delete


 Set shtComments = ActiveSheet

With shtComments

'create and format the comment summary sheet
shtComments.Name = "CellComments"

With .Columns("A:D")

.VerticalAlignment = xlTop
 .WrapText = True

End With

.Columns("B").ColumnWidth = 15
 .Columns("C").ColumnWidth = 60
 .PageSetup.PrintGridlines = True

.[a1] = "Sheet"
.[b1] = "Cell"
'.[c1] = "Value"
.[c1] = "Comment"

.Rows(1).Font.Bold = True

'.Tab.Color = 0
 .Tab.TintAndShade = 0

End With

lRowLoop = 2

For Each shtLoop In ActiveWorkbook.Worksheets

'loop through all worksheets and retrieve the comments
 If shtLoop.Name <> shtComments.Name And shtLoop.Comments.count > 0 Then

On Error Resume Next
 Set rgComments = shtLoop.Cells.SpecialCells(xlCellTypeComments)

If ERR = 0 Then

For Each rgCmt In rgComments.Cells

If VBA.Trim(rgCmt.Comment.text) <> "" Then

shtComments.Cells(lRowLoop, 1) = shtLoop.Name
shtComments.Hyperlinks.Add Anchor:=shtComments.Cells(lRowLoop, 2), Address:="", _
 SubAddress:="'" & shtLoop.Name & "'!" & rgCmt.Address(0, 0), TextToDisplay:=rgCmt.Address(0, 0)
'shtComments.Cells(lRowLoop, 3) = "'" & rgCmt.Text
shtComments.Cells(lRowLoop, 3) = "'" & rgCmt.Comment.text
 lRowLoop = lRowLoop + 1

End If

Next rgCmt

Else

ERR.Clear

End If

End If

Next shtLoop

shtComments.Activate

'clean up
 If Application.WorksheetFunction.CountA(shtComments.Columns(1)) = 1 Then

MsgBox "No comments in workbook"
Application.DisplayAlerts = False
 shtComments.Delete
 Application.DisplayAlerts = True

End If

Application.Calculation = xlCalculationAutomatic 'xl95 uses xlAutomatic
 Application.ScreenUpdating = True
 ActiveWindow.DisplayGridlines = False
 End Sub


