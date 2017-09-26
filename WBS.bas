Attribute VB_Name = "WBS"
Sub wbsGroup()

    
    Set rng1 = Application.InputBox(prompt:="select cell that contains first wbs", Type:=8)
    Set rng2 = Application.InputBox(prompt:="select cell that contains last wbs", Type:=8)
    Dim cell As Range
    
    Range(rng1.Address, rng2.Address).ClearOutline
    c = 0
    For Each cell In Range(rng1.Address, rng2.Address)
        summ = "=sum("
        summ = summ & cell.Address
        Cells(cell.Row, cell.Column + 1) = summ & ")"
        
        
        ind = Len(cell) - Len(Replace(cell, ".", ""))
        If ind > 7 Then
        c = c + 1
        tooBig = tooBig & " " & cell.Value2
        Else
        cell.NumberFormat = "@"
        cell.EntireRow.OutlineLevel = WorksheetFunction.Min(ind + 1, 8)
        End If
        cell.IndentLevel = ind
    Next cell

    If c > 0 Then MsgBox (c & " Items are too large to group" & tooBig)
    
End Sub

Sub sumWBS()
Set r = Application.InputBox("first cell of wbs", Type:=8)

last = r.End(xlDown).Row - r.Row + 1
k = r.Column
kk = Application.InputBox("First cell of Summation", Type:=8).Column
yrs = Application.InputBox("Number of Columns to add (Be Sure to Add a column for Totals", Type:=1)
For i = r.Row To r.End(xlDown).Row

ind = Cells(i, k).IndentLevel
ii = i + 1
tempsubtotal = ""
While ind < Cells(ii, k).IndentLevel

hmm = Cells(ii, k).IndentLevel - ind
If Cells(ii, k).IndentLevel - ind = 1 Then
tempsubtotal = tempsubtotal & Cells(ii, kk).Address(False, False) & ", "
End If
ii = ii + 1
Wend
'newstring = "=subtotal(9," & Left(tempsubtotal, Len(tempsubtotal) - 2) & ")"
If Not tempsubtotal = "" Then
    Cells(i, kk) = "=sum(" & Left(tempsubtotal, Len(tempsubtotal) - 2) & ")"
    Set rngYR = Range(Cells(i, kk), Cells(i, kk + yrs - 1))
    Cells(i, kk).Copy (rngYR)
End If
Next i
End Sub
