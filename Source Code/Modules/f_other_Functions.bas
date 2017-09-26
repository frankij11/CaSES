Attribute VB_Name = "f_other_Functions"
Option Explicit


Sub FullCalculation()
Application.CalculateFull
End Sub


Function getComment(c As Range) As String
Application.Volatile (True)
On Error Resume Next
getComment = ""
getComment = c.Comment.text
End Function
Public Function indexCaSES(Range_Lookup As Range, Row_Lookup, Column_Lookup)
Attribute indexCaSES.VB_Description = "Looks up the value that corresponds to the Row and column specificied"
Attribute indexCaSES.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim rowMatch, colMatch As Range
    Set rowMatch = Range(Range_Lookup(1, 1).Address, Range_Lookup(Range_Lookup.Rows.count, 1).Address)
    Set colMatch = Range(Range_Lookup(1, 1).Address, Range_Lookup(1, Range_Lookup.Columns.count).Address)
    indexCaSES = Range_Lookup(WorksheetFunction.Match(Row_Lookup, rowMatch, 0), WorksheetFunction.Match(Column_Lookup, colMatch, 0))
End Function

Private Function WBSLev(WBS, level)
Dim wbsArr As Variant
Dim wbsArr2 As Variant
Dim i

wbsArr = Split(WBS)
WBS = wbsArr(0)
wbsArr2 = Split(WBS, ".")
For i = 1 To 8
ReDim Preserve wbsArr2(1 To 8)
If wbsArr2(i) = "" Then wbsArr2(i) = 0
Next i

WBSLev = wbsArr2(level) * 1


End Function



Function midpointAsh(First, last, b)

midpointAsh = ((1 / (last - First + 1)) * ((((last + 0.5) ^ (1 + b)) - ((First - 0.5) ^ (1 + b))) / (1 + b))) ^ (1 / b)
End Function

Public Function LCunit(T1, LC, RC, Priors, LotQty)

Dim b, c As Double
Dim f, l As Integer
Dim mp As Double

b = Log(LC) / Log(2)
c = Log(RC) / Log(2)

f = Priors + 1
l = f + LotQty - 1
mp = midpointAsh(f, l, b)
On Error Resume Next
LCunit = (T1 * mp ^ b * LotQty ^ c) * LotQty

If LCunit = "" Then LCunit = 0


End Function

'Function LCunit(x, q, T1, b, r)
'LCunit = T1 * x ^ b * q ^ r
'End Function


Public Function UnitC(Cost, FromD, ToD) As Variant
Attribute UnitC.VB_Description = "Returns specified units (in $) based on a given unit"
Attribute UnitC.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
FromD = VBA.VBA.Trim(VBA.VBA.LCase(FromD))
ToD = VBA.VBA.Trim(VBA.VBA.LCase(ToD))
Dim conver As Double

Select Case FromD
Case "$", "dollar", "hour", "hr", "hrs"
FromD = 1
Case "$k", "hours k", "hrs k"
FromD = 1000
Case "$m"
FromD = 1000000
Case "$b"
FromD = 1000000000
End Select

Select Case ToD
Case "$", "dollar", "hour", "hr", "hrs"
ToD = 1
Case "$k", "hours k", "hrs k"
ToD = 1000
Case "$m"
ToD = 1000000
Case "$b"
ToD = 1000000000
End Select

conver = FromD / ToD
UnitC = Cost * FromD / ToD

End Function

