Attribute VB_Name = "M_traceBack_Formula_Auditing"
Option Explicit
Sub Formula_Auditing()
TraceBackNavigator.Show

End Sub


Sub ListFormulas()
    Dim start As Double
    Dim FormulaCells As Range, cell As Range
    Dim FormulaSheet As Worksheet
    Dim row As Integer
    start = Time
'   Create a Range object for all formula cells
    On Error Resume Next
    Set FormulaCells = Range("A1").SpecialCells(xlFormulas, 23)
    
'   Exit if no formulas are found
    If FormulaCells Is Nothing Then
        MsgBox "No Formulas."
        Exit Sub
    End If
    
'   Add a new worksheet
    Application.ScreenUpdating = False
    Set FormulaSheet = ActiveWorkbook.Worksheets.Add
    
    FormulaSheet.Name = Left("Formulas in " & FormulaCells.Parent.Name, 20)
    
'   Set up the column headings
    With FormulaSheet
        Range("A1") = "Address"
        Range("B1") = "Formula"
        Range("C1") = "Value"
        Range("A1:C1").Font.Bold = True
    End With
    
'   Process each formula
    row = 2
    For Each cell In FormulaCells
        Application.StatusBar = Format((row - 1) / FormulaCells.count, "0%")
        With FormulaSheet
            Cells(row, 1) = cell.Address _
                (RowAbsolute:=False, ColumnAbsolute:=False)
            Cells(row, 2) = " " & cell.Formula
            Cells(row, 3) = cell.Value
            row = row + 1
        End With
    Next cell
    
'   Adjust column widths
    FormulaSheet.Columns("A:C").AutoFit
    Application.StatusBar = False
    MsgBox ("Finished in " & Time - start)
End Sub



Function removeQuote(text As String)
Dim easy, eB, i
easy = Split(text, VBA.Chr(34))
    eB = text
    For i = 1 To UBound(easy, 1)
      eB = Replace(eB, VBA.Chr(34) & easy(i) & VBA.Chr(34), "")
        
        i = i + 1
    Next i
text = eB
removeQuote = text
End Function

Function sepParen(text As String) As Variant
Dim oP()
Dim cc As Integer
Dim pp As Integer
Dim oPar As Integer
Dim c
c = 1
pp = 1
text = "(" & text & ")"
Do Until c = Len(text) + 1
If Mid(text, c, 1) = "(" Then
    cc = 1
    'pp = pp + 1
    oPar = 1
    ReDim Preserve oP(1 To pp)
    Do Until oPar = 1 And Mid(text, c + cc, 1) = ")"
       If Mid(text, c + cc, 1) = "(" Then
        oPar = oPar + 1
            Else
                If Mid(text, c + cc, 1) = ")" Then oPar = oPar - 1
        'Debug.Print (Mid(b, c + cc, 1))
        End If
        cc = cc + 1
    Loop
    temp = SplitEx(Mid(text, c + 1, cc - 1), True, ",")
    For i = LBound(temp, 1) To UBound(temp, 1)
        If Len(temp(i)) - Len(Replace(Replace(temp(i), "(", ""), ")", "")) > 0 Then 'do nothing because next loop will catch parenthesis
        Else
        ReDim Preserve oP(1 To pp)
        oP(pp) = VBA.Trim(temp(i))
        'Debug.Print (oP(pp))
        pp = pp + 1
        End If
    Next i
    
    
End If
c = c + 1
Loop
ReDim Preserve oP(1 To pp)
sepParen = oP
End Function

Function isRange(M)
Dim temp, M, c, cc
temp = Len(M) - Len(Replace(M, "'", ""))
If temp > 0 Then
    c = WorksheetFunction.Find("'", M)
    cc = WorksheetFunction.Find("'", M, c + 1)
    'wks = Mid(M, c + 1, cc - 2)
    'M = Replace(M, "'" & wks & "'", "")
    If Mid(M, 1, 1) = "!" Then M = Right(M, Len(M) - 1)
End If
    On Error GoTo ERR
Set result = Range(M)
isRange = True
Exit Function
ERR:
isRange = False

End Function



Public Function determineRanges(a As String) As Variant
Dim result(0 To 20, 1 To 4)
On Error Resume Next
a = Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1)
Dim b()
ReDim Preserve b(1 To 50, 0 To 10)
a = removeQuote(a)
c = sepParen(a)
For i = LBound(c, 1) To UBound(c, 1)
'Debug.Print (c(i))
temp = SplitEx(c(i), True, "=", "+", "-", ",", "/", "*", "^", "&")
big = WorksheetFunction.Max(UBound(b, 2), UBound(temp, 1))
'ReDim Preserve b(1 To i, 0 To big)
For ii = LBound(temp, 1) To UBound(temp, 1)
b(i, ii) = temp(ii)
'Debug.Print (b(i, ii))
Next ii
Next i
report = ActiveCell.Address
counter = 0
For i = LBound(b, 1) To UBound(b, 1)
For ii = LBound(b, 2) To UBound(b, 2)

If b(i, ii) = "" Then
Else
'Debug.Print (isRange(b(i, ii)) & " " & b(i, ii))
'ReDim Preserve result(1 To counter, 1 To 4)
result(counter, 1) = b(i, ii)
result(counter, 2) = Range(b(i, ii)).Formula
result(counter, 3) = Range(b(i, ii)).Value
Debug.Print (result(counter, 1) & " " & result(counter, 2) & " " & result(counter, 3))
counter = counter + 1
'report = report & Chr(13) & isRange(b(i, ii)) & " " & b(i, ii)
End If
Next ii
Next i
determineRanges = result
'MsgBox (report)
End Function
Function SplitEx(ByVal InString As String, IgnoreDoubleDelmiters As Boolean, _
        ParamArray Delims() As Variant) As String()
    Dim Arr() As String
    Dim Ndx As Long
    Dim N As Long
    
    If Len(InString) = 0 Then
        ReDim Preserve Arr(0 To 0)
        SplitEx = Arr
        Exit Function
    End If
    If IgnoreDoubleDelmiters = True Then
        For Ndx = LBound(Delims) To UBound(Delims)
            N = InStr(1, InString, Delims(Ndx) & Delims(Ndx), vbTextCompare)
            Do Until N = 0
                InString = Replace(InString, Delims(Ndx) & Delims(Ndx), Delims(Ndx))
                N = InStr(1, InString, Delims(Ndx) & Delims(Ndx), vbTextCompare)
            Loop
        Next Ndx
    End If
    
    
    ReDim Arr(1 To Len(InString))
    For Ndx = LBound(Delims) To UBound(Delims)
        InString = Replace(InString, Delims(Ndx), Chr(1))
    Next Ndx
    Arr = Split(InString, Chr(1))
    SplitEx = Arr
End Function


Public Function tracePrecedents() As Variant
Application.ScreenUpdating = False
ActiveSheet.ClearArrows
Dim test
Dim a()
Dim b, Target, testN, testL, i, ii, c, temp
b = ActiveCell.Address(, , , True)
Set Target = Range(b)
ReDim Preserve a(1 To 1)
On Error GoTo ext:
Target.ShowPrecedents
test = True
testN = True
testL = True
i = 1
ii = 1
c = 1
'First Link
Do Until testL = False Or Range(Target.NavigateArrow(True, 1, ii).Address)(1, 1).Address = Target.Address
    On Error GoTo second:
    ReDim Preserve a(1 To c)
    temp = Target.NavigateArrow(True, 1, ii).Address
    'Debug.Print (temp)
    'ReDim Preserve a(1 To i, 1 To 20)
    a(c) = Target.NavigateArrow(True, 1, ii).Address
    a(c) = "'" & Range(a(c)).Parent.Name & "'!" & a(c)
    Debug.Print ("a(" & c & ")" & a(c))
    ii = ii + 1
    c = c + 1
Loop

'Second and on arrow
second:
i = 2

testL = False
Do Until Range(Target.NavigateArrow(True, i, 1).Address(, , , True))(1, 1).Address(, , , True) = Target.Address(, , , True)

    ReDim Preserve a(1 To c)
    a(c) = Target.NavigateArrow(True, i, 1).Address
    a(c) = "'" & Range(a(c)).Parent.Name & "'!" & a(c)
    Debug.Print ("a(" & c & ")" & a(c))
    i = i + 1
    c = c + 1
Loop
ext:
tracePrecedents = a
Target.Select
ActiveSheet.ClearArrows
Application.ScreenUpdating = True
End Function

Function traceDependents() As Variant
Application.ScreenUpdating = False
ActiveSheet.ClearArrows
Dim test
Dim a()
Dim b, Target, testN, testL, i, ii, c, temp
b = ActiveCell.Address(, , , True)
Set Target = Range(b)
ReDim Preserve a(1 To 1)
On Error GoTo ext:
Target.ShowDependents
test = True
testN = True
testL = True
i = 1
ii = 1
c = 1
'First Link
Do Until testL = False Or Range(Target.NavigateArrow(False, 1, ii).Address)(1, 1).Address = Target.Address
    On Error GoTo second:
    ReDim Preserve a(1 To c)
    temp = Target.NavigateArrow(False, 1, ii).Address(, , , True)
    'Debug.Print (temp)
    'ReDim Preserve a(1 To i, 1 To 20)
    a(c) = Target.NavigateArrow(False, 1, ii).Address
    a(c) = "'" & Range(a(c)).Parent.Name & "'!" & a(c)
    Debug.Print ("a(" & c & ")" & a(c))
    ii = ii + 1
    c = c + 1
Loop

'Second and on arrow
second:
i = 2

testL = False
Do Until Range(Target.NavigateArrow(False, i, 1).Address(, , , True))(1, 1).Address(, , , True) = Target.Address(, , , True)

    ReDim Preserve a(1 To c)
    a(c) = Target.NavigateArrow(False, i, 1).Address
    a(c) = "'" & Range(a(c)).Parent.Name & "'!" & a(c)
    Debug.Print ("a(" & c & ")" & a(c))
    i = i + 1
    c = c + 1
Loop
ext:
traceDependents = a
Target.Select
ActiveSheet.ClearArrows
Application.ScreenUpdating = True
End Function

