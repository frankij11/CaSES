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

