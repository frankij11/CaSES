Attribute VB_Name = "p_Factor"
Option Explicit

Function Factor(FactorID, FactorColumn As Range, SumColumn As Range, Cost, FY, Optional StartDate, Optional Duration, Optional LaborCategory, Optional Repeat, Optional Cycle) As Variant
Application.Volatile (True)
Dim Item As Object
Dim splitItem
Dim IndividualFactors
Dim sumRange As Range
Dim count As Integer
Dim SumAnchorRange
Set SumAnchorRange = Range(Cells(SumColumn.row, SumColumn.Column).Address)


''''''''''count = 0
''''''''''Dim secondCount
''''''''''secondCount = -1


''''''''''For Each Item In FactorColumn
''''''''''    secondCount = secondCount + 1 'keep a count to know how man rows to offset the sum range
''''''''''    'If Item contains FactorID then add cell to sum range from SumColumn
''''''''''    IndividualFactors = Split(Cells(Item.row, Item.Column), ",")
''''''''''    For Each splitItem In IndividualFactors
''''''''''        If VBA.VBA.Trim(VBA.VBA.LCASE(splitItem)) = VBA.VBA.Trim(VBA.VBA.LCASE(FactorID)) Then
''''''''''            If count = 0 Then
''''''''''                count = 1
''''''''''                Set SumRange = Range(SumAnchorRange.Offset(secondCount).Address)
''''''''''
''''''''''            Else
''''''''''                Set SumRange = Union(SumRange, Range(SumAnchorRange.Offset(secondCount).Address))
''''''''''
''''''''''            End If
''''''''''        End If
''''''''''    Next ' Split Item
''''''''''Next ' Item
               
               
               
               
FactorID = "*" & FactorID & "*"
Factor = WorksheetFunction.SumIf(FactorColumn, FactorID, SumColumn) * Cost

End Function

Sub enterFactor()
Dim row As Integer
Dim fycol As Integer
Dim Cost As String
Dim col As Integer
Dim i As Integer
Dim ident As Variant
Dim a
Dim thing
Dim txt As String
Dim count As Integer
Dim YrsToAdd As Integer
Dim Factor As String
Dim laborRow, laborCol As Integer


row = ActiveCell.row
col = Range("Model_Factor").Column
fycol = Range("model_first_year").Column
If Cells(row, Range("model_labor_category").Column).Value = "" Then
    Cost = Cells(row, Range("model_cost").Column).Address
Else
    Cost = "LaborRate(" & Cells(row, Range("model_labor_category").Column).Address & "," & Range("Model_First_Year").Address(True, False) & ")"
    
End If
YrsToAdd = Range("Num_Yrs_est").Value
   
        If Cells(row, col).Value = "" Then
        'do nothing
        Else
        ident = Split(Cells(row, col).Value, ",")
        Factor = ident(0)
        End If
    
    For i = Range("model_factor").row + 1 To 500
        a = Split(Cells(i, col), ",")
        For Each thing In a
            If VBA.VBA.LCase(VBA.VBA.Trim(Factor)) = VBA.VBA.LCase(VBA.VBA.Trim(thing)) Then
                 
                If i = row Then
                'do nothing
                Else
                    If count = 0 Then
                    txt = "=SUM(" & Cells(i, fycol).Address(True, False)
                    Else
                    txt = txt & "," & Cells(i, fycol).Address(True, False)
                    End If
                count = count + 1
                End If
            End If
        Next
    Next
    
    

    
    
    If count = 0 Then
    Range(Cells(row, fycol), Cells(row, fycol + YrsToAdd)).ClearContents
    Else
    Cells(row, fycol) = txt & ") * " & Cost
    'Cells(Row, fycol).Copy
    Range(Cells(row, fycol), Cells(row, fycol + YrsToAdd - 1)).FillRight
    'Selection.PasteSpecial (xlFormulas)
    '(Range(Cells(Row, fycol), Cells(Row, fycol + 50)))
    End If

End Sub
