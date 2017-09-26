Attribute VB_Name = "f_Inflation"
Option Explicit

Function BYtoBY(Index, FromYR, ToYR, Cost)
Attribute BYtoBY.VB_Description = "Returns a specified Base Year amount given a Base Year (uses Inflation tables located on inflation worksheet)"
Attribute BYtoBY.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
Dim raw As Variant
Dim nToYR, bFrom, bTo, Indice As Integer
Dim rawYR, rawIndice, rawTable As Range

Set rawYR = Range("Inflation_Raw[#headers]")
Set rawIndice = Range("Inflation_Raw[Raw Index]")
Set rawTable = Range("Inflation_Raw")


'nToYR = ToYR
'If ToYR > WorksheetFunction.Max(rawYR) Then nToYR = WorksheetFunction.Max(rawYR)

bFrom = WorksheetFunction.Match(CStr(FromYR), rawYR, False)
bTo = WorksheetFunction.Match(CStr(ToYR), rawYR, False)
Indice = WorksheetFunction.Match(Index, rawIndice, False)

BYtoBY = (Cost / rawTable(Indice, bFrom)) * rawTable(Indice, bTo)

End Function

Function BYtoTY(Index, FromYR, ToYR, Cost)
Attribute BYtoTY.VB_Description = "Returns a specified Then Year amount given a Base Year (uses Inflation tables located on inflation worksheet)"
Attribute BYtoTY.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
Dim wtdTable, wtdYR, wtdIndice As Range
Dim rawTable, rawYR, rawIndice As Range
Dim nToYR, cFrom, bTo, r_Indice, w_Indice As Integer

Set wtdYR = Range("Inflation_Weighted[#headers]")
Set wtdTable = Range("Inflation_Weighted")
Set wtdIndice = Range("Inflation_Weighted[Weighted Index]")

Set rawYR = Range("Inflation_Raw[#headers]")
Set rawTable = Range("Inflation_raw")
Set rawIndice = Range("Inflation_raw[Raw Index]")

'If nToYR > 2060 Then nToYR = 2060


cFrom = WorksheetFunction.Match(CStr(FromYR), rawYR, False)
bTo = WorksheetFunction.Match(CStr(ToYR), wtdYR, False)
r_Indice = WorksheetFunction.Match(Index, rawIndice, False)
w_Indice = WorksheetFunction.Match(Index, wtdIndice, False)

BYtoTY = (Cost / rawTable(r_Indice, cFrom)) * wtdTable(w_Indice, bTo)


End Function

Function TYtoBY(Index, FromYR, ToYR, Cost)
Attribute TYtoBY.VB_Description = "Returns a specified Base Year amount given a Then Year (uses Inflation tables located on inflation worksheet)"
Attribute TYtoBY.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
Dim wtdTable, wtdYR, wtdIndice As Range
Dim rawTable, rawYR, rawIndice As Range
Dim nToYR, cFrom, bTo, r_Indice, w_Indice As Integer

Set wtdYR = Range("Inflation_Weighted[#headers]")
Set wtdTable = Range("Inflation_Weighted")
Set wtdIndice = Range("Inflation_Weighted[Weighted Index]")

Set rawYR = Range("Inflation_Raw[#headers]")
Set rawTable = Range("Inflation_raw")
Set rawIndice = Range("Inflation_raw[Raw Index]")

'If nToYR > 2060 Then nToYR = 2060


cFrom = WorksheetFunction.Match(CStr(FromYR), rawYR, False)
bTo = WorksheetFunction.Match(CStr(ToYR), wtdYR, False)
r_Indice = WorksheetFunction.Match(Index, rawIndice, False)
w_Indice = WorksheetFunction.Match(Index, wtdIndice, False)

TYtoBY = (Cost / wtdTable(w_Indice, cFrom)) * rawTable(r_Indice, bTo)

End Function


Function TYCalc(Cost As Range, Indice, BaseYear, FirstYear) As Variant
Attribute TYCalc.VB_Description = "Returns a Then Year amount based on Base Year calculations. (optimized to run as an array)"
Attribute TYCalc.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
Dim FYBegin, FYEnd As Integer
Dim newCost As Range
Dim CurrentYear
Dim thing

    If WorksheetFunction.Sum(Cost) = 0 Then GoTo done
    If Indice = 0 Or BaseYear = 0 Then
        TYCalc = "Missing Indice or BY"
        GoTo done
    End If
    
   
Dim results() As Variant
ReDim results(1 To 1, FirstYear To FirstYear + 100)
Dim i As Integer
Set newCost = Range(Cells(Cost.row, Cost.Column).Address)

CurrentYear = FirstYear
For Each thing In Cost
    If thing > 0 Then
    
        results(1, CurrentYear) = BYtoTY(Indice, BaseYear, CurrentYear, thing)

        Set newCost = Union(newCost, thing)
        If WorksheetFunction.Sum(Cost) = WorksheetFunction.Sum(newCost) Then Exit For
       
    End If
     CurrentYear = CurrentYear + 1
Next

TYCalc = results
done:
End Function
