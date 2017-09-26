Attribute VB_Name = "f_FiscalYearCalc"
Option Explicit

'Calculations are necessary for other phasing functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function DATEtoFYint(curr_date As Date)
Attribute DATEtoFYint.VB_Description = "Returns an integer fiscal Year given a date"
Attribute DATEtoFYint.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
'Determines Julian date associated fiscal year
If Month(curr_date) >= 1 And Month(curr_date) <= 9 Then

    DATEtoFYint = Year(curr_date)
Else
    DATEtoFYint = Year(curr_date) + 1
End If
End Function


Function DATEtoFYfrac(D)
Attribute DATEtoFYfrac.VB_Description = "Returns a Fiscal Year Date in format YYYY.Frac (i.e. 2014.25 reads as 2nd Quarter FY 2014)"
Attribute DATEtoFYfrac.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
'converts Julian Date to YYYY.QTR in Fiscal Years
Dim a As Date
Dim FY, b, c
a = CDate(D)
FY = DATEtoFYint(a)
b = "October 1, " & FY - 1
b = CDate(b)
c = WorksheetFunction.YearFrac(a, b, 1)
DATEtoFYfrac = FY + c
End Function

Function FYtoDate(FiscalYear As Double) As Date
Attribute FYtoDate.VB_Description = "Returns Julian date given a Fiscal Year  (YYYY.Frac)."
Attribute FYtoDate.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
Dim beginFY, endFY, daysInYear
beginFY = "October 1, " & Int(FiscalYear) - 1
endFY = "September 30, " & Int(FiscalYear)
daysInYear = DateValue(endFY) - DateValue(beginFY) + 1

FYtoDate = DateAdd("d", (FiscalYear - Int(FiscalYear)) * daysInYear, beginFY)


End Function




