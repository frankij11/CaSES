Attribute VB_Name = "p_Rayleigh"
Option Explicit


Public Function RayleighPhasing(StartDate, Duration, Cost, delta, FY)
Attribute RayleighPhasing.VB_Description = "Returns a specified phasing profile based on the Rayleigh distribution.Time phasing functions allow for dynamic phasing of any cost estimate, cost element, or WBS element"
Attribute RayleighPhasing.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
If delta = 0 Then delta = 0.03
If StartDate > 4000 Then StartDate = DATEtoFYfrac(StartDate)

Dim s As Double
Dim EndDate As Double
Dim currentFY As Integer
Dim T1, T0 As Double
Dim CDF1, CDF0 As Double
Dim PDF As Double
Dim BYConv, UnitConv As Double
Dim LaborRates As Double
Dim LaborCat, LaborYr As Integer
Dim countYR                         ' counts the number of years from start date to calc labor year

Dim res()
ReDim res(1 To 1, FY To FY + 100)

'calculate sigma (s) as function of duration
s = (Duration ^ 2 / (Abs(Log(delta)) * 2)) ^ 0.5
Cost = Cost / (1 - delta)
' RaleighPhasing Macro
' (StartDate, Duration, Cost, A, FY)
'
EndDate = StartDate + Duration - 0.0000001



For currentFY = Int(StartDate) To Int(EndDate)
If currentFY < FY Then ' Do Nothing
Else
countYR = countYR + 1
T1 = (WorksheetFunction.Min(EndDate, currentFY + 1) - StartDate) / (EndDate - StartDate)
T0 = (WorksheetFunction.Min(EndDate, currentFY) - StartDate) / (EndDate - StartDate)



If T0 < 0 Then T0 = 0
If T1 < 0 Then T1 = 0
'If t1 > 1 Then t1 = 1
If T1 >= 0 And T1 <= 1 Then
     
    CDF1 = 1 - Exp(-((T1 * Duration) ^ 2) / (2 * s ^ 2))
    CDF0 = 1 - Exp(-((T0 * Duration) ^ 2) / (2 * s ^ 2))
    
    PDF = CDF1 - CDF0
        
    
        
    res(1, currentFY) = PDF * Cost
ElseIf Int(EndDate) = currentFY Then
    CDF1 = 1
    CDF0 = 1 - Exp(-(T0 ^ 2) / (2 * s ^ 2))
    PDF = CDF1 - CDF0
    
    
    
    
    res(1, currentFY) = PDF * Cost
End If
End If
Next currentFY

RayleighPhasing = res


End Function
