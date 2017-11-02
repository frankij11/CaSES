Attribute VB_Name = "p_Beta"
Option Explicit


Public Function BetaPhasing(StartDate, Duration, Cost, a, b, FY) ' (StartDate, Duration, Cost, A, FY)
Attribute BetaPhasing.VB_Description = "Returns a specified phasing profile based on the modified beta distribution by Richard Whitlock of NASA\n\nShould be based on analogous programs, COSTAT allows a fit of actual data"
Attribute BetaPhasing.VB_ProcData.VB_Invoke_Func = " \n19"
'
' BetaPhasing Macro
' (StartDate, Duration, Cost, A, FY)
' Phasing based on NASA study
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.Volatile (True)
Dim T1 As Double        'current periods cumulative time
Dim T0 As Double        'previous periods cumulative time
Dim results() As Variant   'Results is a variable to hold the final answer
Dim CDF1 As Double      'current periods cumulative cost
Dim CDF0 As Double      'previous periods cumulative cost
Dim PDF As Double       'current periods cost
Dim EndDate As Double
Dim currentFY As Integer
Dim LaborRates As Double
Dim LaborCat, LaborYr As Integer
Dim countYR                         ' counts the number of years from start date to calc labor year

ReDim results(1 To 1, FY To FY + 100)    'Bogus size array, would like to fix to be dynamic/more efficient

'More Efficient Algorithm to hold results array (Not working yet)
            ''''With Application.Caller
            ''''    ccc = .Columns.Count
            ''''    rrr = .Rows.Count
            ''''End With
            ''''ReDim Preserve results(1 To rrr, 1 To ccc)


' If Start Date is greater than 4000 we assume that this is a julian date and we convert back to YEAR.FRAC format
' YEAR.FRAC format : Year = Fiscal Year, FRAC = fraction of Fiscal Year
    If StartDate > 4000 Then StartDate = DATEtoFYfrac(StartDate)
    ' Assumes that work finished one day before End Date
    ' Necessary in order for calculations to be accurate
    EndDate = StartDate + Duration - 0.001 / 365

'This For loop determines the amount to phase in each year
'Only loops through through start and end dates

For currentFY = Int(StartDate) To Int(EndDate)
    If currentFY < FY Then
    'do nothing
    Else
    
    T1 = (WorksheetFunction.Min(EndDate, currentFY + 1) - StartDate) / (EndDate - StartDate)
    T0 = (WorksheetFunction.Min(EndDate, currentFY) - StartDate) / (EndDate - StartDate)
        If T0 < 0 Then T0 = 0
        If T1 < 0 Then T1 = 0
        If T1 >= 0 And T1 <= 1 Then
            'BetaPhasing Cumulative phasing curve given time T, alpha and beta
            CDF1 = 10 * (T1 ^ 2) * ((1 - T1) ^ 2) * (a + (b * T1)) + (T1 ^ 4) * (5 - (4 * T1))
            CDF0 = 10 * (T0 ^ 2) * ((1 - T0) ^ 2) * (a + (b * T0)) + (T0 ^ 4) * (5 - (4 * T0))
            PDF = CDF1 - CDF0
            
        End If
            results(1, currentFY) = PDF * Cost
    End If
    Next currentFY
    BetaPhasing = results
'If a > 1 Or a < 0 Or b < 0 Or a + b > 1 Then BetaPhasing = "Choose proper Alpha and Beta"
End Function
                        Public Function CaSES_Phasing_Beta(StartDate, Duration, Cost, a, b, FY)
                                    CaSES_Phasing_Beta = BetaPhasing(StartDate, Duration, Cost, a,b,FY)
                        End Function
                    
