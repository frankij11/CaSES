Attribute VB_Name = "p_Custom"
Option Explicit

Function resizeRange(sRange As Range) As Variant
Application.Volatile (True)
Dim done As Boolean
Dim thing
Dim count
For Each thing In sRange
    If count > 0 Then
        count = count + 1
        Set resizeRange = Union(resizeRange, thing)
    ElseIf WorksheetFunction.Sum(thing) <> 0 Then
        count = count + 1
        Set resizeRange = thing
    End If
    If WorksheetFunction.Sum(sRange) = WorksheetFunction.Sum(resizeRange) Then
        If WorksheetFunction.Sum(resizeRange) = 0 Then Set resizeRange = thing
        Exit Function
    End If
Next
End Function


Public Function CustomPhasing(exSD, exDur, StartDate, Duration, YearlyAmount As Range, ScheduleSlip As Double, FY) As Variant
Attribute CustomPhasing.VB_Description = "Returns a specified phasing profile based on a custom (user defined) distribution. Time phasing functions allow for dynamic phasing of any cost estimate, element, or WBS element."
Attribute CustomPhasing.VB_ProcData.VB_Invoke_Func = " \n19"
Application.Volatile (True)
Dim EndDate, exEndDate, EndDateSlip As Double
Dim SD_FY, ED_FY As Integer
Dim exSD_Frac, exED_Frac, ED_FracSLip, SD_Frac, ED_Frac  As Double
Dim slip As Double
Dim i, t As Integer
Dim l As Double
Dim SlipPoint As Integer
Dim countYR
 
Dim costString As String
    costString = Replace(YearlyAmount.Formula, "=SUM", "")
    If costString = "" Or costString = "0" Then GoTo endFunction Else Set YearlyAmount = Range(costString)
    
'resize YearlyAmount range to only include non Zero values
Set YearlyAmount = resizeRange(YearlyAmount)

'Caculate the cumulative time at which the schedule slips
SlipPoint = WorksheetFunction.Round(YearlyAmount.count * ScheduleSlip, 0)
SlipPoint = WorksheetFunction.Max(SlipPoint, 1)


'convert Julian Date to FY.QTR format
If exSD > 4000 Then exSD = DATEtoFYfrac(exSD)
If StartDate > 4000 Then StartDate = DATEtoFYfrac(StartDate)


'Dim p(1 To 1, 1 To 100)
Dim p() As Variant
ReDim p(1 To 1, FY To FY + 100)
 
EndDate = StartDate + Duration - 0.0000001
exEndDate = exSD + exDur - 0.0000001
EndDateSlip = StartDate + exDur - 0.0000001

SD_FY = Int(StartDate)
ED_FY = Int(EndDate)

exSD_Frac = (Int(exSD) + 1) - exSD
exED_Frac = exEndDate - Int(exEndDate)

SD_Frac = (WorksheetFunction.Min(Int(StartDate) + 1, EndDate)) - StartDate
ED_Frac = EndDate - Int(EndDate)
ED_FracSLip = EndDateSlip - Int(EndDateSlip)

' ScheduleSlip = 1 then add activities
Dim Cost() As Variant 'stores cost when when schedule slip is eqaul 0
Dim estimate() As Variant 'stores cost when when schedule slip is greater than 0

''''Normalize each year to a full year
''''i.e. Convert YearlyAmount partial years to full amounts

If YearlyAmount.count = 1 Then 'if the YearlyAmount range is a single cell then convert to an array
    ReDim estimate(1 To 1, 1 To 1)
    ReDim Cost(1 To 1, 1 To 1)
    If ScheduleSlip > 0 Then
        Cost(1, 1) = YearlyAmount / (exED_Frac)
        estimate(1, 1) = YearlyAmount / (exED_Frac)
    Else
        Cost(1, 1) = YearlyAmount
    End If

Else
    Cost = YearlyAmount
    estimate = YearlyAmount
   If ScheduleSlip > 0 Then
    ReDim Preserve Cost(1 To 1, 1 To YearlyAmount.count + 1) 'add additional year to hold partial year
    Cost(1, 1) = Cost(1, 1) / exSD_Frac
    Cost(1, YearlyAmount.count) = Cost(1, YearlyAmount.count) / exED_Frac
    
    ReDim estimate(1 To 1, 1 To UBound(Cost, 2))
If exSD_Frac >= SD_Frac Then
    estimate(1, 1) = Cost(1, 1) * SD_Frac
    slip = exSD_Frac - SD_Frac
    For i = 2 To UBound(estimate, 2) - 2
        estimate(1, i) = Cost(1, i - 1) * slip + Cost(1, i) * (1 - slip)
    Next i
    l = WorksheetFunction.Min(1 - slip, exED_Frac) * (Cost(1, YearlyAmount.count))
    estimate(1, YearlyAmount.count) = Cost(1, YearlyAmount.count - 1) * slip + l
    estimate(1, YearlyAmount.count + 1) = Cost(1, YearlyAmount.count) * exED_Frac - l
Else
    slip = SD_Frac - exSD_Frac
    ReDim Preserve estimate(1 To 1, 1 To YearlyAmount.count)
    estimate(1, 1) = Cost(1, 1) * exSD_Frac + Cost(1, 2) * slip
    For i = 2 To UBound(estimate, 2) - 2
    estimate(1, i) = Cost(1, i + 1) * slip + Cost(1, i) * (1 - slip)
    Debug.Print ("YR" & i & "cost = " & estimate(1, i))
    Next i
        
    estimate(1, YearlyAmount.count - 1) = (Cost(1, YearlyAmount.count - 1) * (1 - slip)) + WorksheetFunction.Min(slip, exED_Frac) * Cost(1, YearlyAmount.count)
    estimate(1, YearlyAmount.count) = Cost(1, YearlyAmount.count) * ED_Frac - l
End If
      If estimate(1, UBound(estimate, 2)) = 0 Then ReDim Preserve estimate(1 To 1, 1 To UBound(Cost, 2) - 1)
    End If
    estimate(1, UBound(estimate, 2)) = estimate(1, UBound(estimate, 2)) / ED_FracSLip
    estimate(1, 1) = estimate(1, 1) / SD_Frac
End If

' ScheduleSlip = 0 then use only Cost range
    '' works for Procurement Type activtities
t = 0
Dim currentFY 'determines the current fiscal year during iterations
Dim convFactor 'factor to convert estimate BY, $, and Labor Rate
Select Case ScheduleSlip
        
    Case 0
        
        'For i = SD_FY - FY + 1 To (SD_FY - FY) + UBound(Cost, 2)
        For currentFY = SD_FY To SD_FY + UBound(Cost, 2) - 1
        t = t + 1
            If currentFY < FY Then 'Current Fiscal Year is less than initial Fiscal YEar then do nothing
                'Do Nothing
            Else
                
                p(1, currentFY) = Cost(1, t)
            End If
        Next currentFY
            

' ScheduleSlip > 0 then schedule slip happens during X% of cumaltive time of program
    '' works for Staffing/burn rate Type activtities
    Case Is > 0
        
        For currentFY = SD_FY To ED_FY
            t = t + 1
            If currentFY < FY Then 'Current Fiscal Year happens before first available year
                'Do Nothing
            Else          'Calculate cost for Current Fiscal Year
            
                If SD_FY = currentFY Then 'If start date is equal to Current Fiscal Year then multiply first year value by appropriate fraction
                    p(1, currentFY) = estimate(1, 1) * SD_Frac
                
                ElseIf ED_FY = currentFY Then 'If end date is equal to Current FY then multiply last value by appropriate fraction
                    p(1, currentFY) = estimate(1, WorksheetFunction.Min(t, UBound(estimate, 2))) * ED_Frac
                
                ElseIf ED_FY - SD_FY + 1 <= UBound(estimate, 2) Then
                    p(1, currentFY) = estimate(1, WorksheetFunction.Min(t, UBound(estimate, 2)))
                
                Else
                    If t <= SlipPoint Then
                        p(1, currentFY) = estimate(1, WorksheetFunction.Min(t, UBound(estimate, 2)))
                    
                    ElseIf t > SlipPoint And t <= (SlipPoint + ED_FY - SD_FY - UBound(estimate, 2) + 1) Then
                        
                        p(1, currentFY) = estimate(1, SlipPoint)
                    Else
                        p(1, currentFY) = estimate(1, t - (ED_FY - SD_FY - UBound(estimate, 2) + 1))
                    
                    End If
                End If
        End If
    Next currentFY
 End Select
 
CustomPhasing = p

endFunction:
End Function

Sub trialCust()
Dim a
Dim b As Range
Set b = Range("AM19")
a = CustomPhasing(0, 0, 2020, 27, Range("BY_First_Year"), b, 1, 2016, "APN", "$", "Government Composite Rate (TY$)")
End Sub
