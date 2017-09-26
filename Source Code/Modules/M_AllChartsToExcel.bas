Attribute VB_Name = "M_AllChartsToExcel"
Option Explicit

Sub M_AllChartsToPPT()
Dim PPApp As Object
Dim PPpres As Object
Dim shObjC
Dim sh As Object 'sheets
Dim ch 'chartobject
Dim location
On Error Resume Next
Set location = ActiveWorkbook.ActiveSheet
For Each sh In ActiveWorkbook.Worksheets
    For Each ch In sh.ChartObjects
        shObjC = 1
     Next
    If shObjC = 1 Then Exit For
Next

If ActiveWorkbook.Charts.count = 0 And shObjC = 0 Then
    MsgBox ("No Charts Found")
    Exit Sub
End If
    
Application.ScreenUpdating = False
' Reference existing instance of PowerPoint
Set PPApp = CreateObject("Powerpoint.Application")
Set PPpres = PPApp.presentations.Add


For Each sh In Sheets

    Select Case sh.Type
        
        Case -4167 'if sheet type is a WORKSHEET then
        
        For Each ch In sh.ChartObjects
            ch.Chart.PlotVisibleOnly = False
            
            ch.Activate
            ch.Copy
            PasteToPPT False, Title:=sh.Name
            
        Next
        Case Else 'if sheet type is a CHART
            sh.PlotVisibleOnly = False
            sh.ChartArea.Copy
            PasteToPPT False, Title:=sh.Name
        'Case Else 'if neither a Worksheet or Chart, tell me what you are
            Debug.Print sh.Type & " : " & sh.Name
         
    End Select
    
Next


'Set view to First Sheet in workbook
'Workbooks(Target.Parent.Parent.Name).Worksheets(Target.Parent.Name).Range(Target).Select
On Error Resume Next
Excel.ActiveWorkbook.Activate
ActiveWorkbook.Activate
location.Select

' Clean up
pptFont ("Times New Roman")
Set PPpres = Nothing
Set PPApp = Nothing

End Sub

Sub pptPasteAllChartsheet()
On Error Resume Next

If ActiveWorkbook.Charts.count = 0 Then
    MsgBox ("No charts found in this workbook")
    Exit Sub
End If

Dim PPApp, PPpres As Object
Set PPApp = CreateObject("Powerpoint.Application")
Set PPpres = PPApp.presentations.Add
    Dim ch As Object
    For Each ch In ActiveWorkbook.Charts
        ch.Chart.PlotVisibleOnly = False
        ch.ChartArea.Copy
        PasteToPPT Title:=ch.Name
    Next
    pptFont ("Times New Roman")
End Sub


Sub pptPasteCurrentCharts()
On Error Resume Next
Dim PPApp, PPpres, ch As Object

Dim sh As Object

'''''''''
'look for charts if no charts then exit routine
Set sh = ActiveSheet.ChartArea
If sh Is Nothing And ActiveSheet.ChartObjects.count = 0 Then
    MsgBox ("No Charts Found on this sheet")
    Exit Sub
End If

'''''''''
'Start Routine: Create presentation
Set PPApp = CreateObject("Powerpoint.Application")
Set PPpres = PPApp.presentations.Add
    Select Case ActiveSheet.Type
        Case -4167
            For Each ch In ActiveSheet.ChartObjects
                ch.Chart.PlotVisibleOnly = False
                ch.Activate
                ch.Copy
                PasteToPPT Title:=ActiveSheet.Name
            Next
        Case Else
            ActiveSheet.ChartArea.Copy
            PasteToPPT Title:=ActiveSheet.Name
    
    End Select
    pptFont ("Times New Roman")
End Sub



Sub PasteToPPT(Optional NewPresentation As Boolean, Optional Title As String)
If IsMissing(NewPresentation) Then NewPresentation = True
'Dim XLApp As Object
Dim PPApp As Object
Dim PPpres As Object
Dim SlideCount As Integer
Dim PPSlide As Object
Dim PPchart As Object
'Set XLApp = GetObject(, Excel.Application)
On Error Resume Next
    Set PPApp = GetObject(, "Powerpoint.application")
    If PPApp Is Nothing Then Set PPApp = CreateObject("Powerpoint.Application")
On Error GoTo 0
On Error Resume Next
    If NewPresentation Then
        Set PPpres = PPApp.presentations.Add
    Else
        Set PPpres = PPApp.activepresentation
        If PPpres Is Nothing Then Set PPpres = PPApp.presentations.Add
    End If
On Error GoTo 0

SlideCount = PPpres.Slides.count
Set PPSlide = PPpres.Slides.Add(SlideCount + 1, 11)
PPApp.ActiveWindow.View.GotoSlide PPSlide.SlideIndex
With PPSlide
    ' paste and select the chart picture
    Set PPchart = .Shapes.Paste
    
    
    'PPchart.Select
    PPchart.Height = 5.5 * 72
    PPchart.Width = 8.6 * 72
    PPchart.Top = 120
    ' align the chart
    PPchart.Align msoAlignCenters, True
    'PPchart.Align msoAlignMiddles, True
    'PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
    'PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True

    .Shapes.Placeholders(1).TextFrame.TextRange.text = Title '"sTitle"
    .Shapes.Placeholders(1).TextFrame.TextRange.Font.Name = "Times New Roman" '"sTitle"
    pptFont ("Times New Roman")
End With

End Sub

Sub pptFont(Optional fName)
On Error Resume Next
If IsMissing(fName) Then fName = "Times New Roman"
    Dim fnt
    Dim PPApp As Object
    Set PPApp = GetObject(, "Powerpoint.Application")
    For Each fnt In PPApp.activepresentation.Fonts
        fnt.Name = fName '"Times New Roman"
    Next
End Sub

