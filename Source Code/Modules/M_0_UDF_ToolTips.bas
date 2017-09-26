Attribute VB_Name = "M_0_UDF_ToolTips"
Option Explicit
Sub LoadUDF_ToolTips()


Application.ScreenUpdating = False
ThisWorkbook.IsAddin = False

Dim udf
Dim desc As String
Dim ArgDesc()
Dim ArgDesc1 As Range
Dim count
For Each udf In ThisWorkbook.Worksheets("VBA Functions").Range("UDFs[Macro]")
    count = 0
    desc = udf.Offset(, 2).Value
    Set ArgDesc1 = ThisWorkbook.Worksheets("VBA Functions").Range(udf.Offset(, 5).Address)
    Do Until ArgDesc1 = ""
        
        ReDim Preserve ArgDesc(0 To count)
        ArgDesc(count) = ArgDesc1.Value
        Set ArgDesc1 = ArgDesc1.Offset(, 1)
        count = count + 1
    Loop
    Call toolTips(udf.Value, desc, ArgDesc)
    ReDim ArgDesc(0 To 0)

Next


ThisWorkbook.IsAddin = True
Application.ScreenUpdating = True
End Sub

Private Sub toolTips(udf As String, Descript As String, ArgDesc)
On Error Resume Next

Application.MacroOptions Macro:=udf, _
    Description:=Descript, _
    ArgumentDescriptions:=ArgDesc, _
    Category:="CaSES"
    
End Sub
