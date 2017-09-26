Attribute VB_Name = "M_PivotTables"
Option Explicit
Public Sub PivotFieldsToSum()
' Cycles through all pivot data fields and sets to sum
' Created by Dr Moxie
On Error Resume Next
    Dim pf As PivotField
    With Selection.PivotTable
        .ManualUpdate = True
        For Each pf In .DataFields
            With pf
                .Function = xlSum
                .NumberFormat = "$#,##0"
            End With
        Next pf
        .ManualUpdate = False
    End With
End Sub


