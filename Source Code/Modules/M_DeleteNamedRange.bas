Attribute VB_Name = "M_DeleteNamedRange"
Option Explicit
Sub M_Delete_NamedRange()

Dim ExName As Name
Dim continue
On Error Resume Next

If ActiveWorkbook.Names.count = 0 Then

MsgBox ("No Named Ranges are included in Active Workbook")

Else

For Each ExName In ActiveWorkbook.Names
    If ExName.Visible = True Or InStr(1, ExName.Value, "#Ref") Then
        continue = MsgBox("Delete Named Range: " & vbNewLine & ExName.Name & " " & ExName, vbYesNoCancel)
        If continue = vbYes Then
            ExName.Visible = True
            ExName.Delete
        
        ElseIf continue = vbCancel Then
            Exit Sub
        
        End If
    End If
Next
    
    On Error GoTo 0
    
End If
        
End Sub
