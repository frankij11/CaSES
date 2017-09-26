Attribute VB_Name = "M_Add_WBS_Sheets"
Option Explicit
Option Base 1

Public Sub M_WBSElements_To_Tabs()
    
    TabstoSheets_WBS_Option.Show

End Sub


Sub createWBSTabs(Concat As Boolean)

Dim location
Set location = ActiveSheet

'This module will allow users to select a group of WBS elements and then create tabs for each one

Dim wbsRange As Range

'This section will set the size of WBS Element array to number of rows selected by user

Set wbsRange = Application.InputBox("Choose range that includes WBS Elements", _
        Title:="WBS Element Location", _
        Type:=8)
' modify wbsRange to be only the first column
Set wbsRange = Range(wbsRange(1, 1), wbsRange(wbsRange.Rows.count, 1))

' modify wbsRange to be visible cells only
If wbsRange.count <> 1 Then
    Set wbsRange = wbsRange.SpecialCells(xlCellTypeVisible)
    
End If

'if too many cells are selected Allow user to abort
If wbsRange.count > 75 Then
    MsgBox (wbsRange.count & " WBS Elements selected. Please choose less than 50 WBS elements and try again")
    Exit Sub
ElseIf wbsRange.count > 25 Then
    
    Dim cont
    cont = MsgBox(wbsRange.count & " WBS Elements have been selected." & vbNewLine & vbNewLine & "Would you like to continue?", vbOKCancel)
    If cont = vbCancel Then
        MsgBox ("Please select fewer WBS elements and try again")
        Exit Sub
    End If
End If
    

'loop through each WBS name in WBS range
'Create generic calc with WBS name
Dim wbsCell
Dim wbsName

For Each wbsCell In wbsRange
    If wbsCell.Value <> "" Then
        'make
        If Concat Then
            wbsName = CleanWorksheetName(VBA.Trim(wbsCell.Value) & " " & VBA.Trim(wbsCell.Offset(, 1).Value))
        Else
           wbsName = CleanWorksheetName(VBA.Trim(wbsCell.Value))
        End If
        
        addGenericCalc (wbsName)
    End If
Next

location.Activate
    
End Sub

Function CleanWorksheetName(ByRef strName As String) As String
    Dim varBadChars As Variant
    Dim varChar As Variant
    'remove extra spaces
    Dim c As Integer 'count variable
    Do While InStr(1, strName, "  ") > 0
        c = c + 1
        strName = Replace(strName, "  ", " ")
        If c > 20 Then Exit Do
    Loop
    varBadChars = Array(":", "/", "\", "?", "*", "[", "]")
     
     'correct string for forbidden characters
    For Each varChar In varBadChars
        Select Case varChar
        Case ":"
            strName = Replace(strName, varChar, vbNullString)
        Case "/"
            strName = Replace(strName, varChar, "-")
        Case "\"
            strName = Replace(strName, varChar, "-")
        Case "?"
            strName = Replace(strName, varChar, vbNullString)
        Case "*"
            strName = Replace(strName, varChar, vbNullString)
        Case "["
            strName = Replace(strName, varChar, "(")
        Case "]"
            strName = Replace(strName, varChar, ")")
        End Select
    Next varChar
     
     'correct string for worksheet length requirement
    strName = VBA.Left(strName, 31)
     
    CleanWorksheetName = strName
End Function
