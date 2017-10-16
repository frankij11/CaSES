VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TraceBackNavigator 
   Caption         =   "Traceback Navigator"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11205
   OleObjectBlob   =   "TraceBackNavigator.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TraceBackNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Userform1_initialize()
TraceBackNavigator.Show

End Sub

Private Sub CheckBox1_Click()
'''Need to add logic back in minimze screen
End Sub

Private Sub CommandButton1_Click()
Dim i
ListBox5.Clear
    With ActiveWorkbook
        For i = 1 To .Worksheets.count
            If .Worksheets(i).Visible = True Then ListBox5.AddItem .Worksheets(i).Name
        Next i
    End With
UserForm_Click

End Sub



Private Sub CommandButton2_Click()
ListFormulas

End Sub

Private Sub CommandButton3_Click()
Application.Dialogs(xlDialogEvaluateFormula).Show


End Sub

Private Sub CommandButton4_Click()
FormulaFormatter_Click
End Sub

Private Sub ListBox1_Click()

Dim addr, x
Dim ws As Worksheet
Dim rng As Range
Dim M
'On Error Resume Next
M = ListBox4.Value

addr = Split(ListBox1.Value, "!")
Set ws = Worksheets(Replace(addr(0), "'", ""))
ws.Select
Set rng = ws.Range(addr(1))
rng.Select
ListBox1.Selected(ListBox1.ListIndex) = -1

'UserForm_Click
End Sub



Private Sub ListBox4_Click()
Dim addr, x
Dim ws As Worksheet
Dim rng As Range
Dim M
'On Error Resume Next
M = ListBox4.Value

addr = Split(ListBox4.Value, "!")
Set ws = Worksheets(Replace(addr(0), "'", ""))
ws.Select
Set rng = ws.Range(addr(1))
rng.Select
ListBox4.ListIndex = -1


'UserForm_Click
End Sub
Private Sub ListBox6_Click()
Dim x
Dim addr
Dim ws As Worksheet
Dim rng As Range
Dim M
'On Error Resume Next
M = ListBox6.Value
addr = Split(ListBox6.Value, "!")
Set ws = Worksheets(Replace(addr(0), "'", ""))
ws.Select
Set rng = ws.Range(addr(1))
rng.Select
ListBox6.ListIndex = -1
'UserForm_Click
End Sub

Private Sub ListBox5_Click()
On Error Resume Next
Worksheets(ListBox5.Value).Select
End Sub





Private Sub UserForm_Activate()
    Dim i
    'MakeFormResizable
    'Me.TreeView1.LineStyle = tvwRootLines
    'Tre
    ListBox5.Clear
    With ActiveWorkbook
        For i = 1 To .Worksheets.count
            If .Worksheets(i).Visible = True Then ListBox5.AddItem .Worksheets(i).Name
        Next i
    End With
    With Me
        'This will create a vertical scrollbar
        '.ScrollBars = fmScrollBarsVertical
        
        'Change the values of 2 as Per your requirements
        .ScrollHeight = .InsideHeight * 2
        .ScrollWidth = .InsideWidth * 1.5
    End With
End Sub

Private Sub UserForm_Click()
Dim i, lb4add, lb4txt, tar, wks, tempp2, a, LB1
'For c = ListBox4.ListCount - 1 To 0 Step -1
'    If ListBox4.Selected(c) = VBA.Trim(Replace(ActiveCell.Address(, , , True), "[" & ActiveWorkbook.Name & "]", "")) Then GoTo 22
'
'Next
ListBox5.Clear
    With ActiveWorkbook
        For i = 1 To .Worksheets.count
            If .Worksheets(i).Visible = True Then ListBox5.AddItem .Worksheets(i).Name
        Next i
    End With




lb4txt = VBA.Trim(Replace(ActiveCell.Address(, , , True), "[" & ActiveWorkbook.Name & "]", ""))
'MsgBox (ListBox4.ListCount)
lb4add = True
For i = 0 To ListBox4.ListCount - 1
   
    If ListBox4.List(i, 0) = lb4txt Then
    ListBox4.RemoveItem (i)
    GoTo addLable
    End If
Next
addLable:
If lb4add = True Then ListBox4.AddItem lb4txt

    
TextBox1 = ActiveCell.Formula
TextBox2 = Replace(TextBox1.Value, "$", "")
ListBox1.Clear
ListBox2.Clear
ListBox3.Clear
ListBox6.Clear
Dim tempp()
tar = ActiveCell.Address(, , , True)
Set tar = Range(tar)
wks = tar.Parent.Name
    wks = "'" & wks & "'!"
    Debug.Print (wks)
tempp = tracePrecedents()
tempp2 = traceDependents()
    a = UBound(tempp, 1)

On Error Resume Next 'GoTo ErrorHand
For i = UBound(tempp, 1) To LBound(tempp, 1) Step -1
        
    
    'wbk = "[NCCA_GATOR MODEL v45.xlsx]"
    
    LB1 = Replace(tempp(i), wks, "")
    Debug.Print ("LB1: " & LB1)
  
    TextBox2 = Replace(TextBox2.Value, Replace(LB1, "$", ""), WorksheetFunction.Sum(Range(tempp(i))))
    TextBox2 = Replace(TextBox2.Value, Replace(LB1, "$", ""), Range(tempp(i)).Value)
    'TextBox2 = Replace(TextBox2.Value, Replace(LB1, "$", ""), Range(tempp(i)).Value2)
            'TextBox2 = Replace(TextBox1, tempp(i), range(tempp(i), , 1)
    With ListBox1
        '.AddItem Replace(tempp(i), wks, "")
        .AddItem tempp(i)
    End With
    With ListBox2
        If Range(tempp(i)).NumberFormat <> "General" Then
        .AddItem VBA.Format(Range(tempp(i)).Formula, Range(tempp(i)).NumberFormat)
        Else
        .AddItem VBA.Format(Range(tempp(i)).Formula, "#,##0.##")
        End If
        
    End With
    With ListBox3
       
        If Range(tempp(i)).NumberFormat <> "General" Then
        .AddItem VBA.Format(Range(tempp(i)).Value2, Range(tempp(i)).NumberFormat)
        Else
        .AddItem VBA.Format(Range(tempp(i)).Value, "#,##0.##")
        End If
       
       
    End With
Next i

For i = 1 To UBound(tempp2, 1)
    With ListBox6
        '.AddItem Replace(tempp2(i), wks, "")
        .AddItem tempp2(i)
        Debug.Print (tempp2(i))
    End With
Next i

Exit Sub
ERR:
    ReDim Preserve tempp(1 To 1, 1 To 1)
ErrorHand:
End Sub
