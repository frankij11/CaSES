Attribute VB_Name = "M_CellFormatter"
Option Explicit

Private Function HasDependents(rngCheck As Range) As Boolean
   Dim lngSheetCounter As Long, lngRefCounter As Long, rngDep As Range
   On Error Resume Next

   With rngCheck
      .ShowDependents False
      Set rngDep = .NavigateArrow(False, 1, 1)
      If rngDep.Address(external:=True) = rngCheck.Address(external:=True) Then
         HasDependents = False
      Else
         HasDependents = (ERR.Number = 0)
      End If
      .ShowDependents True
   End With
   rngCheck.Select

End Function

Sub printGuide()
Dim guide As Worksheet
Application.ScreenUpdating = False
On Error Resume Next


Set guide = ActiveWorkbook.Worksheets("Format Guide")
If guide Is Nothing Then
    Set guide = ActiveWorkbook.Worksheets.Add
    guide.Name = "Format Guide"
End If
guide.Select
guide.Cells(2, 2) = "Text"
guide.Cells(3, 2) = "Constants"
guide.Cells(4, 2) = "Formulas"
guide.Cells(5, 2) = "Constants (No Dependents)"
guide.Cells(6, 2) = "Formulas (No Dependents)"

guide.Cells(2, 3) = "CY17$M"
guide.Cells(3, 3) = 20
guide.Cells(3, 4) = "This type of cell HAS dependents and is used in some calculations. You should focus on these values"
guide.Range(Cells(3, 4), Cells(3, 14)).Merge (True)

Dim i
For i = 0 To 11
    guide.Cells(4, 3).Offset(, i) = "=round(C3 * rand(),2)"
    If i = 6 Then guide.Cells(4, 3).Offset(, i) = 25
Next
Range("I4").AddComment ("Notice that this cell is hardcoded? Formatting allows for easy inspection")
Range("I4").Comment.Visible = True
Range("i4").Comment.Shape.Left = Range("P4").Left
Range("i4").Comment.Shape.Top = Range("P4").Top

guide.Cells(5, 3) = 15
guide.Cells(5, 4) = "This type of cell has no dependents and is not used in any calculations. You can likely ignore these values"
guide.Range(Cells(5, 4), Cells(5, 14)).Merge (True)
guide.Cells(6, 3) = "=sum(C4:N4)"
guide.Cells(6, 4) = "These type of cell represents either a Summary or Formula that is not used"
guide.Range(Cells(6, 4), Cells(6, 14)).Merge (True)

Columns("B:B").EntireColumn.AutoFit
Columns("A:A").ColumnWidth = 1.25
ActiveWindow.Zoom = 85
Range("A1").Select
Application.ScreenUpdating = True

guide.Cells.Interior.ColorIndex = 0
guide.Cells.Borders.LineStyle = xlNone
guide.Range("B2:C2, B3:N6").Borders.LineStyle = xlContinuous
guide.Range("B2:B6, C2, D3,D5,D6").Interior.ColorIndex = 15 'gray
MsgBox ("It can be difficult to review worksheets without formatting as this sheet shows. Take a second to see how unorganized this sheet is before continuing.")

formatWorksheet
MsgBox ("Now I will color cells that DO NOT have dependents a different color." & vbNewLine & vbNewLine & "Hardcoded (inputs): cells that have no dependents are likely not used in the model. Coloring them black shows they are not used in calculations." & vbNewLine & vbNewLine & "Formulas with out depedendents can either be a Summary or not used at all. Coloring them dark blue will help to identify them.")
colorCellsWithOutDependents




End Sub

Sub colorCellsWithOutDependents()
Dim location
Set location = Selection
On Error Resume Next
    Dim rng1, rng2 As Range
    Set rng1 = Cells.SpecialCells(xlCellTypeConstants, 1)
    Set rng2 = Cells.SpecialCells(xlCellTypeFormulas, 1)
    Dim tot1, tot2
    tot1 = rng1.count
    tot2 = rng2.count
    Dim continue
    If tot1 + tot2 > 500 Then continue = MsgBox(tot1 & ": Hardcoded cells found." & vbNewLine & tot2 & ": Formulas found" & vbNewLine & "The macro may take longer than normal." & vbNewLine & vbNewLine & "Would you like to contiue?", vbYesNo)
    
    If continue <> 7 Then
    
    Dim MainBar As ProgressBar
    Set MainBar = New ProgressBar
    MainBar.ShowBar "Main Bar: Coloring Cells..."
    MainBar.Top = MainBar.Top + MainBar.Height
    MainBar.progress 1 / 2
    colorDependentsConst
    MainBar.progress 2 / 2
    colorDependentsForm
    
    MainBar.Terminate (2)
    End If
location.Select
End Sub

Sub colorDependentsConst()
On Error Resume Next
    Dim rng As Range
    Dim cel As Range
    Dim i, pct, tot As Integer
    Dim subBar As ProgressBar
    Set subBar = New ProgressBar
    subBar.ShowBar "Coloring Constant Cells"
    Set rng = Cells.SpecialCells(xlCellTypeConstants, 1)
    tot = rng.count
    For Each cel In rng
        i = i + 1
        pct = Round(i / tot, 2)
        subBar.progress pct
        If Not HasDependents(cel) Then
            cel.Interior.ColorIndex = 1
            cel.Font.ColorIndex = 2
        End If
    
    Next
    subBar.Terminate (1)
End Sub

Sub colorDependentsForm()
On Error Resume Next
    Dim rng As Range
    Dim cel As Range
    Dim i, pct, tot As Integer
    Dim subBar As ProgressBar
    Set subBar = New ProgressBar
    subBar.ShowBar "Coloring Constant Cells"
    Set rng = Cells.SpecialCells(xlCellTypeFormulas, 1)
    tot = rng.count
    For Each cel In rng
        
        i = i + 1
        pct = Round(i / tot, 2)
        subBar.progress pct
        If Not HasDependents(cel) Then
            cel.Interior.ColorIndex = 5 'RGB(0, 0, 0)
            cel.Font.ColorIndex = 2
            cel.Font.Bold = True
        End If
    
    Next
    subBar.Terminate (1)
End Sub

Sub formatEntireWorkbook()
Application.ScreenUpdating = False
On Error Resume Next
Dim Target As Range
Set Target = Selection
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Select
        formatWorksheet
    Next
Application.ScreenUpdating = True
Workbooks(Target.Parent.Parent.Name).Activate
Sheets(Target.Parent.Name).Activate
Target.Select

End Sub

Sub formatWorksheet()
On Error Resume Next
Dim Target As Range
Set Target = Selection
    'fConstText (14540253)   'light grey
    'fFormText (14540253)    'light grey
    fConstNum (13434879)    'light yellow
    fFormNum (16772300)     'light blue
    'colorDependentsConst
    'colorDependentsForm
    
    'linked to a specific tab
    
Target.Select
End Sub

Private Sub fBlanks(Optional i_color, Optional i_font, Optional i_bold, Optional i_pattern)
ActiveWindow.DisplayGridlines = False
If IsMissing(i_color) Then
    MsgBox ("What color should blank cells be?" & vbNewLine & "Choose on next screen")
    i_color = PickNewColor(Selection.Interior.Color)
End If

If IsMissing(i_font) Then i_font = "Times New Roman"
If IsMissing(i_bold) Then i_bold = False
If IsMissing(i_pattern) Then i_pattern = xlSolid

Dim formatRng As Range
'Cells.Select
''''Format Blank Cells'''''''''''''''''''''
On Error GoTo ERR:
    Set formatRng = Cells.SpecialCells(xlCellTypeBlanks)

    With formatRng.Interior
        .Color = i_color
    
    End With
    

Exit Sub


''''Format Blank Cells'''''''''''''''''''''
ERR:
MsgBox ("An error occurred")
 End Sub



    
''''Format Constants Numbers Cells'''''''''''''''''''''
    

''''Format Constants Texts Cells'''''''''''''''''''''
Private Sub fConstText(Optional i_color, Optional i_font, Optional i_bold, Optional i_pattern, Optional i_font_color)
ActiveWindow.DisplayGridlines = False

If IsMissing(i_color) Then
    MsgBox ("What color should cells containing Text be?" & vbNewLine & "Choose cell interior color on next screen")
    i_color = PickNewColor(Selection.Interior.Color)
    Debug.Print i_color
End If

If IsMissing(i_font) Then i_font = "Times New Roman"
If IsMissing(i_bold) Then i_bold = False
If IsMissing(i_pattern) Then i_pattern = xlSolid
If IsMissing(i_font_color) Then i_font_color = 0
   
Dim formatRng As Range
   On Error GoTo ERR:
   Set formatRng = Cells.SpecialCells(xlCellTypeConstants, 2)
    With formatRng
 
        .Interior.Color = i_color

    End With
    
 
''''Format Constants Text Cells'''''''''''''''''''''
Exit Sub
ERR:
End Sub

Private Sub fConstNum(Optional i_color, Optional i_font, Optional i_bold, Optional i_pattern, Optional i_font_color)
ActiveWindow.DisplayGridlines = False

If IsMissing(i_color) Then
    MsgBox ("What color should cells containing Text be?" & vbNewLine & "Choose cell interior color on next screen")
    i_color = PickNewColor(Selection.Interior.Color)
    
End If

If IsMissing(i_font) Then i_font = "Times New Roman"
If IsMissing(i_bold) Then i_bold = False
If IsMissing(i_pattern) Then i_pattern = xlSolid
If IsMissing(i_font_color) Then i_font_color = 0
   
Dim formatRng As Range
    On Error GoTo ERR:
    Set formatRng = Cells.SpecialCells(xlCellTypeConstants, 1)
    With formatRng

        .Interior.Color = i_color

    End With
    

''''Format Constants Text Cells'''''''''''''''''''''
Exit Sub
ERR:
End Sub

Private Sub fFormText(Optional i_color, Optional i_font, Optional i_bold, Optional i_pattern, Optional i_font_color)
ActiveWindow.DisplayGridlines = False
On Error GoTo ERR:
If IsMissing(i_color) Then
    MsgBox ("What color should cells containing a formula with Text be?" & vbNewLine & "Choose cell interior color on next screen")
    i_color = PickNewColor(Selection.Interior.Color)
End If

If IsMissing(i_font) Then i_font = "Times New Roman"
If IsMissing(i_bold) Then i_bold = True
If IsMissing(i_pattern) Then i_pattern = xlSolid
If IsMissing(i_font_color) Then i_font_color = 0

Dim formatRng As Range
Set formatRng = Cells.SpecialCells(xlCellTypeFormulas, 6)

 With formatRng
        .Interior.Color = i_color

 End With
    

''''Format Constants Text Cells'''''''''''''''''''''
Exit Sub
ERR:


End Sub

Private Sub fFormNum(Optional i_color, Optional i_font, Optional i_bold, Optional i_pattern, Optional i_font_color)
ActiveWindow.DisplayGridlines = False
If IsMissing(i_color) Then
    MsgBox ("What color should cells containing a Formula be?" & vbNewLine & "Choose cell interior color on next screen")
    i_color = PickNewColor(Selection.Interior.Color)
    Debug.Print i_color
End If

If IsMissing(i_font) Then i_font = "Times New Roman"
If IsMissing(i_bold) Then i_bold = False
If IsMissing(i_pattern) Then i_pattern = xlSolid
If IsMissing(i_font_color) Then i_font_color = 0
  
  ''''Format Formula Numbers Cells'''''''''''''''''''''
Dim rng As Range
    'Range("A1").Select
On Error GoTo ERR:
   Set rng = Cells.SpecialCells(xlCellTypeFormulas, 1)
   With rng
        .Interior.Color = i_color

    End With
    

Exit Sub
ERR:
End Sub







'Picks new color
Private Function PickNewColor(Optional i_OldColor As Double = xlNone) As Double
Const BGColor As Long = 13160660  'background color of dialogue
Const ColorIndexLast As Long = 32 'index of last custom color in palette

Dim myOrgColor As Double          'original color of color index 32
Dim myNewColor As Double          'color that was picked in the dialogue
Dim myRGB_R As Integer            'RGB values of the color that will be
Dim myRGB_G As Integer            'displayed in the dialogue as
Dim myRGB_B As Integer            '"Current" color
  
  'save original palette color, because we don't really want to change it
  myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
  
  If i_OldColor = xlNone Then
    'get RGB values of background color, so the "Current" color looks empty
    Color2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
  Else
    'get RGB values of i_OldColor
    Color2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
  End If
  
  'call the color picker dialogue
  If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, _
     myRGB_R, myRGB_G, myRGB_B) = True Then
    '"OK" was pressed, so Excel automatically changed the palette
    'read the new color from the palette
    PickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
    'reset palette color to its original value
    ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
  Else
    '"Cancel" was pressed, palette wasn't changed
    'return old color (or xlNone if no color was passed to the function)
    PickNewColor = i_OldColor
  End If
End Function

'Converts a color to RGB values
Sub Color2RGB(ByVal i_color As Long, _
              o_R As Integer, o_G As Integer, o_B As Integer)
  o_R = i_color Mod 256
  i_color = i_color \ 256
  o_G = i_color Mod 256
  i_color = i_color \ 256
  o_B = i_color Mod 256
End Sub





