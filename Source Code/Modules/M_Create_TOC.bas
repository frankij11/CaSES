Attribute VB_Name = "M_Create_TOC"
Option Explicit
'Referenced from http://www.vbaexpress.com/kb/getarticle.php?kb_id=120
 
Sub TEST_CreateTOC1()
    Call CreateTOC(False, False)
End Sub
 
Sub TEST_CreateTOC2()
    Call CreateTOC(True, True)
End Sub
 
Sub TEST_CreateTOC3()
    Call CreateTOC(False, True)
End Sub
 
Sub TEST_CreateTOC4()
    Call CreateTOC(True, False)
End Sub
 
Sub CreateTOC(Optional ByVal IncludeHiddenSheets As Boolean = False, _
    Optional ByVal AddHomeLinkOnSheets As Boolean = False)
     'Referenced from http://www.vbaexpress.com/kb/getarticle.php?kb_id=120
     ' IncludeHiddenSheets
     '   Boolean
     '   Specifies whether or not hidden sheets should be included in the Table of Contents
     '
     ' AddHomeLinkOnSheets
     '   Boolean
     '   Specifies whether or not a link should be placed in each sheet linking back to the
     '   Table of Contents. This will only be placed on worksheets (i.e. not chart sheets),
     '   will not work with a protected sheet, and will overwrite anything in the cell
     '   specified in the destination [address] constant below (under declared variables).
     '
     'Use cases:
     'Call CreateTOC(False, False)
     '   This will create a Table of Contents which excludes hidden sheets and does not add a link
     '   back to itself
     '
     'Call CreateTOC(True, True)
     '   This will create a Table of Contents which includes hidden sheets and also includes a link
     '   back to itself.
     '*** CAUTION: Specifying a cell in each sheet will 1) only work on worksheets (i.e. not chart sheets),
     '               overwrite anything in the destination cell (unless worksheet is protected)
     '
     'Call CreateTOC(False, True)
     '   This will create a Table of Contents which excludes hidden sheets and also includes a link
     '   back to itself.
     '*** CAUTION: Specifying a cell in each sheet will 1) only work on worksheets (i.e. not chart sheets),
     '               overwrite anything in the destination cell (unless worksheet is protected)
     '
     'Call CreateTOC(True, False)
     '   This will create a Table of Contents which includes hidden sheets and does not add a link
     '   back to itself
     '
     'Declare all variables
    Dim TOCBook As Workbook
    Dim CheckSheet As Worksheet
    Dim TOC As Worksheet
    Dim ChartButton As Shape
    Dim NewRow As Long
    Dim SheetCount As Long
    Dim CellLeft
    Dim CellTop
    Dim CellHeight
    Dim CellWidth
    Dim SheetName As String
    Dim Prompt As String
    Dim CellR1C1Address As String
     
     'Set a constant to the name of the Table of Contents
    Const TOCName As String = "TOC"
    Const HomeCell As String = "A1"
    Const StartRow As Long = 5
     
     'Check if a workbook is open or not.  If no workbook is open, quit.
    If ActiveWorkbook Is Nothing Then
        MsgBox "You must have a workbook open first!", vbInformation, "No Open Book"
        Exit Sub
    End If
    Set TOCBook = ActiveWorkbook
     
    On Error Resume Next
    Set TOC = TOCBook.Worksheets("TOC")
    On Error GoTo 0
    If Not TOC Is Nothing Then
        If MsgBox("Table of contents already exists. Overwrite?", vbYesNo + vbDefaultButton2, "Overwrite TOC?") <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        TOC.Delete
        Set TOC = Nothing
    End If
    Set TOC = TOCBook.Worksheets.Add(before:=TOCBook.Sheets(1))
    TOC.Name = TOCName
    TOC.Columns(1).ColumnWidth = 1
     
    TOC.Cells(StartRow - 3, "B").Value = "TABLE OF CONTENTS"
    If IncludeHiddenSheets Then
        TOC.Cells(StartRow - 2, "B").Value = "Hidden sheets are italicized"
        TOC.Cells(StartRow - 2, "B").Font.Size = 10
        NewRow = StartRow
    Else
        NewRow = StartRow - 1
    End If
     
    For SheetCount = 1 To TOCBook.Sheets.count
        SheetName = TOCBook.Sheets(SheetCount).Name
        If TOCBook.Sheets(SheetName).Name = TOCName Then GoTo SkipSheet
        If Not IncludeHiddenSheets And TOCBook.Sheets(SheetName).Visible <> xlSheetVisible Then GoTo SkipSheet
        If IsChart(SheetName) Then
             '** Sheet IS a Chart Sheet
             'Set variables for button dimensions.
            CellLeft = TOC.Range("B" & NewRow).Left
            CellTop = TOC.Range("B" & NewRow).Top
            CellWidth = TOC.Range("B" & NewRow).Width
            CellHeight = TOC.Range("B" & NewRow).RowHeight
            CellR1C1Address = "R" & NewRow & "C3"
             'Add button to cell dimensions.
            Set ChartButton = TOC.Shapes.AddShape(msoShapeRoundedRectangle, CellLeft, CellTop, CellWidth, CellHeight)
            ChartButton.Select
             'Use older technique to add Chart sheet name to button text.
            ExecuteExcel4Macro "FORMULA(""=" & CellR1C1Address & """)"
             'Format shape to look like hyperlink and match background color (transparent).
            Selection.ShapeRange.Fill.ForeColor.SchemeColor = 0
            Selection.Font.Underline = xlUnderlineStyleSingle
            Selection.Font.ColorIndex = 0
            Selection.ShapeRange.Fill.Visible = msoFalse
            Selection.ShapeRange.Line.Visible = msoFalse
            Selection.OnAction = "GotoChart"
            Selection.Name = SheetName
        Else
             '** Sheet is NOT a Chart sheet. Add a hyperlink to A1 of each sheet.
            TOC.Range("B" & NewRow).Hyperlinks.Add Anchor:=TOC.Range("B" & NewRow), Address:="#'" & SheetName & "'!A1", TextToDisplay:=SheetName
            If AddHomeLinkOnSheets Then
                If TOCBook.Sheets(SheetName).Type = xlWorksheet Then
                    If TOCBook.Sheets(SheetName).ProtectContents = False Then
                        If TOCBook.Sheets(SheetName).Range(HomeCell).Value <> "" And TOCBook.Sheets(SheetName).Range(HomeCell).Value <> "TOC" Then
                            TOCBook.Sheets(SheetName).Range(HomeCell).EntireRow.Insert
                        End If
                        TOCBook.Sheets(SheetName).Range(HomeCell).Value = "TOC"
                        TOCBook.Sheets(SheetName).Range(HomeCell).Hyperlinks.Add Anchor:=TOCBook.Sheets(SheetName).Range("A1"), Address:="#'" & TOCName & "'!A1", TextToDisplay:=TOCName
                    End If
                End If
            End If
        End If
         'Add name and format sheet name on TOC
        TOC.Range("B" & NewRow).Value = SheetName
        TOC.Range("B" & NewRow).HorizontalAlignment = xlLeft
        TOC.Range("B" & NewRow).Font.Italic = CBool(TOCBook.Sheets(SheetName).Visible <> xlSheetVisible)
        TOC.Range("B" & NewRow).Font.ColorIndex = 5
         'Increment row
        NewRow = NewRow + 1
SkipSheet:
    Next SheetCount
     
    TOC.Activate
    TOC.Cells(1, 1).Select
    ActiveWindow.DisplayGridlines = False
End Sub
 
Public Function IsChart(cName As String, Optional ChartBook As Workbook) As Boolean
     
     'Will return True or False if sheet is a Chart sheet object or not.
     'Can be used as a worksheet function.
    Dim tmpChart As Chart
    If ChartBook Is Nothing Then
        If ActiveWorkbook Is Nothing Then Exit Function
        Set ChartBook = ActiveWorkbook
    End If
     
     'Function will be determined if the object is not errored
    On Error Resume Next
    IsChart = IIf(ChartBook.Charts(cName) Is Nothing, False, True)
    On Error GoTo 0
     
End Function
 
Sub GotoChart(Optional Placebo As String = "")
     
     'This routine is to be assigned to button Object for Chart sheets only
     'as Chart sheets don't have cell references to hyperlink to.
     
    On Error Resume Next
    ActiveWorkbook.Charts(Application.Caller).Activate
    On Error GoTo 0
    If ERR.Number <> 0 Then Exit Sub
     
     'Optional: zoom Chart sheet to fit screen.
     'Depending on screen resolution, this may need adjustment(s).
    ActiveWindow.Zoom = 80
     
End Sub

