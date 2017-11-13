Attribute VB_Name = "M_WBS"
Option Explicit

Sub wbsGroupInd(Optional wbsRange, Optional wbsType)
Dim c 'costant counter
Dim rng
Dim cell As Range
Dim ind 'number of cell indent
Dim wbsArray As Variant
Dim tooBig As String ' contains which elements are greater than level 8
Dim cella

On Error GoTo errHandle
If IsMissing(wbsType) Then wbsType = Application.InputBox(Prompt:="If WBS contains Periods Enter 0" & VBA.vbNewLine & "If WBS Contains Indent Enter 1", Type:=1)
If IsMissing(wbsRange) Then Set wbsRange = Application.InputBox(Prompt:="select cells that contains wbs", Type:=8)
    'Set rng2 = Application.InputBox(Prompt:="select cell that contains last wbs", Type:=8)
   
    
    'Range(rng1.Address, rng2.Address).ClearOutline
    'Selection.ClearOutline
    c = 0
    For Each cell In wbsRange
        'summ = "=sum("
        'summ = summ & cell.Address
        'Cells(cell.Row, cell.Column + 1) = summ & ")"
        
        Select Case wbsType
            Case 0
                If cell.Value <> "" Then
                    cella = Split(cell, " ")
                    cella = cella(0)
                    ind = Len(cella) - Len(Replace(cella, ".", ""))
                    wbsArray = Split(cella, ".")
                    If wbsArray(UBound(wbsArray, 1)) = 0 Then ind = ind - 1
                    cell.IndentLevel = ind
                Else
                    ' do nothing
                End If
            
            Case 1
                ind = cell.IndentLevel
        End Select
        
        If ind > 7 Then
            c = c + 1
            tooBig = tooBig & " " & cell.Value2
            cell.EntireRow.OutlineLevel = WorksheetFunction.Min(ind + 1, 8)
        Else
            'cell.NumberFormat = "@"
            cell.EntireRow.OutlineLevel = WorksheetFunction.Min(ind + 1, 8)
        End If
        
    Next cell

    If c > 0 Then MsgBox (c & " Items are too large to group" & tooBig)
On Error GoTo 0
On Error Resume Next
     With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
    End With
Exit Sub
errHandle:
End Sub

Sub sumWBS(Optional wbsRange, Optional sumRange, Optional Years, Optional RollupRange)
Dim last, k, kk, i, ii, ind, tempsubtotal, hmm, rngYR, RollUp
Dim nextWBS, WBS

'On Error Resume Next
On Error GoTo errHandle
If IsMissing(wbsRange) Then Set wbsRange = Application.InputBox("Select cells of wbs", Type:=8)

If IsMissing(sumRange) Then Set sumRange = Application.InputBox("First cell of Summation", Type:=8)
If IsMissing(Years) Then Years = Application.InputBox("Number of Columns to add (Be Sure to Add a column for Totals", Type:=1)
If IsMissing(RollupRange) Then RollUp = False Else RollUp = True
'last = r.End(xlDown).row - r.row + 1
'k = r.Column


kk = sumRange.Column

'Years = Application.InputBox("Number of Columns to add (Be Sure to Add a column for Totals", Type:=1)
For Each WBS In wbsRange 'i = r.row To r.End(xlDown).row

    ind = WBS.IndentLevel
    ii = i + 1
    Set nextWBS = WBS.Offset(1)
    tempsubtotal = ""
    
    If ind >= nextWBS.IndentLevel Then
        If RollUp Then
            If Cells(WBS.row, RollupRange.Column) = "ROLL UP" Then
                Cells(WBS.row, RollupRange.Column).ClearContents
                If Cells(WBS.row, kk).HasArray Then Cells(WBS.row, kk).CurrentArray.ClearContents
                Range(Cells(WBS.row, kk), Cells(WBS.row, kk + Years - 1)).ClearContents
            End If
        End If
    End If
    
    
    While ind < nextWBS.IndentLevel
    
        hmm = nextWBS.IndentLevel - ind
        If nextWBS.IndentLevel - ind = 1 Then
            tempsubtotal = tempsubtotal & Cells(nextWBS.row, kk).Address(False, False) & ", "
        End If
        Set nextWBS = nextWBS.Offset(1)
        ii = ii + 1
    Wend
    'newstring = "=subtotal(9," & Left(tempsubtotal, Len(tempsubtotal) - 2) & ")"
    If Not tempsubtotal = "" Then
        
        'delete array
        If Cells(WBS.row, kk).HasArray Then Cells(WBS.row, kk).CurrentArray.ClearContents
        'insert summing formula
        Cells(WBS.row, kk) = "=sum(" & VBA.Left(tempsubtotal, Len(tempsubtotal) - 2) & ")"
        Set rngYR = Range(Cells(WBS.row, kk), Cells(WBS.row, kk + Years - 1))
        Cells(WBS.row, kk).Copy (rngYR)
        If RollUp Then Cells(WBS.row, RollupRange.Column) = "ROLL UP"
    End If
Next
Exit Sub
errHandle:
'MsgBox ("Routine did not complete")
End Sub

Sub childWBS(Optional wbsRange, Optional childRange, Optional levRange, Optional parentLevel)
Dim ii, ind, noWBS, prevWBS
Dim nextWBS, WBS
If IsMissing(wbsRange) Then Set wbsRange = Application.InputBox("Select cells of wbs", Type:=8)
If IsMissing(levRange) Then Set levRange = Application.InputBox("Select destination for parent wbs", Type:=8)

If IsMissing(childRange) Then Set childRange = Application.InputBox("Select destination for parent child relationship", Type:=8)
If IsMissing(parentLevel) Then parentLevel = 3
    

For Each WBS In wbsRange 'i = r.row To r.End(xlDown).row
    Set prevWBS = WBS
    ind = WBS.IndentLevel
    'Debug.Print wbs.Value
    Set nextWBS = WBS.Offset(1)
    'Debug.Print nextWBS.IndentLevel
    'Debug.Print nextWBS.Address
    If ind >= nextWBS.IndentLevel Then
        Cells(WBS.row, childRange.Column) = "Child"
        
        Do

            Debug.Print prevWBS.Address
            Set prevWBS = prevWBS.Offset(-1)
            If prevWBS.IndentLevel = parentLevel - 1 Then
                Cells(WBS.row, levRange.Column) = prevWBS.Value
                Exit Do
            End If
        Loop
        
        
    Else
        Cells(WBS.row, childRange.Column) = "Parent"
    End If
   
Next
End Sub

Sub WBSNumbering(Optional rngDesc As range)
'From http://j.modjeska.us/?p=31'Renumber tasks on a project plan'Associate this code with a button or other control on your spreadsheet
' inserts WBS numbering to the left of selection/range
    On Error Resume Next        'Hide page breaks and disable screen updating (speeds up processing)        ActiveSheet.DisplayPageBreaks = False        Dim c As range, d As range, n As Long    Dim depth As Long               'How many "decimal" places for each task    Dim wbsArray() As Long          'Master array holds counters for each WBS level    Dim iBase As Long              'Whole number sequencing variable    Dim strWBS As String               ' WBS string    Dim iLoop As Integer            ' loop variable        If rngDesc Is Nothing Then Set rngDesc = Application.InputBox("Choose WBS", "WBS Number Generator", Type:=8)  'Selection    Application.ScreenUpdating = False    If rngDesc.Columns.Count > 1 Then Exit Sub        If rngDesc.Rows = ActiveSheet.Rows Then _        n = rngDesc.SpecialCells(xlCellTypeLastCell).Row        iBase = 0                     'Initialize whole numbers    ReDim wbsArray(0 To 0) As Long  'Initialize WBS ennumeration array        For Each c In rngDesc        'Ignore empty tasks        If c <> "" And c.EntireRow.Hidden = False Then            depth = c.IndentLevel                        If 0 = depth Then                ' this is a primary task                iBase = iBase + 1                strWBS = CStr(iBase)                ReDim wbsArray(0 To 0)            Else                ' depth > 0                ReDim Preserve wbsArray(0 To depth) As Long                                ' zero based indexing on wbsArray                depth = depth - 1                                ' add one to WBS element at this depth                wbsArray(depth) = IIf(wbsArray(depth) <> 0, wbsArray(depth), 0) + 1                                ' clear deeper elements                If wbsArray(depth + 1) <> 0 Then                    For iLoop = depth + 1 To UBound(wbsArray)                        wbsArray(iLoop) = 0                    Next iLoop                End If                                ' build WBS string                strWBS = CStr(iBase)                For iLoop = 0 To depth                    strWBS = strWBS & "." & CStr(wbsArray(iLoop))                Next iLoop            End If                        ' insert WBS string one cell left            With c.Offset(0, -1)                .NumberFormat = "@"                .Value = strWBS                .Errors(xlNumberAsText).Ignore = True            End With                        ' bold previous line if subtasks exist            If 1 = wbsArray(depth) Then _                d.Offset(, -1).Resize(, 2).Font.Bold = True                            Set d = c            If c.Row = n Then Exit For        End If    Next c' Unhide page breaks and disable screen updating (speeds up processing)    Application.ScreenUpdating = True    'ActiveSheet.DisplayPageBreaks = True    End Sub                                                                                                    
