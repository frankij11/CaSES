VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FlatFile_USF 
   Caption         =   "Flat File Creator "
   ClientHeight    =   3647
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   5040
   OleObjectBlob   =   "FlatFile_USF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FlatFile_USF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Flat_File_Run_Click()
'Application created by Nicholas Lanham and Kevin Joy
    
    FlatFile_USF.Hide
    
    'Declare Current as a worksheet object variable.
    Dim WrkSheet As Worksheet
    Dim CopyArea_Header As Range
    Dim CopyArea_DataRow As Range
    Dim CopyArea_DataRowRange As String
    Dim WS_Count As Integer
    Dim Copy_Count As Integer
    Dim I As Integer
    Dim Transpose_Data_CopyRange As Range
    Dim First_Loop_WrkSheet As Worksheet
    Dim Last_Loop_WrkSheet As Worksheet
        
    'Set MS Excel error messages to false
    Application.DisplayAlerts = False
    
    Dim LB_I As Long
    Dim Array_FlatFile() As Variant
    Dim Array_Paste() As Variant
    Dim Size As Integer
    Dim AI As Integer
    Dim C As Integer
    Dim NUMSELECT As Integer
    Dim MatchTab As Variant
    Dim X As Variant
    Dim Y As Variant
    Dim Z As Integer
    Dim Cel As Range
    Dim CopyEachValue As Range
    Dim Paste_Destination As Range
       
    
    'Code to count how many items have been selected to size the array
    For C = 0 To FlatFile_USF.FlatFile_FF_LB1.ListCount - 1
        If FlatFile_USF.FlatFile_FF_LB1.Selected(C) = True Then
            NUMSELECT = NUMSELECT + 1
        End If
    Next C
    Debug.Print NUMSELECT
    
    'Code to Size Array
    ReDim Array_FlatFile(NUMSELECT)
    
    'Code to add selected data tabs to array
    For LB_I = 0 To FlatFile_USF.FlatFile_FF_LB1.ListCount - 1
        If FlatFile_USF.FlatFile_FF_LB1.Selected(LB_I) = True Then
         'Debug.Print CSDRFF_USF.CSDR_FF_LB1.List(LB_I)
         
        Array_FlatFile(AI) = FlatFile_USF.FlatFile_FF_LB1.List(LB_I)
        Debug.Print Array_FlatFile(AI)
        AI = AI + 1
            
        End If
    Next

    
    'Add the DataTable worksheet to the begining of workbook
    On Error Resume Next
    If ActiveWorkbook.Sheets("DataTable") Is Nothing Then
    On Error GoTo 0
        ActiveWorkbook.Sheets.Add(Before:=Worksheets(1)).Name = "DataTable"
        'Define last flatfile format table output sheet name
        Set Last_Loop_WrkSheet = ActiveWorkbook.Sheets("DataTable")
     Else
        Confirm_DataTable_MsgBox = MsgBox("Datatable tab already exists. Do you want to overwrite with a new Datatable?", vbYesNo, "Check if DataTable Exists")
        If Confirm_DataTable_MsgBox = vbYes Then
            ActiveWorkbook.Sheets("DataTable").Delete
            ActiveWorkbook.Sheets.Add(Before:=Worksheets(1)).Name = "DataTable"
        'Define last flatfile format table output sheet name
        Set Last_Loop_WrkSheet = ActiveWorkbook.Sheets("DataTable")
        Else
            Exit Sub
        End If
     End If
    
    'Count the number of worksheets in workbook
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    'Define last flatfile format table output sheet name
    Set Last_Loop_WrkSheet = ActiveWorkbook.Sheets("DataTable")
    
     'Select user defined header information and paste header information into DataTable sheet
    On Error Resume Next
    Set CopyArea_Header = Application.InputBox("Enter the range that includes the header labels for the new datatable." & vbCrLf & "Reminder: These cells will be the labels for the new data table." & vbCrLf & "Tool Created By: Nick Lanham and Kevin Joy", Type:=8)
         
        If CopyArea_Header Is Nothing Or CopyArea_Header.Columns.Count > 1 Then
            MsgBox ("Please try again by selecting a single column as the range that includes header information")
            Exit Sub
        Else
        
        With CopyArea_Header
            .Copy
        End With
 
        
            
        'Past all headers into the datatable sheet
        Last_Loop_WrkSheet.Range("D2").Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        Application.CutCopyMode = False
        ActiveCell.Offset(0, -1).Value = "Tab Name"
        With Activcell
            .Font.Bold = True
        End With
                        
        Last_Loop_WrkSheet.Range("C2", Last_Loop_WrkSheet.Range("C2").End(xlToRight)).Select
            With Selection
                .WrapText = True
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
           
        End If
            
        'Allows user to select raw data input range and use this range forthe following copy/paste loop
        On Error Resume Next
        Set CopyArea_DataRow = Application.InputBox("Enter the range that includes the data for the new flat file" & vbCrLf & "Reminder: This will be the data copied from each tab in the activeworkbook." & vbCrLf & "Tool Created By: Nick Lanham and Kevin Joy", Type:=8)
        
        MsgBox (Range("CopyArea_DataRow").Columns.Count)
        
        If CopyArea_DataRow Is Nothing Or CopyArea_DataRow.Columns.Count > 1 Then
                MsgBox ("Please try again by selecting a single column as the range that includes data for flat file output.")
                Exit Sub
        Else
            CopyArea_DataRowRange = CopyArea_DataRow.Address
        End If
        Debug.Print CopyArea_DataRowRange
        'Loop through all of the worksheets in the active workbook.
        'Clear ListBox integer counter
        
        'LB_I = 1
        
        'For X = LBound(Array_FlatFile) To UBound(Array_FlatFile)
        '   Debug.Print Array_FlatFile(X)
        'Next X
        
        
        X = Nothing
        
        Last_Loop_WrkSheet.Activate
        Last_Loop_WrkSheet.Range("D3").Select
        
        For I = 2 To WS_Count
        
            Set CopyEachValue = ActiveWorkbook.Sheets(I).Range(CopyArea_DataRowRange)
            'Debug.Print CopyEachValue.Address

            ReDim Array_Paste(1 To CopyArea_DataRow.Rows.Count)
            'Debug.Print "Array Paste =" & CopyArea_DataRow.Rows.Count

            For X = LBound(Array_FlatFile) To UBound(Array_FlatFile)
                If ActiveWorkbook.Sheets(I).Name = Array_FlatFile(X) Then
                    For Each Cel In CopyEachValue
                        Y = 1
                        Array_Paste(Y) = Cel.Value
                        Debug.Print "These are the cell values =" & Cel.Value
                        ActiveCell.Value = Array_Paste(Y)
                        ActiveCell.Offset(0, 1).Select
                        Y = Y + 1
                    Next Cel
                    
                ActiveCell.Offset(0, -1 * (1 + CopyArea_DataRow.Rows.Count)).Select
                
                With Last_Loop_WrkSheet
                    .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
                    "'" & ActiveWorkbook.Sheets(I).Name & "'" & "!A1", TextToDisplay:=ActiveWorkbook.Sheets(I).Name
                End With
                
                ActiveCell.Offset(1, 1).Select
                  
                End If
                
            'ActiveCell.Offset(0, -1 * CopyArea_DataRow.Rows.Count).Select
            'ActiveCell = ActiveWorkbook.Sheets(I).Name
                            
            'Adds hyperlink to tab names

            
            Next X
            
        Next I
    
    'Format newly created flat file data table
    Last_Loop_WrkSheet.Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    With Selection.Borders
        .xlContinuous
        .ColorIndex = vbBlack
        .Weight = xlAutomatic
    End With
    
    'clear excel memory to avoid out of memory error
    Set Last_Loop_WrkSheet = Nothing
    Set CopyArea_Header = Nothing
    Set CopyArea_DataRow = Nothing
    Set Transpose_Data_CopyRange = Nothing
    'Set LB_I = Nothing
    'Set Array_FlatFile() = Nothing
    'Set Size = Nothing
    'Set AI = Nothing
    'Set C = Nothing
    'Set NUMSELECT = Nothing
         
         

End Sub

Private Sub UserForm_Initialize()

Dim WS As Worksheet

For Each WS In ActiveWorkbook.Worksheets

FlatFile_FF_LB1.AddItem (WS.Name)

Next WS

End Sub
