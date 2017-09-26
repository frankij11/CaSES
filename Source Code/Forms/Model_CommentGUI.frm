VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Model_CommentGUI 
   Caption         =   "NCCA Model Comment and Documentation Data Entry Tool"
   ClientHeight    =   3570
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   9000.001
   OleObjectBlob   =   "Model_CommentGUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Model_CommentGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function wsExists(wksName As String) As Boolean
    On Error Resume Next
    wsExists = Len(Worksheets(wksName).Name)
    On Error GoTo 0
End Function
Function rangeExists(rangeName As String) As Boolean
    On Error Resume Next
    rangeExists = Len(Range(rangeName).Name)
    On Error GoTo 0
End Function
Set modelcomments = addSheet("Comment Tracker")


Function addSheet(SheetName As String, Optional wbName As String) As Worksheet
    If IsMissing(wbName) Then
        Set addSheet = ActiveWorkbook.Worksheets.Add
    Else
        Set addSheet = Workbooks(wbName).Worksheets.Add
    End If
    
    On Error Resume Next
    addSheet.Name = SheetName
    Do Until i = 10
        i = i + 1
        
        If Len(addSheet.Name) - Len(Replace(addSheet.Name, SheetName, "")) > 0 Then Exit Function
        addSheet.Name = SheetName & "(" & i & ")"
    Loop
End Function
Private Sub CB_Maximize_Click()

On Error Resume Next
    Application.WindowState = xlMaximized
    Model_CommentGUI.Show vbModeless

End Sub

Private Sub CB_Minimize_Click()

On Error Resume Next
    Application.WindowState = xlMinimized
    Model_CommentGUI.Show vbModeless

End Sub










Private Sub ComboBox_DocumentName_Change()
On Error GoTo NonXL:
    Dim testBool As Boolean
    Dim testWB As Workbook
    testBool = False
    Set testWB = Workbooks(ComboBox_DocumentName.Value)
    testWB.Activate
    TextBox_Reference.Value = Replace(Selection.Address(False, False, , True), "[" & Selection.Parent.Parent.Name & "]", "")
    testBool = True
    Exit Sub
On Error GoTo 0

NonXL:
'On Error GoTo NonWRD:
'    Dim wd As Object
'    Set wd = GetObject(, "Word.Application")
'    wd.documents(ComboBox_DocumentName.Value).Activate
'    wd.Visible = True
'    Application.WindowState = xlMinimized
'
'    Model_CommentGUI.Show vbModeless
'On Error GoTo 0

NonWRD:

NonPPT:
    TextBox_Reference.Value = ""
End Sub

Private Sub CommandButton_AddComment_Click()
Call insertData
End Sub
Private Sub insertData()

' Future Adds:
'This portion of code eliminates the active screen updating so users do not see the spreadsheet switching to and from the ModelComments tab in the ActiveWorkbook
Dim OriginalActiveWorksheet As Variant

Application.ScreenUpdating = False
Set OriginalActiveWorksheet = Application.ActiveSheet
Dim modelcomments As Object
On Error Resume Next
Workbooks(Me.ComboBox_Workbook.Value).Activate

'Set modelcomments = Workbooks(Me.ComboBox_Workbook.Value).Worksheets("Comment Tracker").Range("Comment_Tracker")
Set modelcomments = Workbooks(Me.ComboBox_Workbook.Value).Worksheets(Range("Comment_Tracker").Parent.Name).Range("Comment_Tracker")
On Error GoTo 0
'This section of code determines whether a tab named ModelComments is included within the Activeworkbook
    If modelcomments Is Nothing Then
        Set modelcomments = addSheet("Comment Tracker", Me.ComboBox_Workbook.Value)
        ThisWorkbook.Worksheets("Comment Tracker").Cells.Copy
        modelcomments.Range("A1").PasteSpecial (xlPasteAll)
        Workbooks(Me.ComboBox_Workbook.Value).Activate
        modelcomments.Select
        Range("A1").Select
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 70

        'This ends the first portion of the primary If Statement that determines if the Activeworkbook contains a sheet named ModelComments
        Set modelcomments = Nothing
    End If

    'Debug.Print Range("Comment_Tracker").Parent.Name
    Set modelcomments = Workbooks(Me.ComboBox_Workbook.Value).Worksheets(Range("Comment_Tracker").Parent.Name)
    'MsgBox ("Worksheet does Exist")

    'This section finds the last used row within the Activeworkbook Model Comments Tab
  
    Dim TheLastRow As Integer
    'Dim DataEntryRow As Long
    Dim nextRow As Integer
    'nextRow = Range("comment_tracker").SpecialCells(xlCellTypeLastCell).Offset(1).Row
    Dim nextCell As Range
    'modelcomments.Activate
    Set nextCell = modelcomments.Range("Comment_Tracker[[#Headers],[ID]]")
    Do Until nextCell = ""
        Set nextCell = nextCell.Offset(1)
    Loop
    nextRow = nextCell.row
      
    modelcomments.Cells(nextRow, Range("comment_tracker[id]").Column) = WorksheetFunction.Max(modelcomments.Range("comment_tracker[id]")) + 1
    modelcomments.Cells(nextRow, Range("comment_tracker[Date Open]").Column) = DTPicker_CommentGenerated.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Date Due]").Column) = DTPicker_CommentResponse.Value
  
    'Determine if you want to add reference

    modelcomments.Cells(nextRow, Range("comment_tracker[Document Type]").Column) = ComboBox_DocumentType.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Document Name]").Column) = ComboBox_DocumentName.Value 'ActiveSheet.Name
    modelcomments.Cells(nextRow, Range("comment_tracker[Reference]").Column) = TextBox_Reference 'ActiveSheet.Name
    'modelcomments.Cells(nextRow, Range("comment_tracker[cell]").Column) = Selection.Address(False, False)

  
  
    modelcomments.Cells(nextRow, Range("comment_tracker[Subject]").Column) = TextBox_Subject.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Commenter POC]").Column) = TextBox_CommentOwner.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Assigned To]").Column) = ComboBox_CommentDeliverTo.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Comment Type]").Column) = ComboBox_CommentType.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Critical]").Column) = ComboBox_Critical.Value
    modelcomments.Cells(nextRow, Range("comment_tracker[Comment]").Column) = TextBox_CommentData.Value
      

'This portion of code reactivates screenupdating and brings the userback to the last ActiveSheet before using the Comment Data Entry Tool
ActiveWindow.DisplayGridlines = False
Application.ScreenUpdating = True
OriginalActiveWorksheet.Activate
Set modelcomments = Nothing


TextBox_CommentData.text = ""
TextBox_Subject = ""


End Sub


Private Sub CommandButton_Close_Click()
Unload Me
End Sub

Private Sub Label12_Click()
On Error Resume Next
ComboBox_DocumentName.Value = Selection.Parent.Parent.Name
TextBox_Reference.Value = Replace(Selection.Address(False, False, , True), "[" & Selection.Parent.Parent.Name & "]", "")
End Sub



Private Sub UserForm_Activate()

With Me.ComboBox_Workbook
    Dim wkbk As Workbook
    For Each wkbk In Application.Workbooks
        If VBA.UCase(wkbk.Name) <> "PERSONAL.XLSB" Then .AddItem wkbk.Name
    Next
    .Value = ActiveWorkbook.Name
End With

TextBox_CommentOwner.Value = Application.UserName

With Me.ComboBox_CommentDeliverTo
    .AddItem "SYSCOM"
    .AddItem "NCCA"
    .AddItem "CAPE"
    .Value = "SYSCOM"
End With

With Me.ComboBox_CommentType
    .AddItem "Comprehensive"
    .AddItem "Documentation"
    .AddItem "Accuracy"
    .AddItem "Credible"
    .Value = "Documentation"
End With

With Me.ComboBox_Critical
    .AddItem "Minor"
    .AddItem "Major"
    .AddItem "Admin Only"
    .Value = "Major"
End With


With Me.ComboBox_DocumentType
    .AddItem "CARD"
    .AddItem "Model"
    .AddItem "No Update Required"
    .Value = "CARD"
End With


With Me.ComboBox_DocumentName
    Dim i As Integer
    
    For Each wkbk In Application.Workbooks
        If VBA.LCase(wkbk.Name) <> "personal.xlsb" Then
            .AddItem wkbk.Name
        End If
    Next
    
    On Error Resume Next
    Dim wrd As Object
    Dim doc As Object
    Set wrd = GetObject(, "Word.Application")
    For i = 1 To wrd.documents.count
        .AddItem wrd.documents(i).Name
    Next
    
    
    Dim ppt As Object
    Dim pres As Object
    
    Set ppt = GetObject(, "PowerPoint.Application")
    For i = 1 To ppt.presentations.count
        .AddItem ppt.presentations(i).Name
    Next
    On Error GoTo 0
    .Value = ActiveWorkbook.Name
End With
'Me.ComboBox_UpdateRequired.Selected(1) = True
'Me.ComboBox_CommentType.Selected(1) = True

With Me.DTPicker_CommentGenerated
    .Value = Now
End With

With Me.DTPicker_CommentResponse
    .Value = Now + 7
End With

End Sub

