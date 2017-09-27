VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress Bar"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9210
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ShowBar(Optional Title As String)
    
    'Do Events makes sure the rest of your macro keeps running
    DoEvents
    If IsMissing(Title) Then Title = "Progress Bar"
    
    'Set the Width of the Progressbar to Zero
    Me.Bar.Width = 0
    'Update the Title of the Form
    Me.Caption = Title
    'Initialize the Private Class Variable
    'cFormShowStatus = True
    
    'Show the Form
    Me.Show
    'Repaint the Form
    Me.Repaint

End Sub

Public Sub progress(pctComplete)
DoEvents
Me.Bar.Width = pctComplete * (ProgressBar.Frame1.Width)
Me.Bar.Caption = pctComplete * 100 & "%"
End Sub

Public Sub complete(Optional howLong)
Dim wait
If IsMissing(howLong) Then howLong = 3
wait = Timer + howLong
Do While Timer < wait
DoEvents
Loop
Terminate

End Sub

Public Sub Terminate(Optional howLong)

Me.Hide

End Sub


