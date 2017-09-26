Attribute VB_Name = "Outlook_calendar"
'add reference to Microsoft Outlook
Public Sub CreateAppointment()
 
  Dim oApp As Outlook.Application
  Dim oNameSpace As Namespace
  Dim oItem As AppointmentItem
      
  On Error Resume Next
  ' check if Outlook is running
  Set oApp = GetObject("Outlook.Application")
  If ERR <> 0 Then
    'if not running, start it
    Set oApp = CreateObject("Outlook.Application")
  End If
  
  Set oNameSpace = oApp.GetNamespace("MAPI")
  
  Set oItem = oApp.CreateItem(olAppointmentItem)
  
  With oItem
  
    .Subject = "This is the subject"
    .start = "05/06/2014 11:45"
    .Duration = "01:00"
    
    .AllDayEvent = False
    .Importance = olImportanceNormal
    .Location = "Room 101"
    
    .ReminderSet = True
    .ReminderMinutesBeforeStart = "10"
    .ReminderPlaySound = True
    .ReminderSoundFile = "C:\Windows\Media\Ding.wav"
    
    Select Case 2 ' do you want to display the entry first or save it immediately?
      Case 1
        .Display
      Case 2
        .Display
        .Save
    End Select
  
  End With
    
  Set oApp = Nothing
  Set oNameSpace = Nothing
  Set oItem = Nothing
     
End Sub




Option Explicit
Public Sub CreateOutlookApptz()
   Sheets("Sheet1").Select
    On Error GoTo Err_Execute
      
    Dim olApp As Outlook.Application
    Dim olAppt As Outlook.AppointmentItem
    Dim blnCreated As Boolean
    Dim olNs As Outlook.Namespace
    Dim CalFolder As Outlook.MAPIFolder
    Dim subFolder As Outlook.MAPIFolder
    Dim arrCal As String
    Dim objRecip As Outlook.Recipients
    Dim myAttendee As Outlook.Recipient
    Dim myOptional As Outlook.Recipient
    Dim conflic As Variant
    Dim i As Long
    Dim M As Integer
    M = 1
    On Error Resume Next
    Set olApp = Outlook.Application
      
    If olApp Is Nothing Then
        Set olApp = Outlook.Application
         blnCreated = True
        ERR.Clear
    Else
        blnCreated = False
    End If
      
    On Error GoTo 0
      
    Set olNs = olApp.GetNamespace("MAPI")
    Set CalFolder = olNs.GetDefaultFolder(olFolderCalendar)
          
    i = 2
    Do Until Trim(Cells(i, 1).Value) = ""
    arrCal = Cells(i, 1).Value
    If arrCal = "Calendar" Then
        Set subFolder = CalFolder
    Else
        Set subFolder = CalFolder.Folders(arrCal)
    End If
searchItem = "[subject] =" & Chr(34) & Sheets("Sheet1").Cells(i, 2).Value & Chr(34)
Set olAppt = subFolder.Items.Find(searchItem)  '& "AND [Categories] =" & Chr(34) & Sheet1.Cells(i, 5).Value & Chr(34)

If TypeName(olAppt) = "Nothing" Then
    Set olAppt = subFolder.Items.Add(olAppointmentItem)
End If
        
    'MsgBox subFolder, vbOKCancel, "Folder Name"
    With olAppt
      
    'Define calendar item properties
        .MeetingStatus = olMeeting
        .start = Cells(i, 6) + Cells(i, 7)     '+ TimeValue("9:00:00")
        .End = Cells(i, 8) + Cells(i, 9)       '+TimeValue("10:00:00")
        .Subject = Cells(i, 2)
        .Location = Cells(i, 3)
        .Body = Cells(i, 4)
        .BusyStatus = olBusy
        .ReminderMinutesBeforeStart = Cells(i, 10) * 24 * 60
        .ReminderSet = True
        .Categories = Cells(i, 5)
        .Display
        eml = Cells(i, 13).Value
        SendKeys eml
        
        'recip = Cells(i, 13).Value
        '.Recipients.Add recip
        .Save
      
    If .Conflicts.Count > 0 Or .IsConflict Then
        ReDim Preserve conflic(1 To M, 1 To 2)
        conflic(M, 1) = Cells(i, 2)
        conflic(M, 2) = .Conflicts.Count
        Cells(i, 12) = "HAS CONFLICTS"
        M = M + 1
    End If
    If Cells(i, 11) = "Delete" Then .Delete
    End With
        'olAppt.Display
        'olAppt.Send
          
        i = i + 1
        Loop
    Set olAppt = Nothing
    Set olApp = Nothing
    Exit Sub
      
Err_Execute:
    MsgBox "An error occurred - Exporting items to Calendar."
      
End Sub

Sub CreateStatusReportToBoss()
Dim OutApp As Outlook.Application
Dim OutMail As Outlook.AppointmentItem

Set OutApp = New Outlook.Application
Set OutMail = OutApp.CreateItem(olAppointmentItem)

With OutMail
   .MeetingStatus = olMeeting
   .Location = " happening"
   .Subject = " Event check "
   '.start = "8:00 PM" & format(Date)
   '.End = "9:00 PM" & format(Date)
   .Body = "this is event details"
   
   .Display
   SendKeys "kevin.joy1@navy.mil" ' This line is not working
   End With
   Sheets("sheet2").Select
   OutMail.Save

End Sub
