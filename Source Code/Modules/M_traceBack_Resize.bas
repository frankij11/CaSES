Attribute VB_Name = "M_traceBack_Resize"
'Written: February 14, 2011
'Author:  Leith Ross
'
'NOTE:  This code should be executed within the UserForm_Activate() event.

'Private Declare Function GetForegroundWindow Lib "User32.dll" () As Long
'
'Private Declare Function GetWindowLong _
'  Lib "User32.dll" Alias "GetWindowLongA" _
'    (ByVal hwnd As Long, _
'     ByVal nIndex As Long) _
'  As Long
'
'Private Declare Function SetWindowLong _
'  Lib "User32.dll" Alias "SetWindowLongA" _
'    (ByVal hwnd As Long, _
'     ByVal nIndex As Long, _
'     ByVal dwNewLong As Long) _
'  As Long
'
'Private Const WS_THICKFRAME As Long = &H40000
'Private Const GWL_STYLE As Long = -16
'
'Public Sub MakeFormResizable()
'
'  Dim lStyle As Long
'  Dim hwnd As Long
'  Dim RetVal
'
'    hwnd = GetForegroundWindow
'
'    'Get the basic window style
'     lStyle = GetWindowLong(hwnd, GWL_STYLE) Or WS_THICKFRAME
'
'    'Set the basic window styles
'     RetVal = SetWindowLong(hwnd, GWL_STYLE, lStyle)
'
'End Sub

