VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabstoSheets_WBS_Option 
   Caption         =   "Select WBS Format"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   OleObjectBlob   =   "TabstoSheets_WBS_Option.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabstoSheets_WBS_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CB_OptionA_Click()
    'This module will allow users to select a group of WBS elements and then create tabs for each one

Unload TabstoSheets_WBS_Option
createWBSTabs (True)

End Sub

Public Sub CB_OptionB_Click()

Unload TabstoSheets_WBS_Option
createWBSTabs (False)
End Sub


