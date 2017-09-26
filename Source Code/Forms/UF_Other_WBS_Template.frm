VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Other_WBS_Template 
   Caption         =   "Additional WBS Templates:"
   ClientHeight    =   1050
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6852
   OleObjectBlob   =   "UF_Other_WBS_Template.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Other_WBS_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB_OtherWBS_Click()

 Dim location As Range
    Dim WBS_Appendix_Select
    Dim AppendixList(1) As Variant
       
    WBS_Appendix_Select = UF_Other_WBS_Template.CB_OtherWBS.ListIndex
    Debug.Print UF_Other_WBS_Template.CB_OtherWBS.ListIndex
    
    AppendixList(0) = "Table_NSA"
'    AppendixList(1) = ""
'    AppendixList(2) = ""
'    AppendixList(3) = ""
'    AppendixList(4) = ""
'    AppendixList(5) = ""
'    AppendixList(6) = ""
'    AppendixList(7) = ""
'    AppendixList(8) = ""
'    AppendixList(9) = ""
'    AppendixList(10) = ""
'    AppendixList(11) = ""
       
    Unload UF_Other_WBS_Template

    Set location = Application.InputBox("Choose Destination Cell For First WBS Element", _
        Title:="WBS Input Location", _
        Type:=8)
    
    ThisWorkbook.Sheets("Other WBS Templates").Range(AppendixList(WBS_Appendix_Select)).Copy Destination:=location
    
    
    Call wbsGroupInd(Range(location, location.End(xlDown)), 0)


End Sub


Private Sub UserForm_Initialize()

Dim WBS

For Each WBS In ThisWorkbook.Sheets("Other WBS Templates").Range("Table_OtherWBSTemplate")
    UF_Other_WBS_Template.CB_OtherWBS.AddItem WBS.Value
Next WBS

UF_Other_WBS_Template.CB_OtherWBS.Value = "Please select appendix"

End Sub

