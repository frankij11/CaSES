VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MILSTD881C 
   Caption         =   "MIL-STD-881C WBS Selector Tool:"
   ClientHeight    =   1050
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6852
   OleObjectBlob   =   "UF_MILSTD881C.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_MILSTD881C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB_MILSTD881C_Click()
    Dim location As Range
    Dim WBS_Appendix_Select
    Dim AppendixList(11) As Variant
       
    WBS_Appendix_Select = UF_MILSTD881C.CB_MILSTD881C.ListIndex
    Debug.Print UF_MILSTD881C.CB_MILSTD881C.ListIndex
    
    AppendixList(0) = "Table_Aircraft"
    AppendixList(1) = "Table_Electronic"
    AppendixList(2) = "Table_Missile"
    AppendixList(3) = "Table_Ordnance"
    AppendixList(4) = "Table_Sea"
    AppendixList(5) = "Table_Space"
    AppendixList(6) = "Table_Surface_Vehicle"
    AppendixList(7) = "Table_Surface_Vehicle"
    AppendixList(8) = "Table_UAV_System"
    AppendixList(9) = "Table_Unmanned_Maritime_System"
    AppendixList(10) = "Table_Launch_Vehicle"
    AppendixList(11) = "Table_AIS"
       
    Unload UF_MILSTD881C

    Set location = Application.InputBox("Choose Destination Cell For First WBS Element", _
        Title:="WBS Input Location", _
        Type:=8)
    
    ThisWorkbook.Sheets("MILSTD881C_Datatables").Range(AppendixList(WBS_Appendix_Select)).Copy Destination:=location
    
    
    Call wbsGroupInd(Range(location, location.End(xlDown)), 0)
    
    
  
End Sub

Private Sub UserForm_Initialize()

Dim WBS

For Each WBS In ThisWorkbook.Sheets("MILSTD881C_Datatables").Range("Table_List")
    UF_MILSTD881C.CB_MILSTD881C.AddItem WBS.Value
Next WBS

UF_MILSTD881C.CB_MILSTD881C.Value = "Please select appendix"

End Sub


