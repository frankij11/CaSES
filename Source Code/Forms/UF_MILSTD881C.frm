VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MILSTD881C 
   Caption         =   "MIL-STD-881C WBS Selector Tool:"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
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
       
    Dim wbs_Appendix
    wbs_Appendix = UF_MILSTD881C.CB_MILSTD881C.Value
       
       
    Unload UF_MILSTD881C

    Set location = Application.InputBox("Choose Destination Cell For First WBS Element", _
        Title:="WBS Input Location", _
        Type:=8)
    
    'if multiple cells are selected choose first cell
    If location.count > 1 Then Set location = location(1, 1)
        
    
    
    Dim wbsRange As Range
    'Filter Table
    ThisWorkbook.Sheets("WBS").Range("Table_WBS").AutoFilter 1, wbs_Appendix
    
    
    'set range equal only to values in the current filter
    Set wbsRange = ThisWorkbook.Sheets("WBS").Range("Table_WBS[[WBS Code]:[Element Title:]]").SpecialCells(xlCellTypeVisible)
    
    'Check if location will require over writing
    Dim rngCheck As Range
    Set rngCheck = Range(location, location.Offset(wbsRange.Rows.count - 1, wbsRange.Columns.count - 1))
    Dim cont ' boolean to determine if you should continue or exit routine
    Debug.Print rngCheck.Address
    cont = True 'default set to continue
    If WorksheetFunction.CountA(rngCheck) > 0 Then 'count number of non blank cells

        cont = MsgBox("Do you want to overwrite cells (" & rngCheck.Address & ") ?", vbYesNo)
        If cont = vbNo Then cont = False
    End If
    
    'Copy wbs to location only if user specifies cont
    If cont Then
        wbsRange.Copy 'Destination:=location
        location.PasteSpecial (xlPasteValues)
        Call wbsGroupInd(Range(location, location.End(xlDown)), 0)
    
    End If
        
    
    'ThisWorkbook.Sheets("WBS").Range("Table_WBS").AutoFilter = False
End Sub

Private Sub UserForm_Initialize()

Dim WBS
Dim tmp
For Each WBS In ThisWorkbook.Sheets("WBS").Range("Table_WBS[APPENDIX]")
    If (WBS.Value <> "") And (InStr(tmp, WBS.Value) = 0) Then
      tmp = tmp & WBS.Value
      UF_MILSTD881C.CB_MILSTD881C.AddItem WBS.Value
    End If
Next WBS

UF_MILSTD881C.CB_MILSTD881C.Value = "Please select appendix"

End Sub


