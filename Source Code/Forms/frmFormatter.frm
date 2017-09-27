VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormatter 
   Caption         =   "Formula Formatter"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   OleObjectBlob   =   "frmFormatter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private off_txtFormula_Width As Double
Private off_txtFormula_Height As Double
Private off_txtStatus_Top As Double
Private off_cmdWrite_Top As Double
Private off_txtStatus_Width As Double
Private off_cmdClose_Left As Double
Private off_cmdClose_Top As Double

Private Sub app_SheetChange(ByVal sh As Object, ByVal Target As Range)
    If TypeName(Selection) = "Range" Then
        If Target.Address(False, False, xlA1, True) = Selection.Address(False, False, xlA1, True) Then
            RefreshFormula
        End If
    End If
End Sub

Private Sub app_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    RefreshFormula
End Sub

Public Sub RefreshFormula()
    Dim strFormula As String

    If TypeName(Selection) = "Range" Then
        With Selection(1)
            If .HasFormula Then
                strFormula = FormatFormulaString(.FormulaLocal)
            Else
                strFormula = .Value
            End If
        End With
        txtFormula.text = strFormula
        txtFormula.SelStart = 0
    End If
End Sub

Private Sub cmdWrite_Click()
    If Left(txtFormula.text, 1) = "=" Then
        If FormulaToSelection(CompactFormulaString(txtFormula.text)) Then
            Status
        Else
            Beep
            Status "Unable to write formula"
        End If
    Else
        Beep
        Status "Not a formula"
    End If
End Sub

Private Sub Status(Optional ByVal str As String = "")
    txtStatus.text = str
End Sub

Private Sub UserForm_Activate()
'    Const WS_THICKFRAME = &H40000, GWL_STYLE = (-16)
'    Dim lngHWnd As Long, lngStyle As Long, lngRet As Long
'
'    lngHWnd = FindWindow("ThunderDFrame", Me.Caption)
'    lngStyle = GetWindowLong(lngHWnd, GWL_STYLE)
'    lngRet = SetWindowLong(lngHWnd, GWL_STYLE, lngStyle Or WS_THICKFRAME)

    txtFormula.SetFocus
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    txtFormula_KeyDown KeyCode, Shift
End Sub

Private Sub cmdClose_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    txtFormula_KeyDown KeyCode, Shift
End Sub

Private Sub cmdWrite_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    txtFormula_KeyDown KeyCode, Shift
End Sub

Private Sub txtStatus_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    txtFormula_KeyDown KeyCode, Shift
End Sub

Private Sub txtFormula_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim str As String

    If KeyCode = vbKeyF9 Then
        With txtFormula
            If .SelLength = 0 Then
                str = .text
            Else
                str = Mid(Replace(.text, vbCr, ""), .SelStart + 1, .SelLength)
            End If
        End With

        Status EvaluateFormula(str)
    End If
End Sub

Private Sub UserForm_Initialize()
    off_txtFormula_Width = Me.Width - txtFormula.Width
    off_txtFormula_Height = Me.Height - txtFormula.Height
    off_txtStatus_Top = Me.Height - txtStatus.Top
    off_txtStatus_Width = Me.Width - txtStatus.Width
    off_cmdWrite_Top = Me.Height - cmdWrite.Top
    off_cmdClose_Left = Me.Width - cmdClose.Left
    off_cmdClose_Top = Me.Height - cmdClose.Top

    Set app = Application
    RefreshFormula
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Resize()
    Dim dbl As Double

    On Error Resume Next
    txtStatus.Top = Me.Height - off_txtStatus_Top
    cmdWrite.Top = Me.Height - off_cmdWrite_Top
    cmdClose.Left = Me.Width - off_cmdClose_Left
    cmdClose.Top = Me.Height - off_cmdClose_Top

    If Me.Width > off_txtFormula_Width And Me.Height > off_txtFormula_Height Then
        txtFormula.Visible = True
        txtFormula.Width = Me.Width - off_txtFormula_Width
        txtFormula.Height = Me.Height - off_txtFormula_Height
    Else
        txtFormula.Visible = False
    End If

    If Me.Width > off_txtStatus_Width Then
        txtStatus.Visible = True
        txtStatus.Width = Me.Width - off_txtStatus_Width
    Else
        txtStatus.Visible = False
    End If

End Sub
