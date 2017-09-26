Attribute VB_Name = "M_modFormatter"
Option Explicit

Private Const fmt_BeginNewLine = 2 ^ 0
Private Const fmt_BeginTabDown = 2 ^ 1
Private Const fmt_EndNewLine = 2 ^ 2
Private Const fmt_EndTabUp = 2 ^ 3
Private Const fmt_BeginTabDownOnNonEmptyFunction = 2 ^ 4
Private Const fmt_BeginNewLineOnNonEmptyFunction = 2 ^ 5

Private Const fmt_TabLen = 4

Private Const fmt_FunctionBegin = fmt_BeginNewLine Or fmt_EndNewLine Or fmt_EndTabUp
Private Const fmt_FunctionEnd = fmt_BeginNewLineOnNonEmptyFunction Or fmt_BeginTabDownOnNonEmptyFunction ' Or fmt_EndNewLine
Private Const fmt_FunctionArgument = fmt_EndNewLine
Private Const fmt_ArrayBegin = fmt_BeginNewLine Or fmt_EndNewLine Or fmt_EndTabUp
Private Const fmt_ArrayEnd = fmt_BeginNewLine Or fmt_BeginTabDown
Private Const fmt_ArrayRow = fmt_EndNewLine

Public Sub FormulaFormatter_Click()
    frmFormatter.Show vbModeless
End Sub

Public Function FormulaToSelection(Formula As String) As Boolean
    If TypeName(Selection) = "Range" Then
        On Error Resume Next
        Selection.FormulaLocal = Formula
        If ERR.Number Then
            FormulaToSelection = False
        Else
            FormulaToSelection = True
        End If
        On Error GoTo 0
    End If
End Function

Public Function CompactFormulaString(Formula As String) As String
    Dim udtTokens() As Token, strFormula As String, i As Long, str As String

    strFormula = StripLFCR(Formula)
    udtTokens() = ParseFormula(strFormula)

    strFormula = "="
    For i = 0 To TokenCount(udtTokens())
        With udtTokens(i)
            If (.lngType And tkt_OperandText) = tkt_OperandText Then
                strFormula = strFormula & """" & Replace(.strValue, """", """""") & """"

            ElseIf (.lngType And tkt_OperandReferenceWksQual) = tkt_OperandReferenceWksQual Then
                str = .strValue
                If InStr(1, str, "'") > 0 Or InStr(1, str, " ") > 0 Then str = "'" & Replace(str, "'", "''") & "'"
                strFormula = strFormula & str
            ElseIf (.lngType And tkt_WhiteSpace) = tkt_WhiteSpace Then
                'do nothing
            Else
                strFormula = strFormula & .strValue
            End If
        End With
    Next
    CompactFormulaString = strFormula
End Function

Public Function EvaluateFormula(Formula As String) As String
    Dim var As Variant, str As String, strTemp As String, i As Long, j As Long, k As Long
    
    var = Application.Evaluate(StripLFCR(Formula))
    If ERR.Number Then
        str = "Unable to evaluate formula"
    Else
        If Right(TypeName(var), 2) = "()" Then
            On Error Resume Next
            k = 0: Do: k = k - (LBound(var, k + 1) * 0 = 0): Loop Until ERR.Number
            On Error GoTo 0
            If k = 1 Then
                str = ""
                For i = LBound(var) To UBound(var)
                    str = str & IIf(str = "", "", Application.International(xlColumnSeparator)) & var(i)
                Next
            ElseIf k = 2 Then
                str = ""
                For i = LBound(var, 1) To UBound(var, 1)
                    strTemp = ""
                    For j = LBound(var, 2) To UBound(var, 2)
                        strTemp = strTemp & IIf(strTemp = "", "", Application.International(xlColumnSeparator)) & var(i, j)
                    Next
                    str = str & IIf(str = "", "", Application.International(xlRowSeparator)) & strTemp
                Next
            End If
            str = Application.International(xlLeftBrace) & str & Application.International(xlRightBrace)
        ElseIf TypeName(var) = "Error" Then
            str = "Unable to evaluate formula"
        Else
            str = var
        End If
    End If

    EvaluateFormula = str
End Function

Public Function StripLFCR(str As String) As String
    StripLFCR = Replace(Replace(str, vbLf, ""), vbCr, "")
End Function

Public Function FormatFormulaString(Formula As String) As String
    Dim udtTokens() As Token, strFormula As String, i As Long, strMain As String, strPart As String
    Dim str As String, lngTab As Long, blnNewLine As Boolean, lngTemp As Long, bln As Boolean

    strFormula = Formula

    udtTokens() = ParseFormula(strFormula)

    strMain = "="
    blnNewLine = True
    lngTab = 0
    For i = 0 To TokenCount(udtTokens)
        With udtTokens(i)
            If (.lngType And tkt_Function) = tkt_Function Then
                If strPart <> "" Then
                    WritePart strMain, strPart, lngTab, blnNewLine, 0
                    strPart = ""
                End If

                If (.lngType And tkt_Begin) = tkt_Begin Then
                    WritePart strMain, .strValue, lngTab, blnNewLine, fmt_FunctionBegin

                ElseIf (.lngType And tkt_End) = tkt_End Then
                    lngTemp = fmt_FunctionEnd
                    bln = (udtTokens(i - 1).lngType And (tkt_Function Or tkt_Begin)) = (tkt_Function Or tkt_Begin)

                    If (lngTemp And fmt_BeginNewLineOnNonEmptyFunction) = fmt_BeginNewLineOnNonEmptyFunction Then
                        If bln Then blnNewLine = False Else lngTemp = lngTemp Or fmt_BeginNewLine
                    Else
                        lngTemp = lngTemp Or fmt_BeginNewLine
                    End If
                    If (lngTemp And fmt_BeginTabDownOnNonEmptyFunction) = fmt_BeginTabDownOnNonEmptyFunction Then
                        If bln Then lngTab = lngTab - 1 Else lngTemp = lngTemp Or fmt_BeginTabDown
                    Else
                        lngTemp = lngTemp Or fmt_BeginTabDown
                    End If
                    WritePart strMain, .strValue, lngTab, blnNewLine, lngTemp

                End If

            ElseIf (.lngType And tkt_FunctionArgument) = tkt_FunctionArgument Then
                If strPart <> "" Then
                    WritePart strMain, strPart, lngTab, blnNewLine, 0
                    strPart = ""
                End If

                WritePart strMain, .strValue, lngTab, blnNewLine, fmt_FunctionArgument

            ElseIf (.lngType And tkt_Array) = tkt_Array Then
                If strPart <> "" Then
                    WritePart strMain, strPart, lngTab, blnNewLine, 0
                    strPart = ""
                End If

                If (.lngType And tkt_Begin) = tkt_Begin Then
                    WritePart strMain, .strValue, lngTab, blnNewLine, fmt_ArrayBegin

                ElseIf (.lngType And tkt_End) = tkt_End Then
                    WritePart strMain, .strValue, lngTab, blnNewLine, fmt_ArrayEnd

                End If

            ElseIf (.lngType And tkt_ArrayRow) = tkt_ArrayRow Then
                If strPart <> "" Then
                    WritePart strMain, strPart, lngTab, blnNewLine, 0
                    strPart = ""
                End If

                WritePart strMain, .strValue, lngTab, blnNewLine, fmt_ArrayRow

            ElseIf (.lngType And tkt_OperandText) = tkt_OperandText Then
                strPart = strPart & """" & Replace(.strValue, """", """""") & """"

            ElseIf (.lngType And tkt_OperandReferenceWksQual) = tkt_OperandReferenceWksQual Then
                str = .strValue
                If InStr(1, str, "'") > 0 Or InStr(1, str, " ") > 0 Then str = "'" & Replace(str, "'", "''") & "'"
                strPart = strPart & str

            Else
                strPart = strPart & .strValue

            End If
        End With
    Next
    If strPart <> "" Then WritePart strMain, strPart, lngTab, blnNewLine, 0

    FormatFormulaString = strMain
End Function

Sub WritePart(strMain As String, strPart As String, lngTab As Long, blnNewLine As Boolean, lngFormat As Long)
    If (lngFormat And fmt_BeginTabDown) = fmt_BeginTabDown Then lngTab = lngTab - 1
    If (lngFormat And fmt_BeginNewLine) = fmt_BeginNewLine Then blnNewLine = True
    If Not strMain = "" And blnNewLine Then
        strMain = strMain & vbNewLine & String(lngTab * fmt_TabLen, " ")
        blnNewLine = False
    End If
    strMain = strMain & strPart
    If (lngFormat And fmt_EndNewLine) = fmt_EndNewLine Then blnNewLine = True
    If (lngFormat And fmt_EndTabUp) = fmt_EndTabUp Then lngTab = lngTab + 1
End Sub
