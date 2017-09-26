Attribute VB_Name = "a_AboutAddin"


'MIT License
'
'Copyright (c) 2017 CaSES
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
Option Explicit

Sub About_CT()
Dim CT_Date
Dim CT_version

CT_Date = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")
CT_version = ThisWorkbook.Name
CT_version = Replace(CT_version, "_", " Version: ")
CT_version = Replace(CT_version, ".xlam", "")
CT_version = Replace(CT_version, ".xlsm", "")

MsgBox Prompt:="Originally Created by:" _
                & vbNewLine & "Nicholas Lanham and Kevin Joy" _
                & vbNewLine & vbNewLine & "Major Contributors:" _
                & vbNewLine & "Duncan Thomas, Naval Center for Cost Analysis" _
                & vbNewLine & vbNewLine & "This tool was developed using an open source concept and is available to all users at no cost." _
                & vbNewLine & vbNewLine & "This version was last updated on: " & vbNewLine & CT_Date & vbNewLine & CT_version _
                & vbNewLine & vbNewLine & "Add-In location: " & ThisWorkbook.Path, _
        Title:="Cost and Schedule Estimating Suite (CaSES)"

End Sub
