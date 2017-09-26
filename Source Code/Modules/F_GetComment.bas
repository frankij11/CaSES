Attribute VB_Name = "F_GetComment"
Option Explicit

' Referenced from http://chandoo.org/wp/2009/09/03/get-cell-comments/
Function F_GetComment(incell) As String
 ' aceepts a cell as input and returns its comments (if any) back as a string
 On Error Resume Next
 F_GetComment = incell.Comment.text
End Function
