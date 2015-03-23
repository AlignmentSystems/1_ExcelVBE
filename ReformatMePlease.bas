Attribute VB_Name = "ReformatMePlease"
Option Explicit

Sub TestThisCode()
Const cstrMethodName As String = "ReformatMePlease.TestThisCode  "
Const string1 As String = "Hello"
Const string2 As String = "world"
Const string3 As String = "!"

If Len(string2) = Len(string3) Then
    If Len(string1) = Len(string3) Then
        Debug.Print "blah"
    Else
        Debug.Print "blah blah"
    End If
    Debug.Print "blah blah blah"
Else
    Debug.Print "blah blah blah blah"
End If

End Sub
