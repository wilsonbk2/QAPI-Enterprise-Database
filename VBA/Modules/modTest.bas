Attribute VB_Name = "modTest"
Option Compare Database
Option Explicit

Global strTest As String
'Public strTest As String
Private Operator As String

Public Function PublicTest() As String
    PublicTest = "Public Access Modifier"
End Function

Private Function PrivateTest() As String
    PrivateTest = "Private Access Modifier"
End Function

Public Function CallToPrivateTest() As String
    CallToPrivateTest = PrivateTest
End Function

Private Sub Whatever()
    strTest = "blah"
End Sub
