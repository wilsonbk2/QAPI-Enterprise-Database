Option Explicit

'The first function was a password encryption process to keep employees passwords within the application safe and secure
'The second function was used to log users activity across the entire app, by recording their alias, computer number, domain, what they were doing, etc.
'User access function was used to either allow a user to perform a operation or be denied access. This included access to different areas of the app as well.
'Verify function was interfaced with a satellite database connection in the human resources department to ensure and verify the employee worked at the facility.
''

Global g_Enabled As Boolean

Public Function Encrypt(strvalue As String) As String

Const LowerAlpha    As String = "abcdefghijklmnopqrstuvwxyz"
Const LowerSub      As String = "zebrascdfghijklmnopqtuvwxy" 'zebras
Const UpperAlpha    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const UpperSub      As String = "ZEBRASCDFGHIJKLMNOPQTUVWXY" 'ZEBRAS

Dim lngI            As Long
Dim lngE            As Long
Dim strEncrypt      As String
Dim strLetter       As String

If strvalue & "" = "" Then Exit Function

For lngI = 1 To Len(strvalue)

    strLetter = Mid(strvalue, lngI, 1)

    Select Case Asc(strLetter)

        Case 65 To 90 'Uppercase
            'Find position in alpha string
            For lngE = 1 To Len(UpperAlpha)
                If Mid(UpperAlpha, lngE, 1) = strLetter Then GoTo USub
            Next
USub:
            strEncrypt = strEncrypt & Mid(UpperSub, lngE, 1)

        Case 97 To 122 'Lowercase
            'Find position in alpha string
            For lngE = 1 To Len(LowerAlpha)
                If Mid(LowerAlpha, lngE, 1) = strLetter Then GoTo LSub
            Next
LSub:
            strEncrypt = strEncrypt & Mid(LowerSub, lngE, 1)

        Case Else 'Do not substitute

            strEncrypt = strEncrypt & strLetter

    End Select

Next

'Now pass this string through ROT13 for another tier of security

For lngI = 1 To Len(strEncrypt)
    Encrypt = Encrypt & Chr(Asc(Mid(strEncrypt, lngI, 1)) + 13)
Next

End Function

Public Sub Logging(Activity As String, Accessed As String)
    CurrentDb.Execute "INSERT INTO tblUserLog (UserName, Activity, Accessed, UserComputerName, ComputerName, DomainName) Values('" & TempVars("UserName").value & "', '" & Activity & "', '" & Accessed & "', '" & TempVars("UserComputerName").value & "', '" & TempVars("ComputerName").value & "', '" & TempVars("UserDomain").value & "')"
End Sub

Public Function UserAccess(FormName As String) As Boolean

   UserAccess = Nz(DLookup("HaveAccess", "tbl3EmployeeAccess", "EmployeeType_ID=" & TempVars("EmployeeType") & " AND FormName='" & FormName & "'"), False)

End Function

Public Function VerifyStatus(VERIDESC As String)

     VerifyStatus = Nz(DLookup("VERIFYE", "VERIFY", "VERIDESC='" & VERIDESC & "'"), False)
        
End Function