VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub EID_AfterUpdate()
If Not IsNumeric(EID.value) Then
        MsgBox "only numbers allowed"
        Me.EID.BorderColor = vbRed
        Me.EID = Null
    End If
If Len(EID.Text) < 9 Then
   MsgBox "Too short!"
   Me.EID.BorderColor = vbRed
   Me.EID = Null
End If
If Len(EID.Text) > 9 Then
   MsgBox "Too Long!"
   Me.EID.BorderColor = vbRed
   Me.EID = Null
End If
End Sub

Private Sub Form_Close()
    '// load the form in full  transparent mode
    Call SetTranslucent(Me.hwnd, 0)
    '// ensure the form is visible for proper fade effect
    '// (thought still full transparent)
    Me.Visible = True
    '// process fade in effect and provide the opacity value
    '// not to exceed 255
    Call UIProcessFadeOut(Me, 255)
End Sub

Private Sub Form_Current()
    Me.txtUserName.SetFocus
    Me.txtUserName.SelStart = 0
    
    Dim S As String
    S = Globals.Encrypt(Me.txtPassword)
    Me.Encrypt2.value = S
End Sub

Private Sub Form_Load()
DoCmd.MoveSize 9500, 6000, 9000, 5200
 '// load the form in full  transparent mode
    Call SetTranslucent(Me.hwnd, 0)
    '// ensure the form is visible for proper fade effect
    '// (thought still full transparent)
    Me.Visible = True
    '// process fade in effect and provide the opacity value
    '// not to exceed 255
    Call UIProcessFadeIn(Me, 255)
    Me.txtUserName.SetFocus
    Me.txtPassword.SetFocus
'    Me.txtuser.SetFocus
'    Me.txtuser2.SetFocus
'    Me.txtuser3.SetFocus
End Sub

Private Sub btnLogin_Click()
If MsgBox("Are you sure you want to continue? Please make sure your department is correct.", vbYesNo + vbQuestion) = vbYes Then
    Dim strSql As String
    
    If IsNull(Me.firstname) _
     Or IsNull(Me.lastname) _
     Or IsNull(Me.DEPT) _
     Or IsNull(Me.txtUserName) _
     Or IsNull(Me.EID) _
     Or IsNull(Me.txtPassword) Then
     
    MsgBox "You must fill out required feilds.", vbCritical, "Data entry error... All Fields Required"
    Me.firstname.BorderColor = vbRed
    Me.lastname.BorderColor = vbRed
    Me.DEPT.BorderColor = vbRed
    Me.EID.BorderColor = vbRed
    GoTo TheEnd
    End If
    
    
    CurrentDb.Execute "Insert Into tbl1Employees (FirstName, LastName, Password, EmployeeType_ID, Alias, EMPID) " & "Values ('" & Me.firstname & "', '" & Me.lastname & "', '" & Me.Encrypt2 & "', '" & Me.DEPT & "', '" & Me.txtPassword & "', '" & Me.EID & "')"
    CurrentDb.Execute "Insert Into tbl4UserAlias (Alias, UserName) " & "Values ('" & Me.txtPassword & "', '" & Me.txtUserName & "')"

'        Me.MedA.Locked = True
'        Me.MedVAL.Locked = True
        
        MsgBox "Thank you for creating an account. You can now login."
        
        DoCmd.Close acForm, "frmCreateAccount"
        DoCmd.OpenForm "frmLogin"
        
End If


TheEnd:
    
End Sub

Private Sub txtPassword_Enter()
    Me.txtPassword.value = Environ$("UserName")
End Sub

Private Sub txtuser_Enter()
    Me.txtuser.value = Environ$("UserName")
End Sub

Private Sub txtuser2_Enter()
    Me.txtuser2.value = Environ$("UserDomain")
End Sub

Private Sub txtuser3_Enter()
    Me.txtuser3.value = Environ$("computername")
End Sub

Private Sub txtUserName_Enter()
    Me.txtUserName.value = Environ$("computername")
End Sub

