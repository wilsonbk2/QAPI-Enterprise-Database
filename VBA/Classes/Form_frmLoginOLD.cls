VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLoginOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'// transparency value
Private Const UIOpacity As Integer = 175

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
    DoCmd.OpenForm "frmDemo", acNormal, , , , acHidden
    DoCmd.OpenForm "frmLogoutTimer", acNormal, , , , acHidden
    Me.txtUserName.SetFocus
    Me.txtUserName.SelStart = 0
End Sub

Private Sub Form_Load()
DoCmd.MoveSize 9500, 6000, 9000, 3325
    Me.txtUserName.SetFocus
    Me.txtPassword.SetFocus
    Me.txtuser.SetFocus
    Me.txtuser2.SetFocus
    Me.txtuser3.SetFocus
     '// load the form in full  transparent mode
    Call SetTranslucent(Me.hwnd, 0)
'    '// ensure the form is visible for proper fade effect
'    '// (thought still full transparent)
    Me.Visible = True
'    '// process fade in effect and provide the opacity value
'    '// not to exceed 255
    Call UIProcessFadeIn(Me, 255)
End Sub

Private Sub btnLogin_Click()
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset("tbl1Employees", dbOpenSnapshot, dbReadOnly)
    
    rs.FindFirst "UserAlias ='" & Me.UserAlias & "'"
       
    If rs.NoMatch Then
        Me.lblWrongUser.Visible = True
        Me.txtUserName.SetFocus
        If MsgBox("Would you like to create an account?", vbYesNo + vbQuestion) = vbYes Then
        DoCmd.OpenForm "frmAccessCode"
        End If
        Exit Sub
    End If
    Me.lblWrongUser.Visible = False
    
    If rs!Password <> Encrypt(Me.txtPassword) Then
        Me.lblWrongPass.Visible = True
        Me.txtPassword.SetFocus
        If MsgBox("Would you like to create an account?", vbYesNo + vbQuestion) = vbYes Then
        DoCmd.OpenForm "frmAccessCode"
        End If
        Exit Sub
    End If
    Me.lblWrongPass.Visible = False
    
    TempVars("UserName").value = Me.txtUserName.value
    
    TempVars("UserComputerName").value = Me.txtuser.value
    
    TempVars("ComputerName").value = Me.txtuser3.value
    
    TempVars("UserDomain").value = Me.txtuser2.value
    
'    Dim DomainName As String
'    Dim ComputerName As String
'    Dim UserComputerName As String
'    DomainName = Environ$("UserDomain")
'    ComputerName = Environ$("ComputerName")
'    UserComputerName = Environ$("username")

    TempVars("EmployeeType") = rs!EmployeeType_ID.value
    
    If rs!EmployeeType_ID = 3 Then
        Dim prop As Property
        On Error GoTo SetProperty
        Set prop = CurrentDb.CreateProperty("AllowBypassKey", dbBoolean, False)
        
'        CurrentDb.Properties.Append prop
        
SetProperty:
    If MsgBox("Would you like to turn on the Bypass Key?", vbYesNo, "Allow Bypass") = vbYes Then
        CurrentDb.Properties("AllowBypassKey") = True
    Else
        CurrentDb.Properties("AllowBypassKey") = False
    End If
    
    End If
    
    Me.Visible = False
    
    DoCmd.OpenForm "frmDemo", acNormal, , , , acHidden
    DoCmd.OpenForm "frmLogoutTimer", acNormal, , , , acHidden
    
    Globals.Logging "Login QAPI"
    
    DoCmd.OpenForm "Splash Screen"
'    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'    Dim NewIcon As String, Icon
'
'
'NewIcon = "\\esh15\QAPI Committee\QAPI Data\QAPI Backup Documents\Backend\BE\dll\QAPI.ico"
'
'Icon = makeicon.ExtractIcon32(0, NewIcon, 0)
'
'makeicon.SendMessage32 makeicon.GetActiveWindow32(), &H80, 0, Icon 'big Icon
'makeicon.SendMessage32 makeicon.GetActiveWindow32(), &H80, 1, Icon 'small Icon

    If Dir(Left$(CurrentDb.Name, Len(CurrentDb.Name) - 3) & "laccdb") <> "" Then
    MsgBox "Application already open"
    DoCmd.Quit
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Globals.Logging "Logoff QAPI"
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





