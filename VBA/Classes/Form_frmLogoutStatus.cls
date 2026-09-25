VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLogoutStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mdat_StartCountdownTime As Date

Private Sub cmdOK_Click()
   On Error GoTo Err_Handler
   Me.Visible = False
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Private Sub Form_Open(iCancel As Integer)
   On Error GoTo Err_Handler
   
   mdat_StartCountdownTime = DateAdd("n", 3, Now())
   Me!Thetimer = " 3  minutes and  0  seconds "
   
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Private Sub Form_Timer()
   On Error GoTo Err_Handler
   
   Dim intIMins As Integer
   Dim intISecs As Integer
   Dim fLogout As Boolean
   
   fLogout = DLookup("[LogOutAllUsers]", "[tblVersionServer]")
   
   If fLogout = False Then DoCmd.Close acForm, Me.Name
   
   intIMins = DateDiff("s", Now(), mdat_StartCountdownTime) \ 60
   intISecs = DateDiff("s", Now(), DateAdd("n", (intIMins * -1), mdat_StartCountdownTime))
   
   If intIMins = 2 And intISecs = 0 Then
      Me.Visible = True
   End If
   If intIMins = 1 And intISecs = 0 Then
      Me.Visible = True
   End If
   If intIMins = 0 And intISecs = 20 Then
      Me.Visible = True
   End If
   
   If intIMins <= 0 And intISecs <= 0 Then
      DoCmd.Quit
   Else
      On Error Resume Next
'      Me.txtMinsSecs = " " & intIMins & "  minutes and  " & intISecs & "  seconds "
      Me.Thetimer = " " & intIMins & " minutes and " & intISecs & "  seconds "
   End If
   
   
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub




