VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLogoutTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdHide_Click()
   On Error GoTo Err_Handler
   Me.Visible = False
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Private Sub cmdRunNow_Click()
   On Error GoTo Err_Handler
   Me.TimerInterval = 9000
   DoCmd.OpenForm "frmLogoutStatus"
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Private Sub Form_Timer()
   On Error GoTo Err_Handler
   
   Dim fLogout As Boolean
   
   fLogout = DLookup("[LogOutAllUsers]", "[tblVersionServer]")
   If fLogout = True Then
      Me.TimerInterval = 9000
      DoCmd.OpenForm "frmLogoutStatus"
   Else
      Me.TimerInterval = 30000
      If IsLoaded("frmLogoutStatus") Then
         DoCmd.Close acForm, "frmLogoutStatus"
      End If
   End If
   
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Public Function IsLoaded(ByVal sFormName As String) As Boolean
   On Error GoTo Err_Handler
   
   Const conObjStateClosed = 0
   Const conDesignView = 0
   
   If SysCmd(acSysCmdGetObjectState, acForm, sFormName) <> conObjStateClosed Then
      If Forms(sFormName).CurrentView <> conDesignView Then
         IsLoaded = True
      End If
   End If
   
Exit_Here:
   Exit Function
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
   
End Function



