VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo Err_Handler
   DoCmd.Close acForm, Me.Name
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

   DoCmd.MoveSize 200, 200

   'DoCmd.OpenForm "frmLogoutTimer", , , , , acHidden
   DoCmd.OpenForm "frmLogoutTimer", , , , , acWindowNormal
   
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub

Private Sub chkLogOutAllUsers_AfterUpdate()
On Error GoTo Err_Handler
   DoCmd.RunCommand acCmdSaveRecord
   Me.cmdClose.SetFocus
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub
