VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccessCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnLogin_Click()
If Me.AccessConfirm.value = Me.AccessExact.value Then
    DoCmd.OpenForm "frmCreateAccount"
    DoCmd.Close acForm, "frmAccessCode"
    Application.SetOption "Confirm Record Changes", 0
    Application.SetOption "Confirm Action Queries", 0
    Application.SetOption "Confirm Document Deletions", 0
    CurrentDb.Properties("ShowDocumentTabs").value = False
    Else
    MsgBox "Incorrect Access Code.", vbCritical
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

Private Sub Form_Load()
DoCmd.MoveSize 9500, 6000, 4500, 2000
 '// load the form in full  transparent mode
    Call SetTranslucent(Me.hwnd, 0)
    '// ensure the form is visible for proper fade effect
    '// (thought still full transparent)
    Me.Visible = True
    '// process fade in effect and provide the opacity value
    '// not to exceed 255
    Call UIProcessFadeIn(Me, 255)
End Sub
