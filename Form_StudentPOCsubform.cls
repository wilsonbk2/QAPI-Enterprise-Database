VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StudentPOCsubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub POC_FNAME_DblClick(Cancel As Integer)
DoCmd.OpenForm "MiniMenuPOCDialog", , , "[POCID]='" & Me.POCID & "'"
End Sub
