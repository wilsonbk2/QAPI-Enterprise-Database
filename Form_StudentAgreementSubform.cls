VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StudentAgreementSubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub PROG_NAME_DblClick(Cancel As Integer)
DoCmd.OpenForm "MiniMenuAgreementDialog", , , "[AGREE-NUM]='" & Me.AGREENUM & "' And [SCH-ID]='" & Me.SCHID & "'"
End Sub
