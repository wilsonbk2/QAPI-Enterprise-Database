VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_B&GInspectionsD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Combo13_BeforeUpdate(Cancel As Integer)
'If Me.DATECOMP = Null Then
'MsgBox "Please fill in completion date to change status", vbCritical
'Me.DATECOMP.BorderColor = red
'Else
'End If
End Sub

Private Sub DATECOMP_AfterUpdate()
Me.Combo13.Locked = False
End Sub
