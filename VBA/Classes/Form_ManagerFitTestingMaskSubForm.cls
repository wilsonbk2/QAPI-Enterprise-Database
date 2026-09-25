VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ManagerFitTestingMaskSubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub Form_Load()
'Dim varX As Variant
'varX = DLookup("[Full Name]", "SATELLITE", "[EMP ID] = '" & Me.EMPID & "'")
'
'Me.Full_Name = varX
'
'End Sub

Private Sub Combo28_AfterUpdate()
If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
    Requery
    Else
    Me.Undo
    End If
End Sub


