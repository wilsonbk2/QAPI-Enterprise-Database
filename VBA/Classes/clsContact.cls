VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event ShowMessage(Message As String)

Private p_ID As Integer
Private p_CustomerID As Integer
Private p_FirstName As String
Private p_MiddleName As String
Private p_LastName As String
Private p_Email As String

Private Sub class_terminate()
    Globals.g_Enabled = False
End Sub

Public Property Get ID() As Integer
    ID = p_ID
End Property
Public Property Let ID(value As Integer)
    If value > 0 Then
        p_ID = value
    End If
End Property

Public Property Let firstname(value As String)
    p_FirstName = value
End Property
Public Property Get firstname() As String
    firstname = p_FirstName
End Property

Public Sub init(Optional ID As Integer = 1, Optional firstname As String = "Brad")
'    RaiseEvent ShowMessage("Contact Class Initializing")
    p_ID = ID
    p_FirstName = firstname
'    RaiseEvent ShowMessage("Contact Class Initialized")
End Sub

