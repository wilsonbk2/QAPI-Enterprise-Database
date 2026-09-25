Attribute VB_Name = "Resize"
Option Compare Database
''''''CODE START''''''
Private Type Rect
x1 As Long
Y1 As Long
x2 As Long
Y2 As Long
End Type
'

#If Win64 Then
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, r As Rect) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" _
(ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, _
ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" _
(ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#Else
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, r As Rect) As Long
Private Declare Function IsZoomed Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" _
(ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, _
ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" _
(ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If

Sub SizeAccess(ByVal dx As Long, ByVal dy As Long)
'Set size of Access and center on Desktop

Const SW_RESTORE As Long = 9
Dim h As Long
Dim r As Rect
'
On Error Resume Next
'
h = Application.hWndAccessApp
'If maximised, restore
If (IsZoomed(h)) Then ShowWindow h, SW_RESTORE
'
'Get available Desktop size
GetWindowRect GetDesktopWindow(), r
If ((r.x2 - r.x1) - dx) < 0 Or ((r.Y2 - r.Y1) - dy) < 0 Then
'Desktop smaller than requested size
'so size to Desktop
MoveWindow h, r.x1, r.Y1, r.x2, r.Y2, True
Else
'Adjust to requested size and center
MoveWindow h, _
r.x1 + ((r.x2 - r.x1) - dx) \ 2, _
r.Y1 + ((r.Y2 - r.Y1) - dy) \ 2, _
dx, dy, True
End If

End Sub
