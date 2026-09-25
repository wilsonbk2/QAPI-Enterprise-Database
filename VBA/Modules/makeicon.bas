Attribute VB_Name = "makeicon"
Option Compare Database

Private Declare PtrSafe Function GetActiveWindow32 Lib "user32" Alias "GetActiveWindow" () As Integer

Private Declare PtrSafe Function SendMessage32 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare PtrSafe Function ExtractIcon32 Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

