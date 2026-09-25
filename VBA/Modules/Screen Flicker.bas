Attribute VB_Name = "Screen Flicker"
Option Compare Database
Option Explicit


Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SETREDRAW As Long = 11

