Attribute VB_Name = "ShowDB"
Option Compare Database
Option Explicit

'************ Code Start **********
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3

'###############################################
#If VBA7 Then
    Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#ElseIf Win64 Then 'need datatype LongPtr
    Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
#Else '32-bit Office
    Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If
'###############################################

Function SetAccessWindow(nCmdShow As Long)

    'Usage Examples
    'Maximize window:
    ' ?SetAccessWindow(SW_SHOWMAXIMIZED)
    'Minimize window:
    ' ?SetAccessWindow(SW_SHOWMINIMIZED)
    'Hide window:
    ' ?SetAccessWindow(SW_HIDE)
    'Normal window:
    ' ?SetAccessWindow(SW_SHOWNORMAL)
    
    Dim loX As Long
   ' Dim loForm As Form
    On Error Resume Next
    
    loX = apiShowWindow(hWndAccessApp, nCmdShow)
    SetAccessWindow = (loX <> 0)

End Function
