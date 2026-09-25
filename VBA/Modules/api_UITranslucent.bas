Attribute VB_Name = "api_UITranslucent"
Option Compare Database
Option Explicit


Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3

Public Const WS_EX_LAYERED = &H80000

Public Const GWL_EXSTYLE = -20


#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
                        
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
                        ByVal hwnd As Long, _
                        ByVal color As Long, _
                        ByVal bAlpha As Byte, _
                        ByVal alpha As Long) As Boolean
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
                        
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                        ByVal hwnd As Long, _
                        ByVal color As Long, _
                        ByVal bAlpha As Byte, _
                        ByVal alpha As Long) As Boolean
#End If
                        


'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'// PURPOSE:    process opacity/transparency on any particular form. the form's
'//             POPUP property must be set to YES or this will not work
'// PARAMETERS: [in] UIForm - the form we want to fade out
'//             [in] StartOpacity - the final opacity value in which the form
'//                               is to bet set
'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SetTranslucent(hwnd As Long, Opacity As Integer) As Boolean
Dim APIResponse As Long
'// enable error handler
On Error GoTo Err_Handler

    '// put current GWL_EXSTYLE in attrib
    APIResponse = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    '// change GWL_EXSTYLE to WS_EX_LAYERED - makes a window layered
    SetWindowLong hwnd, GWL_EXSTYLE, APIResponse Or WS_EX_LAYERED
    
    '// make transparent (RGB value does not have any effect at this
    SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), Opacity, LWA_ALPHA

Err_Exit:
    Exit Function
    
Err_Handler:
    MsgBox Err.Number & " : " & Err.Description
    
End Function


'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'// PURPOSE:    process fade out effect on any particular form. the form's
'//             POPUP property must be set to YES or this will not work
'// PARAMETERS: [in] UIForm - the form we want to fade out
'//             [in] StartOpacity - the opacity value in which the form was
'//                                 opened, if none was applied, ignore
'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UIProcessFadeOut(UIForm As Form, Optional StartOpacity As Integer = 255)
'// loop counter
Dim i As Integer
    
    For i = StartOpacity To 0 Step -1
        Call SetTranslucent(UIForm.hwnd, i)
        '// this is required for proper fade effect
        '// otherwise you'll just jump to the transparency immediately
        DoEvents
    Next i
End Sub

'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'// PURPOSE:    process fade in effect on any particular form. the form's
'//             POPUP property must be set to YES or this will not work
'// PARAMETERS: [in] UIForm - the form we want to fade out
'//             [in] EndOpacity - the final opacity value in which the form
'//                               is to bet set, if none applied, ignore
'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UIProcessFadeIn(UIForm As Form, Optional EndOpacity As Integer = 255)
'// loop counter
Dim i As Integer

    For i = 1 To EndOpacity
        Call SetTranslucent(UIForm.hwnd, i)
        '// this is required for proper fade effect
        '// otherwise you'll just jump to the transparency immediately
        
        '// you may want to use another method to wait
        DoEvents
    Next i

End Sub






