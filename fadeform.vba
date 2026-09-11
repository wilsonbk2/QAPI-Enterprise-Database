Option Compare Database
Option Explicit

'This vba code uses both 32 and 64 bit architecture in order to handle both older and newer OS systems at that time.
'The codes usage was mainly to polish the final product making the interface fade in and out, and make opening newer pages smoothly fade as well.

#If Win64 Then
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_LAYERED = &H80000

Public Const WS_EX_TRANSPARENT = &H20&

Public Const LWA_ALPHA = &H2&

'Enum for determining the direction of the fade.

Public Enum FadeDirection

   FadeIn = -1

   Fadeout = 0

   Fadezero = 1

   SetOpacity = 1

End Enum

Public Sub FadeForm(frm As Form, Optional Direction As FadeDirection = FadeDirection.FadeIn, _
Optional iDelay As Integer = 0, Optional StartOpacity As Long = 5)

   Dim lOriginalStyle As Long
   Dim iCtr As Integer
   'You can only set a form's opacity if it's Popup property = True.
   If (frm.PopUp = True) Then
       'Get the form window’s handle, and remember its original style.
       lOriginalStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
       SetWindowLong frm.hwnd, GWL_EXSTYLE, lOriginalStyle Or WS_EX_LAYERED
       'If the form’s original style = 0, it hasn’t been faded since it was opened.
       'To get fading to work, we have to set its style to something other than zero.
       If (lOriginalStyle = 0) And (Direction <> FadeDirection.SetOpacity) Then
          'Recursively call this same procedure to set the value.
          FadeForm frm, SetOpacity, , StartOpacity
       End If
       'Depending on the direction of the fade...
       Select Case Direction
          Case FadeDirection.Fadezero
              iCtr = StartOpacity
              SetLayeredWindowAttributes frm.hwnd, 0, CByte(iCtr), LWA_ALPHA
          Case FadeDirection.FadeIn
              'Just in case.
              If StartOpacity < 1 Then StartOpacity = 1
              'Fade the form in by varying its opacity
              'from the value supplied in 'StartOpacity'
              'to 255 (completely opaque).
              For iCtr = StartOpacity To 255 Step 1
                 SetLayeredWindowAttributes frm.hwnd, 0, CByte(iCtr), LWA_ALPHA
                 'Process any outstanding events.
                 DoEvents
                 'Wait a while, so the user can see the effect.
                 Sleep iDelay
              Next
          Case FadeDirection.Fadeout
              'Just in case.
              If StartOpacity < 6 Then StartOpacity = 255
              'Fade the form out by varying its opacity
              'from 255 to 1 (almost transparent).
              For iCtr = StartOpacity To 1 Step -1
                 SetLayeredWindowAttributes frm.hwnd, 0, CByte(iCtr), LWA_ALPHA
                 'Process any outstanding events.
                 DoEvents
                 'Wait a while, so the user can see the effect.
                 Sleep iDelay
              Next
          Case Else 'FadeDirection.SetOpacity.
              'Just in case.
              Select Case StartOpacity
                 Case Is < 1: StartOpacity = 1
                 Case Is > 255: StartOpacity = 255
              End Select
              'Set the form's opacity to a specific value.
              SetLayeredWindowAttributes frm.hwnd, 0, CByte(StartOpacity), LWA_ALPHA
                 'Process any outstanding events.
                 DoEvents
                 'Wait a while, so the user can see the effect.
                 Sleep iDelay
       End Select
   Else
       'The form’s Popup property MUST = True
       DoCmd.Beep
       MsgBox "The form's Popup property must be set to True.", vbOKOnly & vbInformation, "Cannot fade form"
   End If

End Sub
