VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Admin Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'Set an unchangeable variable to the amount (10% for example) to increase or
'decrease the font size with each zoom, in or out.
Const FONT_ZOOM_PERCENT_CHANGE = 0.1


'Create the fontZoom and ctrlKeyIsPressed variables outside of
'the sub definitions so they can be shared between subs
Private fontZoom As Double
Private ctrlKeyIsPressed As Boolean


'Create an enum so we can use it later when pulling the data out of the "Tag" property
Private Enum ControlTag
    FromLeft = 0
    FromTop
    ControlWidth
    ControlHeight
    OriginalFontSize
    OriginalControlHeight
End Enum


Private Sub Command40_Click()
Me.Visible = False
End Sub

Private Sub Form_OnError(DataErr As Integer, Response As Integer)
problem:
MsgBox "test"
MsgBox "Test", vbCritical
End Sub


Private Sub Form_Load()
    'Set the font zoom setting to the default of 100% (represented by a 1 below).
    'This means that the fonts will appear initially at the proportional size
    'set during design time. But they can be made smaller or larger at run time
    'by holding the "Shift" key and hitting the "+" or "-" key at the same time,
    'or by holding the "Ctrl" key and scrolling the mouse wheel up or down.
    fontZoom = 1

    'When the form loads, we need to find the relative position of each control
    'and save it in the control's "Tag" property so the resize event can use it
    SaveControlPositionsToTags Me
    
    Dim notnum As Long
notnum = IsNull(DLookup("[BReason]", "[Bug Resolve]"))
If notnum Then
    Me.NOT1.Visible = False
    Me.NOT2.Visible = False
    Else
    Me.NOT1.Visible = True
    Me.NOT2.Visible = True
End If

End Sub


Private Sub Form_Resize()
    'Set the height of the header and footer before calling RepositionControls
    'since it caused problems changing their heights from inside that sub.
    'The Tag property for the header and footer is set inside the SaveControlPositionsToTags sub
    Me.Section(acHeader).Height = Me.WindowHeight * CDbl(Me.Section(acHeader).Tag)
    Me.Section(acFooter).Height = Me.WindowHeight * CDbl(Me.Section(acFooter).Tag)

    'Call the RepositionControls Sub and pass this form as a parameter
    'and the fontZoom setting which was initially set when the form loaded and then
    'changed if the user holds the "Shift" key and hits the "+" or "-" key
    'or holds the "Ctrl" key and scrolls the mouse wheel up or down.
    RepositionControls Me, fontZoom
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, SHIFT As Integer)

    'PURPOSE: Make the text on the form bigger if "Shift" and "+" are pressed
    'at the same time and smaller if "Shift" and "-" are pressed at the same time.
    'NOTE: Using the "Ctrl" key instead of the "Shift" key conflicts with Access's
    'default behavior of using "Ctrl -" to delete a record, so "Shift" is used instead

    'Was the "Shift" key being held down while the Key was pressed?
    Dim shiftKeyPressed As Boolean
    shiftKeyPressed = (SHIFT And acShiftMask) > 0

    'If so, check to see if the user pressed the "+" or the "-" button at the
    'same time as the "Shift" key. If so, then make the font bigger/smaller
    'by the percentage specificed in the FONT_ZOOM_PERCENT_CHANGE variable.
    If shiftKeyPressed Then

        Select Case KeyCode
            Case vbKeyAdd
                fontZoom = fontZoom + FONT_ZOOM_PERCENT_CHANGE
                RepositionControls Me, fontZoom

                'Set the KeyCode back to zero to prevent the "+" symbol from
                'showing up if a textbox or similar control has the focus
                KeyCode = 0

            Case vbKeySubtract
                fontZoom = fontZoom - FONT_ZOOM_PERCENT_CHANGE
                RepositionControls Me, fontZoom

                'Set the KeyCode back to zero to prevent the "-" symbol from
                'showing up if a textbox or similar control has the focus
                KeyCode = 0

        End Select

    End If

    'Detect if the "Ctrl" key was pressed. This variable
    'will be used later when we detect a mouse wheel scroll event.
    If (SHIFT And acCtrlMask) > 0 Then
        ctrlKeyIsPressed = True
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, SHIFT As Integer)
    'Change the ctrlKeyIsPressed variable to false when
    'any key is let up. This will make sure the form text does
    'not continue to grow/shrink when the mouse wheel is
    'scrolled after the ctrl key is pressed and let up.
    ctrlKeyIsPressed = False
End Sub


Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    'If the "Ctrl" key is also being pressed, then zoom the form in or out
    If ctrlKeyIsPressed Then
        Debug.Print ctrlKeyIsPressed
        'The user scrolled up, so make the text larger
        If Count < 0 Then

            'Make the font bigger by the percentage specificed
            'in the FONT_ZOOM_PERCENT_CHANGE variable
            fontZoom = fontZoom + FONT_ZOOM_PERCENT_CHANGE
            RepositionControls Me, fontZoom

        'The user scrolled down, so make the text smaller
        ElseIf Count > 0 Then

            'Make the font smaller by the percentage specificed
            'in the FONT_ZOOM_PERCENT_CHANGE variable
            fontZoom = fontZoom - FONT_ZOOM_PERCENT_CHANGE
            RepositionControls Me, fontZoom
        End If

    End If
End Sub


Public Sub SaveControlPositionsToTags(frm As Form)
On Error Resume Next

    Dim ctl As Control

    Dim ctlLeft As String
    Dim ctlTop As String
    Dim ctlWidth As String
    Dim ctlHeight As String
    Dim ctlOriginalFontSize As String
    Dim ctlOriginalControlHeight As String

    For Each ctl In frm.Controls

        'Find the relative position of this control in design view
        'e.g.- This control is 5% from the left, 10% from the top, etc.
        'Those percentages can then be saved in the Tag property for this control
        'and used later in the form's resize event
        ctlLeft = CStr(Round(ctl.Left / frm.Width, 4))
        ctlTop = CStr(Round(ctl.Top / frm.Section(ctl.Section).Height, 4))
        ctlWidth = CStr(Round(ctl.Width / frm.Width, 4))
        ctlHeight = CStr(Round(ctl.Height / frm.Section(ctl.Section).Height, 4))

        'If this control has a FontSize property, then capture the
        'control's original font size and the control's original height from design-time
        'These will be used later to calculate what the font size should be when the form is resized
        Select Case ctl.ControlType
            Case acLabel, acCommandButton, acTextBox, acComboBox, acListBox, acTabCtl, acToggleButton
                ctlOriginalFontSize = ctl.FontSize
                ctlOriginalControlHeight = ctl.Height
        End Select

        'Add all this data to the Tag property of the current control, separated by colons
        ctl.Tag = ctlLeft & ":" & ctlTop & ":" & ctlWidth & ":" & ctlHeight & ":" & ctlOriginalFontSize & ":" & ctlOriginalControlHeight

    Next

    'Set the Tag properties for the header and the footer to their proportional height
    'in relation to the height of the whole form (header + detail + footer)
    frm.Section(acHeader).Tag = CStr(Round(frm.Section(acHeader).Height / (frm.Section(acHeader).Height + frm.Section(acDetail).Height + frm.Section(acFooter).Height), 4))
    frm.Section(acFooter).Tag = CStr(Round(frm.Section(acFooter).Height / (frm.Section(acHeader).Height + frm.Section(acDetail).Height + frm.Section(acFooter).Height), 4))

End Sub


Public Sub RepositionControls(frm As Form, fontZoom As Double)
On Error Resume Next

    Dim formDetailHeight As Long
    Dim tagArray() As String

    'Since "Form.Section(acDetail).Height" usually returns the same value (unless the detail section is tiny)
    'go ahead and calculate the detail section height ourselves and store it in a variable
    formDetailHeight = frm.WindowHeight - frm.Section(acHeader).Height - frm.Section(acFooter).Height

    Dim ctl As Control

    'Loop through all the controls on the form
    For Each ctl In frm.Controls

        'An extra check to make sure the Tag property has a value
        If ctl.Tag <> "" Then

            'Split the Tag property into an array
            tagArray = Split(ctl.Tag, ":")

            If ctl.Section = acDetail Then
                'This is the Detail section of the form so use our "formDetailHeight" variable from above
                ctl.Move frm.WindowWidth * (CDbl(tagArray(ControlTag.FromLeft))), _
                                   formDetailHeight * (CDbl(tagArray(ControlTag.FromTop))), _
                                   frm.WindowWidth * (CDbl(tagArray(ControlTag.ControlWidth))), _
                                   formDetailHeight * (CDbl(tagArray(ControlTag.ControlHeight)))
            Else
                ctl.Move frm.WindowWidth * (CDbl(tagArray(ControlTag.FromLeft))), _
                                   frm.Section(ctl.Section).Height * (CDbl(tagArray(ControlTag.FromTop))), _
                                   frm.WindowWidth * (CDbl(tagArray(ControlTag.ControlWidth))), _
                                   frm.Section(ctl.Section).Height * (CDbl(tagArray(ControlTag.ControlHeight)))
            End If

            'Now we need to change the font sizes on the controls.
            'If this control has a FontSize property, then find the ratio of
            'the current height of the control to the form-load height of the control.
            'So if form-load height was 1000 (twips) and the current height is 500 (twips)
            'then we multiply the original font size * (500/1000), or 50%.
            'Then we multiply that by the fontZoom setting in case the user wants to
            'increase or decrease the font sizes while viewing the form.
            Select Case ctl.ControlType
                Case acLabel, acCommandButton, acTextBox, acComboBox, acListBox, acTabCtl, acToggleButton
                    ctl.FontSize = Round(CDbl(tagArray(ControlTag.OriginalFontSize)) * CDbl(ctl.Height / tagArray(ControlTag.OriginalControlHeight))) * fontZoom
            End Select

        End If

    Next

End Sub

Private Sub Command110_Click()

If Globals.VerifyStatus("Exams") And Globals.VerifyStatus("Assessments") = True Then
Me.ADMbtn.BorderColor = vbGreen
Me.ADMchk.Visible = True
Else
Me.ADMbtn.BorderColor = vbRed
Me.ADMchk.Visible = False
End If

If Globals.VerifyStatus("Rounds Audits B1") And Globals.VerifyStatus("Rounds Audits B2") And Globals.VerifyStatus("Patient Complications") And Globals.VerifyStatus("Pressure Sores") = True Then
Me.NRSb1.BorderColor = vbGreen
Me.NRScheck1.Visible = True
Else
Me.NRSb1.BorderColor = vbRed
Me.NRScheck1.Visible = False
End If

If Globals.VerifyStatus("RM Falls") = True Then
Me.RMbtn.BorderColor = vbGreen
Me.RMchk.Visible = True
Else
Me.RMbtn.BorderColor = vbRed
Me.RMchk.Visible = False
End If

If Globals.VerifyStatus("QM CMS") And Globals.VerifyStatus("QMPR") And Globals.VerifyStatus("QMMR") And Globals.VerifyStatus("QMSI") And Globals.VerifyStatus("QMD") = True Then
Me.QMbtn.BorderColor = vbGreen
Me.QMchk.Visible = True
Else
Me.QMbtn.BorderColor = vbRed
Me.QMchk.Visible = False
End If

If Globals.VerifyStatus("Patient Safety Stats") = True Then
Me.PTbtn.BorderColor = vbGreen
Me.PTchk.Visible = True
Else
Me.PTbtn.BorderColor = vbRed
Me.PTchk.Visible = False
End If

'If Globals.VerifyStatus("Diet Stats") = True Then
'Me.DIETbtn.BorderColor = vbGreen
'Me.DIETchk.Visible = True
'Else
'Me.DIETbtn.BorderColor = vbRed
'Me.DIETchk.Visible = False
'End If
Me.DIETbtn.BorderColor = vbBlue

If Globals.VerifyStatus("Patient Medications") And Globals.VerifyStatus("Pharmacy Measures") = True Then
Me.PHARMbtn.BorderColor = vbGreen
Me.PHARMchk.Visible = True
Else
Me.PHARMbtn.BorderColor = vbRed
Me.PHARMchk.Visible = False
End If

If Globals.VerifyStatus("Admis/Discharge") = True Then
Me.HIMbtn.BorderColor = vbGreen
Me.HIMchk.Visible = True
Else
Me.HIMbtn.BorderColor = vbRed
Me.HIMchk.Visible = False
End If

If Globals.VerifyStatus("SDT Completions") = True Then
Me.SDTbtn.BorderColor = vbGreen
Me.SDTchk.Visible = True
Else
Me.SDTbtn.BorderColor = vbRed
Me.SDTchk.Visible = False
End If

If Globals.VerifyStatus("GA By Area") And Globals.VerifyStatus("GA by Unit") = True Then
Me.REHABbtn.BorderColor = vbGreen
Me.REHABchk.Visible = True
Else
Me.REHABbtn.BorderColor = vbRed
Me.REHABchk.Visible = False
End If

If Globals.VerifyStatus("Patient Employment") And Globals.VerifyStatus("Patient Workdays") = True Then
Me.JRCPbtn.BorderColor = vbGreen
Me.JRCPchk.Visible = True
Else
Me.JRCPbtn.BorderColor = vbRed
Me.JRCPchk.Visible = False
End If

If Globals.VerifyStatus("Diversions") = True Then
Me.SWbtn.BorderColor = vbGreen
Me.SWchk.Visible = True
Else
Me.SWbtn.BorderColor = vbRed
Me.SWchk.Visible = False
End If
'Me.SWbtn.BorderColor = vbBlue

If Globals.VerifyStatus("Suicide Assessments") = True Then
Me.PSYbtn.BorderColor = vbGreen
Me.PSYchk.Visible = True
Else
Me.PSYbtn.BorderColor = vbRed
Me.PSYchk.Visible = False
End If

If Globals.VerifyStatus("Laboratory Stats") = True Then
Me.LABbtn.BorderColor = vbGreen
Me.LABchk.Visible = True
Else
Me.LABbtn.BorderColor = vbRed
Me.LABchk.Visible = False
End If

'If Globals.VerifyStatus("") = True Then
'Me.PURbtn.BorderColor = vbGreen
'Me.PURchk.Visible = False
'Else
'Me.PURbtn.BorderColor = vbRed
'End If
Me.PURbtn.BorderColor = vbBlue

If Globals.VerifyStatus("Budget and Prompt Pay") = True Then
Me.FISCALbtn.BorderColor = vbGreen
Me.FISCALchk.Visible = True
Else
Me.FISCALbtn.BorderColor = vbRed
Me.FISCALchk.Visible = False
End If

If Globals.VerifyStatus("MIS Narrative") = True Then
Me.MISbtn.BorderColor = vbGreen
Me.MISchk.Visible = True
Else
Me.MISbtn.BorderColor = vbRed
Me.MISchk.Visible = False
End If

If Globals.VerifyStatus("Hand Hygiene") And Globals.VerifyStatus("IC Stats") = True Then
Me.ICbtn.BorderColor = vbGreen
Me.ICchk.Visible = True
Else
Me.ICbtn.BorderColor = vbRed
Me.ICchk.Visible = False
End If

If Globals.VerifyStatus("B&G Documentation") = True Then
Me.BGbtn.BorderColor = vbGreen
Me.BGchk.Visible = True
Else
Me.BGbtn.BorderColor = vbRed
Me.BGchk.Visible = False
End If

If Globals.VerifyStatus("HREmployee") And Globals.VerifyStatus("HRTurnover") = True Then
Me.HRbtn.BorderColor = vbGreen
Me.HRchk.Visible = True
Else
Me.HRbtn.BorderColor = vbRed
Me.HRchk.Visible = False
End If

If Globals.VerifyStatus("Xray Services") = True Then
Me.Xraybtn.BorderColor = vbGreen
Me.XRcheck.Visible = True
Else
Me.Xraybtn.BorderColor = vbRed
Me.XRcheck.Visible = False
End If

End Sub

Private Sub Command119_Click()
DoCmd.OpenQuery "ResetVerify"
MsgBox "Verification has now been reset for the next month, thank you."
End Sub

Private Sub Command41_Click()
GoToHome.CloseAllForms
End Sub

Private Sub ADMbtn_Click()
DoCmd.OpenForm "ADMISMenu"
End Sub

Private Sub FISCALbtn_Click()
DoCmd.OpenForm "FISCALMenu"
End Sub

Private Sub HIMbtn_Click()
DoCmd.OpenForm "HIMMenu"
End Sub

Private Sub HRbtn_Click()
DoCmd.OpenForm "HRMenu"
End Sub

Private Sub ICbtn_Click()
DoCmd.OpenForm "INFECTMenu"
End Sub

Private Sub JRCPbtn_Click()
DoCmd.OpenForm "JRCPMenu"
End Sub

Private Sub LABbtn_Click()
DoCmd.OpenForm "LABMenu"
End Sub

Private Sub MISbtn_Click()
DoCmd.OpenForm "MISMenu"
End Sub

Private Sub NRSb1_Click()
DoCmd.OpenForm "NursingMenu"
End Sub

Private Sub PHARMbtn_Click()
DoCmd.OpenForm "PHARMMenu"
End Sub

Private Sub PSYbtn_Click()
DoCmd.OpenForm "PSYMenu"
End Sub

Private Sub PTbtn_Click()
DoCmd.OpenForm "PASMenu"
End Sub

Private Sub PURbtn_Click()
DoCmd.OpenForm "PURMenu"
End Sub

Private Sub REHABbtn_Click()
DoCmd.OpenForm "REHABMenu"
End Sub

Private Sub QMbtn_Click()
DoCmd.OpenForm "QMMenu"
End Sub

Private Sub RMbtn_Click()
DoCmd.OpenForm "RMMenu"
End Sub

Private Sub SDTbtn_Click()
DoCmd.OpenForm "SDTMenu"
End Sub

Private Sub SWbtn_Click()
DoCmd.OpenForm "SWMenu"
End Sub

Private Sub Command120_Click()
DoCmd.OpenForm "Expanded Menu"
End Sub

Private Sub COMITbtn_Click()
DoCmd.OpenForm "MISMenu"
End Sub

Private Sub DIETbtn_Click()
DoCmd.OpenForm "DIETMenu"
End Sub

Private Sub BGbtn_Click()
DoCmd.OpenForm "BGMenu"
End Sub

Private Sub Xraybtn_Click()
DoCmd.OpenForm "XRAYMenu"
End Sub
