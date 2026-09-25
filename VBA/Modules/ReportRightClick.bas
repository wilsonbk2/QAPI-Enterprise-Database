Attribute VB_Name = "ReportRightClick"
Option Compare Database

Public Function Create2013_InitialRightClickMenu()
On Error Resume Next

    CommandBars("2013_InitRightClick").DELETE

    Dim cmb As CommandBar 'Object
    Dim cmdBtn01 As CommandBarButton 'Object
    Dim cmdBtn02 As CommandBarButton 'Object
    Dim cmdBtn03 As CommandBarButton 'Object
    Dim cmdBtn04 As CommandBarButton 'Object
    Dim cmdBtn05 As CommandBarButton 'Object
    Dim cmdBtn06 As CommandBarButton 'Object
    Dim cmdBtn07 As CommandBarButton 'Object
    Dim cmdBtn08 As CommandBarButton 'Object
    Dim cmdBtn09 As CommandBarButton 'Object
    Dim cmdBtn10 As CommandBarButton 'Object
    Dim cmdBtn11 As CommandBarButton 'Object
    Dim cmdBtn12 As CommandBarButton 'Object
    Dim cmdBtn13 As CommandBarButton 'Object

    
    '
    Set cmb = CommandBars.Add("2013_InitRightClick", _
               msoBarPopup, False, False)    ' msoBarPopup = 5
        With cmb
              ' add cut, copy, and paste buttons with the "magic number" technique that assigns
              ' appearance and behavior. The magic number goes in as the second parameter


' add customized buttons with our caption and function name -- second param is blank
            Set cmdBtn01 = .Controls.Add(msoControlButton, , , , True)
   ' 01
            With cmdBtn01
                .BeginGroup = True
                .Caption = "View Or Edit Payments For This Property"
                                                .OnAction = "=OpenPayment_Pre()"
            End With
    ' 02
            Set cmdBtn02 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn02
                                        .Caption = "Make A New Payment For This Property"
                                                                        .OnAction = "=OpenIndivPayFormInAddMode ()"
            End With
    ' 03
            Set cmdBtn03 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn03
                                        .Caption = "Tax History Report"
                                                                        .OnAction = "=TaxHistoryReport()"
            End With
    ' 04
            Set cmdBtn04 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn04
                                        .Caption = "Log Manual Phone Call for This Tax Record"
                                                                         .OnAction = "=OpenSingleManualCallLog ()"
            End With
    ' 05
            Set cmdBtn05 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn05
                                        .Caption = "Payoff Calculator"
                                                                          .OnAction = "Tax_Menu_Options.PayoffCalculator"
            End With

    ' 06
            Set cmdBtn06 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn06
                                        .Caption = "Tax Statement"
                                                                          .OnAction = "=PrintStatementOnly ()"
            End With
    ' 07
            Set cmdBtn07 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn07
                                       .Caption = "Tax Statement, Letter and Envelope"
                                                                           .OnAction = "=PrintStatementAndLetter ()"
            End With
    ' 08
            Set cmdBtn08 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn08
                                         .Caption = "View / Revise Associated Property Record"
                                                                          .OnAction = "= OpenPropertyRecord()"
            End With
    ' 09
            Set cmdBtn09 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn09
                                         .Caption = "Collections Report - Payment Ledger By BRT"
                                                                            .OnAction = "Tax_Menu_Options.PayLedger_ByBRT"
            End With

    ' 10
            Set cmdBtn10 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn10
                                        .Caption = "Balancing Report - Payments By Deposit Number"
                                                                             .OnAction = "Tax_Menu_Options.PayLedger_ByDeposit"
            End With
    ' 11
            Set cmdBtn11 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn11
                                      .Caption = "Recalculate Balance Due For This Record"
                                                                             .OnAction = "=RecalcBalanceDue()"
            End With

    ' 12
            Set cmdBtn12 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn12
                                        .Caption = "View Or Create Install Pay Plan"
                                                                             .OnAction = "=OpenInstallPay_Form ()"
            End With
    ' 13
            Set cmdBtn13 = .Controls.Add(msoControlButton, , , , True)
            With cmdBtn13
                                      .Caption = "Print Install Plan Letter, Agreement and Coupons"
                                                                              .OnAction = "=PrintInstallPay_Ltr_Agrmnt_Coup ()"
            End With
   
        End With

'

End Function



Public Sub CreateReportRightClickMenu()

On Error Resume Next
    'for runtime versions, report preview right click option to print is not available,
    'so we need to create a shortcut menu bar here and then assign it in each report property sheet "Shortcut menu bar"
    'just need to create once - can run from immediate window by 'call CreateReportShortcutMenu'
    'can't seem to delete so if you need to change, you will need to create new one with different name
    'to delete - try commandbars.item("xxx").delete
    
  '  CommandBars("2013_ReportRightClick").Delete
    CommandBars("ReportRightClick").DELETE
    
    Dim cmbRightClick As Office.CommandBar
    Dim cmbControl As Office.CommandBarControl

    Set cmbRightClick = CommandBars.Add("ReportRightClick", _
                                        msoBarPopup, False, False)
    With cmbRightClick
        
        ' Add the Print command.
        Set cmbControl = .Controls.Add(msoControlButton, 2521)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Quick Print"
        
        ' Add the Print command.
        Set cmbControl = .Controls.Add(msoControlButton, 15948)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Print ..."
        
'        ' Add the Page Setup... command.
'        Set cmbControl = .Controls.Add(msoControlButton, 247)
'        ' Change the caption displayed for the control.
'        cmbControl.Caption = "Page Setup ..."
        
        ' Add the Mail Recipient (as Attachment)... command.
        Set cmbControl = .Controls.Add(msoControlButton, 2188)
        ' Start a new group.
        cmbControl.BeginGroup = True
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Email Report as an Attachment"
        
'        ' Add the Export To Excel.
'        Set cmbControl = .Controls.Add(msoControlButton, 11723)
'        ' Change the caption displayed for the control.
'        cmbControl.Caption = "Export To Excel"
'
'        ' Add the PDF or XPS command.
'        Set cmbControl = .Controls.Add(msoControlButton, 11725)
'        ' Change the caption displayed for the control.
'        cmbControl.Caption = "Export To Word"
        
        
        ' Add the PDF or XPS command.
        Set cmbControl = .Controls.Add(msoControlButton, 12499)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Save as PDF/XPS"
        
        
'        ' Add the Close command.
'        Set cmbControl = .Controls.Add(msoControlButton, 923)
'        ' Start a new group.
'        cmbControl.BeginGroup = True
'        ' Change the caption displayed for the control.
'        cmbControl.Caption = "Close Report"
    End With
    
    Set cmbControl = Nothing
    Set cmbRightClick = Nothing
End Sub
'
'
'
'End Function



