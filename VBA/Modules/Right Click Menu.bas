Attribute VB_Name = "Right Click Menu"
Option Compare Database


Public Sub CreateMain_copypastecutasc1()

On Error Resume Next
    'for runtime versions, report preview right click option to print is not available,
    'so we need to create a shortcut menu bar here and then assign it in each report property sheet "Shortcut menu bar"
    'just need to create once - can run from immediate window by 'call CreateReportShortcutMenu'
    'can't seem to delete so if you need to change, you will need to create new one with different name
    'to delete - try commandbars.item("xxx").delete
    
  '  CommandBars("2013_ReportRightClick").Delete
    CommandBars("MainRightClick").DELETE
    
    Dim cmbRightClick As Office.CommandBar
    Dim cmbControl As Office.CommandBarControl

    Set cmbRightClick = CommandBars.Add("MainRightClick", _
                                        msoBarPopup, False, False)
    With cmbRightClick
        
        ' Add the Print command.
        Set cmbControl = .Controls.Add(msoControlButton, 21)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Cut"
        
        ' Add the Print command.
        Set cmbControl = .Controls.Add(msoControlButton, 19)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Copy"
        
'        ' Add the Page Setup... command.
'        Set cmbControl = .Controls.Add(msoControlButton, 247)
'        ' Change the caption displayed for the control.
'        cmbControl.Caption = "Page Setup ..."
      
        ' Add the Mail Recipient (as Attachment)... command.
        Set cmbControl = .Controls.Add(msoControlButton, 22)
        ' Start a new group.
    
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Paste"
       
        ' Add the Export To Excel.
        Set cmbControl = .Controls.Add(msoControlButton, 210)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Sort Ascending"
         cmbControl.BeginGroup = True
        ' Add the PDF or XPS command.
        Set cmbControl = .Controls.Add(msoControlButton, 211)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Sort Descending"
        ' Add the Export To Excel.
'        Set cmbControl = .Controls.Add(msoControlButton, 11723)
'        ' Change the caption displayed for the control.
'        cmbControl.Caption = "Export To Excel"
'        cmbControl.BeginGroup = True
        
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



