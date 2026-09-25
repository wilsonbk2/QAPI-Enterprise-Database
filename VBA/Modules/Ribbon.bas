Attribute VB_Name = "Ribbon"
Option Compare Database
Function LoadRibbon()
Dim sXML
  sXML = "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
      "<ribbon startFromScratch=""false"">" & _
      "</ribbon>" & _
      "<backstage>" & _
      "<button idMso=""FileSave"" visible=""false""/>" & _
      "<button idMso=""SaveObjectAs"" visible=""false""/>" & _
      "<button idMso=""FileSaveAsCurrentFileFormat"" visible=""false""/>" & _
      "<button idMso=""FileOpen"" visible=""false""/>" & _
      "<button idMso=""FileCloseDatabase"" visible=""false""/>" & _
      "<tab idMso =""TabInfo"" visible=""false""/>" & _
      "<tab idMso =""TabRecent"" visible=""false""/>" & _
      "<tab idMso =""TabNew"" visible=""false""/>" & _
      "<tab idMso =""TabPrint"" visible=""false""/>" & _
      "<tab idMso =""TabShare"" visible=""false""/>" & _
      "<tab idMso =""TabHelp"" visible=""false""/>" & _
      "<button idMso=""ApplicationOptionsDialog"" visible=""false""/>" & _
      "<button idMso=""FileExit"" visible=""true""/>" & _
      "</backstage>" & _
      "</customUI>"
  MsgBox "Hallo"
  Application.LoadCustomUI "SimpleFile", sXML
End Function

Function LoadRibbon2()
Dim sXML
  sXML = "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
      "<ribbon startFromScratch=""false"">" & _
      "</ribbon>" & _
      "<backstage>" & _
      "<button idMso=""ApplicationOptionsDialog"" visible=""false""/>" & _
      "</backstage>" & _
      "</customUI>"
  Application.LoadCustomUI "SimpleFile2", sXML
End Function

