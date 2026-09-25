Attribute VB_Name = "AutoResize"
Option Compare Database
Option Explicit





Public Function Form_Load()
    'Set the font zoom setting to the default of 100% (represented by a 1 below).
    'This means that the fonts will appear initially at the proportional size
    'set during design time. But they can be made smaller or larger at run time
    'by holding the "Shift" key and hitting the "+" or "-" key at the same time,
    'or by holding the "Ctrl" key and scrolling the mouse wheel up or down.
    fontZoom = 1

    'When the form loads, we need to find the relative position of each control
    'and save it in the control's "Tag" property so the resize event can use it
    SaveControlPositionsToTags Me
End Function















