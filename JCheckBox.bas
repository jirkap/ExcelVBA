Attribute VB_Name = "JCheckBox"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * @param {MSForms.CheckBox} Controls() - Array of checkboxes
' * @link http://stackoverflow.com/questions/37028654/passing-array-of-sheet-controls-to-modify-their-properties
' * @todo This does not work with UserForm checkboxes
' *
' */
Public Sub Check(ByRef Controls() As MSForms.CheckBox)
    Dim i As Long
    For i = LBound(Controls) To UBound(Controls)
        Controls(i) = True
    Next i
End Sub
'/**
' * @param {MSForms.CheckBox} Controls() - Array of checkboxes
' * @link http://stackoverflow.com/questions/37028654/passing-array-of-sheet-controls-to-modify-their-properties
' * @todo This does not work with UserForm checkboxes
' *
' */
Public Sub Uncheck(ByRef Controls() As MSForms.CheckBox)
    Dim i As Long
    For i = LBound(Controls) To UBound(Controls)
        Controls(i) = False
    Next i
End Sub
'/**
' * @param {MSForms.CheckBox} Controls() - Array of checkboxes
' * @link http://stackoverflow.com/questions/37028654/passing-array-of-sheet-controls-to-modify-their-properties
' * @todo This does not work with UserForm checkboxes
' *
' */
Public Sub Disable(ByRef Controls() As MSForms.CheckBox)
    Dim i As Long
    For i = LBound(Controls) To UBound(Controls)
        Controls(i).Enabled = False
    Next i
End Sub
'/**
' * @param {MSForms.CheckBox} Controls() - Array of checkboxes
' * @link http://stackoverflow.com/questions/37028654/passing-array-of-sheet-controls-to-modify-their-properties
' * @todo This does not work with UserForm checkboxes
' *
' */
Public Sub Enable(ByRef Controls() As MSForms.CheckBox)
    Dim i As Long
    For i = LBound(Controls) To UBound(Controls)
        Controls(i).Enabled = True
    Next i
End Sub
