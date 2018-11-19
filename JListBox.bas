Attribute VB_Name = "JListBox"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Populate ListBox with Range of values while skipping empty cells (ignores LstBox.Locked = True)
' * @param {MSForms.ListBox} LstBox - Sheet.lstName
' * @param {Range} SrcRange - Range of values.
' * @param {Integer} SrcColumn – Append data in first column of given range or specify
' *
' */
Public Sub PopulateWithRange(ByRef LstBox As MSForms.ListBox, _
                             ByVal SrcRange As Range, _
                             Optional ByVal SrcColumn As Integer = 1)
    Dim rngCell As Range
    LstBox.Clear
    For Each rngCell In SrcRange.Columns(SrcColumn).Cells
        If rngCell <> vbNullString Then
            Call LstBox.AddItem(rngCell)
        End If
    Next rngCell
End Sub

