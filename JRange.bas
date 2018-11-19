Attribute VB_Name = "JRange"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * @param {Range} Rng
' *
' */
Public Sub ClearAllBorders(Rng As Range)
    With Rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub
'/**
' * Get Range's name out of its child object
' * @param {Range} Rng
' * @returns {String}
' *
' */
Public Function GetName(Rng As Range) As String
    GetName = Rng.Name.Name
End Function
'/**
' * Check if Range is within another
' * @param {Range} SrcRange
' * @param {Range} TargetRange
' * @returns {Boolean}
' *
' */
Public Function IsInRange(SrcRange As Range, TargetRange As Range) As Boolean
    IsInRange = Not (Application.Intersect(SrcRange, TargetRange) Is Nothing)
End Function
'/**
' * Resize named range (expands without limits, shrinks to the minimum size of one cell)
' * @param {Range} NamedRange
' * @param {Integer} Rows - Number of rows to add/remove, ommit to keep
' * @param {Integer} Columns - Number of columns to add/remove, ommit to keep
' * @requires GetName()
' *
' */
Public Sub ResizeNamedRange(NamedRange As Range, _
                            Optional Rows As Integer = 0, _
                            Optional Columns As Integer = 0)
    Dim nme As Name: Set nme = ThisWorkbook.Names.Item(GetName(NamedRange))
    Dim lngRows, lngColumns As Long
    lngRows = 1 ' Min y-size
    lngColumns = 1 ' Min x-size
    If NamedRange.Rows.Count + Rows > 1 Then lngRows = NamedRange.Rows.Count + Rows
    If NamedRange.Columns.Count + Columns > 1 Then lngColumns = NamedRange.Columns.Count + Columns
    nme.RefersTo = nme.RefersToRange.Resize(lngRows, lngColumns)
    Set nme = Nothing ' ?
End Sub
