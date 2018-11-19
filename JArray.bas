Attribute VB_Name = "JArray"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
' Option Compare Text
Option Explicit
Option Private Module
'/**
' * Sort Array of strings in ascending order
' * @param {Variant} Arr - Array of strings
' * @link http://www.java2s.com/Code/VBA-Excel-Access-Word/Data-Type/SortstheListarrayinascendingorder.htm
' *
' */
Public Sub BubbleSort(ByRef Arr As Variant)
    Dim i, j As Integer
    Dim k As Variant
    For i = LBound(Arr) To UBound(Arr) - 1
        For j = i + 1 To UBound(Arr)
            If Arr(i) > Arr(j) Then
                k = Arr(j)
                Arr(j) = Arr(i)
                Arr(i) = k
            End If
        Next j
    Next i
End Sub
'/**
' * @param {Variant} Arr
' * @returns {Long}
' *
' */
Public Function GetLength(Arr As Variant) As Long
    GetLength = UBound(Arr) - LBound(Arr) + 1
End Function
