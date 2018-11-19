Attribute VB_Name = "JTest"
Option Explicit
Option Private Module
Public Function JArray_BubbleSort()
    Dim strArray(1 To 4) As String
    strArray(1) = "New York"
    strArray(2) = "Amsterdam"
    strArray(3) = "Malmö"
    strArray(4) = "Prague"
    Call JArray.BubbleSort(strArray)
    Dim i As Integer
    For i = 1 To UBound(strArray)
        Debug.Print strArray(i)
    Next i
    JArray_BubbleSort = strArray
End Function
Public Sub JArray_GetLength()
    Dim strArray() As String
    strArray = JArray_BubbleSort
    Debug.Print JArray.GetLength(strArray)
End Sub
