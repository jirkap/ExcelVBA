Attribute VB_Name = "JDate"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft Scripting Runtime
' *
' */
Option Explicit
Option Private Module
'/**
' * @returns {Date}
' *
' */
Public Function GetLastDateInMonth(ByVal InputDate As Date) As Date
    Dim y, m, dte As Date
    y = Year(InputDate)
    m = Month(InputDate)
    dte = DateSerial(y, m + 1, 0)
    GetLastDateInMonth = dte
End Function
