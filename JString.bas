Attribute VB_Name = "JString"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Trim string after substring
' * @param {String} Str - String to modify
' * @param {String} Substr
' * @returns {String}
' *
' */
Public Function TrimAfter(Str As String, Substr As String) As String
    TrimAfter = Left(Str, InStr(Str, Substr))
End Function
