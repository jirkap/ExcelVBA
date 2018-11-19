Attribute VB_Name = "JError"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Custom console error report, use with ErrHandler
' * @todo Saving to txt
' *
' */
Public Sub ReportToConsole()
    Debug.Print vbNewLine
    Debug.Print "ERROR REPORT"
    Debug.Print "   Description    : " & Err.Number & " – " & Err.Description
    Debug.Print "   Source         : " & Err.Source
    Debug.Print "   Line           : " & Erl ' For Erl to work, code lines must be manually numbered, otherwise returns 0
    Debug.Print "   Module         : " & Application.VBE.ActiveCodePane.CodeModule.Name
    Debug.Print "   User           : " & Environ("UserName")
    Debug.Print "   Timestamp      : " & Now
    Debug.Print ""
    Err.Clear
End Sub
