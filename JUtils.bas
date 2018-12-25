Attribute VB_Name = "JUtils"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft Scripting Runtime
' *
' */
Option Explicit
Option Private Module
'/**
' * @returns {String}
' *
' */
Public Function GetUsername() As String
    GetUsername = Environ("UserName")
End Function
'/**
' * Returns name of current folder
' * @returns {String}
' *
' */
Public Function GetWorkingFolderName(Optional Wbook As Workbook) As String
    If Wbook Is Nothing Then
        Set Wbook = ThisWorkbook
    End If
    Dim p As String: p = Wbook.path
    GetWorkingFolderName = Right(p, Len(p) - InStrRev(p, "\"))
End Function
'/**
' * Custom error handling
' *
' */
Public Sub ErrHandler(ErrObj As ErrObject, Optional RoutineName As String)
    ' Dim routine As String
    ' routine = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Application.VBE.ActiveCodePane.TopLine, 0) ' This works only when VBE has been activated
    Dim msg As String
    msg = "There has been an error"
    If RoutineName <> vbNullString Then
        msg = msg & " in " & RoutineName
    End If
    msg = msg & ":" & vbNewLine & ErrObj.Description
    
    MsgBox msg, vbExclamation + vbOKOnly, "Error: " & CStr(Err.Number)
    ErrObj.Clear
End Sub
    


