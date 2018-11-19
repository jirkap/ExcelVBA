Attribute VB_Name = "JSQL"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft ActiveX Data Objects 6.1 Library
' *
' */
Option Explicit
Option Private Module
'/**
' * Connect to Excel workbook via ADODB and SQL
' * @param {String} FilePath
' * @param {String} SQLCmd – SQL command
' * @param {Boolean} FirstRowHasFieldNames – Database-like field (column) names must be placed in the first row of db range
' * @returns {ADODB.Recordset}
' * @requires Microsoft ActiveX Data Objects 6.1 Library
' * @todo Excel 12.0 for 2010/2013 files – does not work for xlsx, cannot find installable ISAM error
' *
' * @example
' *
' * Test Queries (modify as needed)
' *
' * "SELECT * FROM [" & SheetName & "$] WHERE [" & Column & "]= " & "'" & LookupValue & "'"
' * "SELECT * FROM [" & SheetName & "$] WHERE [Record Number] = '16'"   ' [Record Number] must be the column's header!
' */
Public Function ADODBRead(FilePath As String, _
                          SQLCmd As String, _
                          Optional FirstRowHasFieldNames As Boolean = True) As ADODB.Recordset
    Dim objRecordset As New ADODB.Recordset
    Dim objConnection As New ADODB.Connection
    Dim objCommand As New ADODB.Command
    Dim strFirstRowHasFieldNames As String
    If FirstRowHasFieldNames = True Then
        strFirstRowHasFieldNames = "Yes"
    Else
        strFirstRowHasFieldNames = "No"
    End If
    With objConnection
        .Provider = "Microsoft.Jet.OLEDB.4.0" ' JET is slightly faster than ACE (revert back to version 4.0 if 12.0 fails)
        .ConnectionString = "Data Source='" & FilePath & "'; Extended Properties='Excel 8.0; HDR=" & strFirstRowHasFieldNames & "; IMEX=1'"
        .Open
    End With
    Set objCommand.ActiveConnection = objConnection
    objCommand.CommandType = adCmdText
    objCommand.CommandText = SQLCmd
    objRecordset.CursorLocation = adUseClient
    objRecordset.CursorType = adOpenDynamic
    objRecordset.LockType = adLockOptimistic
    objRecordset.Open objCommand
    Set objRecordset.ActiveConnection = Nothing ' Disconnect
    If CBool(objCommand.State And adStateOpen) = True Then Set objCommand = Nothing
    If CBool(objConnection.State And adStateOpen) = True Then objConnection.Close
    Set objConnection = Nothing
    Set ADODBRead = objRecordset
End Function

