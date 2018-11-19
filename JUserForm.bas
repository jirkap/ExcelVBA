Attribute VB_Name = "JUserForm"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Cancel/redefine Close button behavior
' * @param {Integer} CloseMode - Event CloseMode (event that triggered QueryClose)
' * @param {String} MsgBody
' * @param {String} MsgTitle
' * @param {Boolean} AskToClose - Allow user to close by clicking OK, False by default
' * @returns {Boolean}
' * @todo The message window is missing icon if AskToClose = True
' *
' * @example
' *
' * Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' *     Cancel = JUserForm.CancelClose(CloseMode, "Body text", "Title text")
' * End Sub
' *
' */
Public Function CancelClose(CloseMode As Integer, _
                            MsgBody As String, _
                            MsgTitle As String, _
                            Optional AskToClose As Boolean = False) As Boolean
    Select Case CloseMode
        Case vbFormControlMenu
            If AskToClose = False Then
                MsgBox MsgBody, vbExclamation, MsgTitle
                CancelClose = True
            Else
                If MsgBox(MsgBody, vbOKCancel, MsgTitle) = vbCancel Then
                    CancelClose = True
                    Else
                        CancelClose = False
                    End If
                End If
        Case vbFormCode ' Unload attempt via VBA
        Case vbAppWindows ' Windows shutting down
        Case vbAppTaskManager ' App kill attempt via Task manager
    End Select
End Function

