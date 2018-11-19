Attribute VB_Name = "JConst"
Option Explicit
Option Private Module

'/**
' * Colors
' *
' */
Public Const clrBlack As Long = 0
Public Const clrWhite As Long = 16777215

'/**
' * Characters
' * @example
' *
' * ChrW(int)
' *
' */
Public Const chrArrowDownRight As Long = 8627

'/**
' * Browse... filters
' *
' */
Public Const strFilterBrowseAllFiles As String = "All Files,*.*"
Public Const strFilterBrowseExcelPlain As String = "Excel Workbook (*.xls*), *.xls*"
Public Const strFilterBrowseExcelMacroEnabled As String = "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm"
Public Const strFilterBrowseTxt As String = "Text Files (*.txt), *.txt"
Public Const strFilterBrowsePdf As String = "PDF Files (*.pdf), *.pdf"
Public Const strFilterBrowseMsg As String = "Outlook Messages (*.msg), *.msg"
Public Const strFilterBrowseJpg As String = "Images (*.jpg; *.jpeg), *.jpg, *.jpeg"

'/**
' * Regular expressions
' *
' */
Public Const strREWordFiles As String = ".+\.([dD][oO][cC](\b|[xX]|[mM]))"  ' .doc | .docx | .docm
Public Const strREExcelFiles As String = ".+\.([xX][lL][sS](\b|[xX]|[mM]))" ' .xls | .xlsx | .xlsm

'/**
' * Alert messages
' *
' */
' (title)
Public Const strMsgTitleActionNotAllowed = "Action not allowed"
Public Const strMsgTitleInformation As String = "Information"
Public Const strMsgTitleError As String = "Error"
' (body)
Public Const strMsgBodyPleaseWait = "Please wait until the operation is finished. "
Public Const strMsgBodyReallyCancelQ As String = "Really cancel? "
Public Const strMsgBodyInvalidPath As String = "Invalid path."
Public Const strMsgBodyEmptyPath As String = "Empty path."

'/**
' * Cell data formats
' *
' */
Public Const strCellFormatDate_dMMMMyyyy As String = "d. MMMM yyyy" ' 01. leden 2016
