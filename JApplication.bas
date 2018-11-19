Attribute VB_Name = "JApplication"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Get Application language
' * @returns {Long} 1029 (CZ), 1033 (US), 2057 (UK), 1030 (DK), 1031 (DE)
' * @link https://msdn.microsoft.com/en-us/goglobal/bb964664
' *
' */
Public Function GetLanguage() As Long
    GetLanguage = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
End Function
