Attribute VB_Name = "JPrint"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * @namespace
' * @property {Variant} AvailablePrinters() - Paths to all available printers
' * @property {Variant} DefaultPrinter - Path to default printer
' * @property {Integer} DefaultPrinter_ArrayPosition - Position of .DefaultPrinter in .AvailablePrinters() (e.g. for ComboBox.ListIndex)
' * @see ListAllPrinters()
' *
' */
Type Printers
    AvailablePrinters() As Variant
    DefaultPrinter As Variant
    DefaultPrinter_ArrayPosition As Integer
End Type
'/**
' * Print out specific sheets (modeless)
' * @param {String} PrinterName
' * @param {Variant} SheetsToPrint() - Array of sheets
' *
' */
Public Sub PrintSpecificSheets(PrinterName As String, ParamArray SheetsToPrint() As Variant)
    Dim i As Integer
    For i = 0 To UBound(SheetsToPrint)
        Call SheetsToPrint(i).PrintOut(ActivePrinter:=PrinterName)
    Next i
End Sub
'/**
' * Get array of all available network printers, default printer and its position in the array
' * @returns {Printers}
' *
' */
Public Function ListAllPrinters() As Printers
    Dim Printers As Printers
    ' .DefaultPrinter
    Dim strActive As String: strActive = Application.ActivePrinter
    Select Case JApplication.GetLanguage
        Case 1029
            strActive = Trim(JString.TrimAfter(strActive, " na Ne")) ' Remove " na Ne##" for CS
        Case 1033, 2057
            strActive = Trim(JString.TrimAfter(strActive, " on Ne")) ' Remove " on Ne##" for US, UK
        Case Else
            strActive = "Lang Error: Update ListAllPrinters() first."
    End Select
    Printers.DefaultPrinter = strActive
    ' .AvailablePrinters(), .DefaultPrinter_ArrayPosition
    Dim objConnPrinters As Object: Set objConnPrinters = CreateObject("WScript.Network").EnumPrinterConnections ' Printer port/name id collection
    ReDim Printers.AvailablePrinters(0 To objConnPrinters.Count \ 2 - 1) ' Set size of the printer names array
    Dim i As Integer
    For i = 0 To UBound(Printers.AvailablePrinters)
        Printers.AvailablePrinters(i) = objConnPrinters.Item(i * 2 + 1)
        If Printers.AvailablePrinters(i) = strActive Then Printers.DefaultPrinter_ArrayPosition = i
    Next i
    Set objConnPrinters = Nothing
    ListAllPrinters = Printers
End Function
