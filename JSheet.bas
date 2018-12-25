Attribute VB_Name = "JSheet"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Embed file in document (same as Insert > Object)
' * @param {Worksheet} Sheet
' * @param {Double} Left
' * @param {Double} Top
' * @param {Double} Width
' * @param {Double} Height
' * @param {String} Filter - All by default
' * @param {Boolean} Linked - False by default
' * @param {Boolean} DisplayAsIcon - True by default
' * @param {String} IconLabel
' * @param {String} ShapeName
' * @todo Extend to other filetypes
' *
' */
Public Sub EmbedFile(Sheet As Worksheet, _
                     Left As Double, Top As Double, _
                     Width As Double, Height As Double, _
                     Optional Filter As String = JConst.strFilterBrowseAllFiles, _
                     Optional Linked As Boolean = False, _
                     Optional DisplayAsIcon As Boolean = True, _
                     Optional IconLabel As String, _
                     Optional ShapeName As String)
    Dim varFile As Variant: varFile = JFn.Excel_BrowseForFile(Filter, "Embed file to worksheet '" & Sheet.Name & "'") ' Open Browse dialog
    If varFile = False Then Exit Sub ' User hit Cancel
    Dim strIconFile As String
    Dim intIconIndex As Integer
    Select Case Filter
        Case JConst.strFilterBrowseExcelPlain: ' Excel
            strIconFile = "C:\Windows\Installer\{90140000-0012-0000-0000-0000000FF1CE}\xlicons.exe"
            intIconIndex = 0
        Case JConst.strFilterBrowsePdf: ' PDF
            strIconFile = "C:\Windows\Installer\{AC76BA86-7AD7-1033-7B44-AB0000000001}\PDFFile_8.ico"
            intIconIndex = 0
    End Select
    If IconLabel = Empty Then IconLabel = JFileSystem.GetFilenameFromPath(varFile) ' Label by filename if not specified
    Dim shpEmbeddedFile As Shape ' varFile icon
    Set shpEmbeddedFile = Sheet.Shapes.AddOLEObject(, varFile, Linked, DisplayAsIcon, strIconFile, intIconIndex, IconLabel, Left, Top, Width, Height)
    With shpEmbeddedFile ' Styling
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .Locked = False
        If ShapeName <> Empty Then .Name = ShapeName ' usually Id
    End With
End Sub
Public Sub RemoveSelectedRows()
    Call Selection.EntireRow.Delete
End Sub
Public Sub RemoveFirstSelectedRow()
    Call Rows(ActiveCell.row).EntireRow.Delete
End Sub
'/**
' * @param {Range} FromRange
' * @param {Long} NumberOfNewRows
' * @param {String} RunForEach - A Sub or Function to run after each new row is created (as Arg1)
' * @link http://www.mrexcel.com/forum/excel-questions/931146-Application-run-each-iteration.html
' *
' */
Public Sub AddRowsBelow(FromRange As Range, NumberOfNewRows As Long, Optional RunForEach As String)
    Dim intNewRowNr As Integer: intNewRowNr = FromRange.row
    Dim i As Integer
    For i = 1 To NumberOfNewRows
        Call FromRange.offset(1).EntireRow.Insert(xlDown)
        If Not IsMissing(RunForEach) Then
            intNewRowNr = intNewRowNr + 1
            Call Application.Run(RunForEach, intNewRowNr)
        End If
    Next
End Sub
'/**
' * @param {Range} FromRange
' * @param {Long} NumberOfNewRows
' *
' */
Public Sub AddRowsAbove(FromRange As Range, NumberOfNewRows As Long)
    Call FromRange.Resize(NumberOfNewRows).EntireRow.Insert
End Sub
'/**
' * Get last filled row in given column
' * @param {Worksheet} Sheet
' * @param {Integer} Column
' * @returns {Integer}
' *
' */
Public Function GetLastRow(Sheet As Worksheet, Column As Integer) As Integer
    GetLastRow = Sheet.Cells(Sheet.Rows.Count, Column).End(xlUp).row
End Function
'/**
' * Get last filled column in given row
' * @param {Worksheet} Sheet
' * @param {Integer} Row
' * @returns {Integer}
' *
' */
Public Function GetLastColumn(Sheet As Worksheet, row As Integer) As Integer
    GetLastColumn = Sheet.Cells(row, Sheet.Columns.Count).End(xlToLeft).Column
End Function
'/**
' * Loop all sheets and change their visibility
' * @param {Workbook} Wbook
' * @param {Integer} SheetToSkip
' *
' */
Public Sub AllSheetsVisibility(Visibility As XlSheetVisibility, Optional SheetToSkip As Worksheet, Optional Wbook As Workbook)
    If Wbook Is Nothing Then
        Set Wbook = ThisWorkbook
    End If
    Dim w As Worksheet
    For Each w In Wbook.Worksheets
        If Not w Is SheetToSkip Or SheetToSkip Is Nothing Then
            w.Visible = Visibility
        End If
    Next w
End Sub
