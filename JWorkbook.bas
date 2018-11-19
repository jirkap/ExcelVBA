Attribute VB_Name = "JWorkbook"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module

'/**
' * @namespace
' * @see GetMetaData()
' *
' */
Type MetaData
    Title As String
    Subject As String
    Author As String
    Keywords As String
    Comments As String
    Template As String
    LastAuthor As String
    RevisionNumber As Integer
    ApplicationName As String
    LastPrintDate As Date
    LastSaved As Date ' (!) Throws Automation Error in case the workbook has never been saved
    CreationDate As Date
    LastSaveTime As Date
    ' TotalEditingTime As Date
    ' NumberOfPages As Integer
    ' NumberOfWords As Integer
    ' NumberOfCharacters As Integer
    ' NumberOfCharactersWithSpaces As Integer
    Security As String
    Category As String
    Format As String
    Manager As String
    Company As String
    ' NumberOfBytes As Long
    ' NumberOfLines As Integer
    ' NumberOfParagraphs As Integer
    ' NumberOfSlides As Integer
    ' NumberOfNotes As Integer
    ' NumberOfHiddenSlides As Integer
    ' HyperlinkBase As String
End Type
'/**
' * Get Workbook metadata
' * @param {Workbook} Workbook – Use ActiveWorkbook if not provided
' * @returns {Metadata} See Metadata for more details
' *
' */
Public Function GetMetadata(Optional Workbook As Workbook) As MetaData
    If Not IsMissing(Workbook) Then Set Workbook = ActiveWorkbook
    Dim MetaData As MetaData
    With Workbook
        MetaData.Title = .BuiltinDocumentProperties("Title")
        MetaData.Subject = .BuiltinDocumentProperties("Subject")
        MetaData.Author = .BuiltinDocumentProperties("Author")
        MetaData.Keywords = .BuiltinDocumentProperties("Keywords")
        MetaData.Comments = .BuiltinDocumentProperties("Comments")
        MetaData.Template = .BuiltinDocumentProperties("Template")
        MetaData.LastAuthor = .BuiltinDocumentProperties("Last Author")
        MetaData.RevisionNumber = .BuiltinDocumentProperties("Revision Number")
        MetaData.ApplicationName = .BuiltinDocumentProperties("Application Name")
        MetaData.LastPrintDate = .BuiltinDocumentProperties("Last Print Date")
        MetaData.LastSaved = .BuiltinDocumentProperties("Last Save Time")
        MetaData.CreationDate = .BuiltinDocumentProperties("Creation Date")
        MetaData.LastSaveTime = .BuiltinDocumentProperties("Last Save Time")
        ' Metadata.TotalEditingTime = .BuiltinDocumentProperties("Total Editing Time") ' Exists in Excel but returns 0:00:00
        ' Metadata.NumberOfPages = .BuiltinDocumentProperties("Number of Pages") ' Word
        ' Metadata.NumberOfWords = .BuiltinDocumentProperties("Number of Words") ' Word
        ' Metadata.NumberOfCharacters = .BuiltinDocumentProperties("Number of Characters") ' Word
        ' Metadata.NumberOfCharactersWithSpaces = .BuiltinDocumentProperties("Number of Characters (with spaces)") ' Word
        MetaData.Security = .BuiltinDocumentProperties("Security")
        MetaData.Category = .BuiltinDocumentProperties("Category")
        MetaData.Format = .BuiltinDocumentProperties("Format")
        MetaData.Manager = .BuiltinDocumentProperties("Manager")
        MetaData.Company = .BuiltinDocumentProperties("Company")
        ' Metadata.NumberOfBytes = .BuiltinDocumentProperties("Number of Bytes") ' ?
        ' Metadata.NumberOfLines = .BuiltinDocumentProperties("Number of Lines") ' Word?
        ' Metadata.NumberOfParagraphs = .BuiltinDocumentProperties("Number of Paragraphs") ' Word?
        ' Metadata.NumberOfSlides = .BuiltinDocumentProperties("Number of Slides") ' PowerPoint
        ' Metadata.NumberOfNotes = .BuiltinDocumentProperties("Number of Notes") ' ?
        ' Metadata.NumberOfHiddenSlides = .BuiltinDocumentProperties("Number of Hidden Slides") ' PowerPoint
        ' Metadata.HyperlinkBase = .BuiltinDocumentProperties("Hyperlink Base") ' ?
    End With
    GetMetadata = MetaData
End Function
'/**
' * Copy range from many Workbooks and paste it into one
' * @param {String} DirPath - Path to directory containing source workbooks
' * @param {Range} SrcRange
' * @param {Workbook} TargetWorkbook
' * @param {Range} TargetRange
' * @todo TargetWorkbook must be open, replace .Activate()
' *
' */
Public Sub CopyFromManyWorkbooksToOne(DirPath As String, SrcRange As Range, TargetWorkbook As Workbook, TargetRange As Range)
    Dim i As Integer: i = 1
    Dim strFile As String: strFile = Dir(DirPath)
    Do Until strFile = ""
        Call Workbooks.Open(DirPath & strFile)  ' Open next workbook
        Call Workbooks(TargetWorkbook.Name).Activate
        TargetWorkbook.TargetRange.Value = Workbooks(strFile).SrcRange
        Call Workbooks(strFile).Close(False)
        i = i + 1
        strFile = Dir()
    Loop
End Sub
'/**
' * @param {String} Path
' * @param {Boolean} Visible - Allows to open file in hidden window
' * @param {Boolean} Activate - If True, all objects/methods become available without the need to reference the Workbook object
' * @param {Boolean} UpdateLinks
' * @param {Boolean} ReadOnly
' * @returns {Workbook}
' * @todo Fix error caused by opening already opened file not in ReadOnly mode
' *
' */
Public Function OpenWorkbook(ByVal Path As String, _
                             Optional ByVal Visible As Boolean = True, _
                             Optional ByVal Activate As Boolean = True, _
                             Optional ByVal UpdateLinks As Boolean = True, _
                             Optional ByVal ReadOnly As Boolean = True) As Workbook
    Dim Wbk As Workbook: Set Wbk = Workbooks.Open(Path, UpdateLinks, ReadOnly)
    Dim strName As String: strName = JFileSystem.GetFilenameFromPath(Path)
    Windows(strName).Visible = Visible
    If Activate = True Then
        Call Wbk.Activate
    End If
    Set OpenWorkbook = Wbk
End Function
'/**
' * @param {String} Filename - Suggested filename
' * @param {String} Filter - Filetype filter, show all by default
' * @param {String} Title - Dialog window title
' * @returns {Variant}
' *
' */
Public Function SaveAs(Optional Filename As String = "", _
                       Optional Filter As String = JConst.strFilterBrowseAllFiles, _
                       Optional Title As String = "Save As") As Variant
    Dim varFile As Variant: varFile = Application.GetSaveAsFilename(Filename, Filter, , Title)
    If LCase(varFile) = "False" Then ' User hit Cancel
        SaveAs = False
    Else
        SaveAs = varFile
    End If
End Function

