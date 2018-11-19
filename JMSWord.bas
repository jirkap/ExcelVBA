Attribute VB_Name = "JMSWord"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Option Explicit
Option Private Module
'/**
' * @namespace
' * @property {Word.Application} WordApp
' * @property {Word.Document} WordDoc
' * @requires Microsoft Word 14.0 Object Library
' * @see LaunchAndOpenDocument()
' *
' */
Type WordDocument
    WordApp As Word.Application
    WordDoc As Word.Document
End Type
'/**
' * Launch Word, open a document and return references to both
' * @param {String} Path - Path to document
' * @param {Boolean} Visible - Word application visibility
' * @returns {WordDocument}
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Function LaunchAndOpenDocument(Path As String, Visible As Boolean) As WordDocument
    On Error GoTo ErrHandler
    Dim Word As WordDocument
    Set Word.WordApp = New Word.Application
    Word.WordApp.Visible = Visible
    Set Word.WordDoc = Word.WordApp.Documents.Add(Path)
    LaunchAndOpenDocument = Word

ErrHandler:
    Set Word.WordApp = Nothing
    Set Word.WordDoc = Nothing
End Function
'/**
' * @param {Boolean} Visible - Word application visibility
' * @returns {Word.Application}
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Function Launch(Optional Visible = True) As Word.Application
    On Error GoTo ErrHandler
    Dim WordApp As Word.Application: Set WordApp = New Word.Application
    WordApp.Visible = Visible
    Set Launch = WordApp
    
ErrHandler:
    Set WordApp = Nothing
End Function
'/**
' * Open document in already running instance of Word application
' * @param {Word.Application} WordApp
' * @param {String} Path – Path to a Word document
' * @param {Boolean} ROnly – Open as ReadOnly
' * @returns {Word.Document}
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Function OpenDocument(WordApp As Word.Application, _
                             Path As String, _
                             Optional ROnly = True) As Word.Document
    'On Error GoTo ErrHandler
    Set OpenDocument = WordApp.Documents.Open(Path, , ROnly, False)
'ErrHandler:
 '   Set OpenDocument = Nothing
End Function
'/**
' * Fill Word.Document.FormFields by iterating them and passing array values one by one
' * @param {Word.Document} WordDoc – Document reference
' * @param {String} Values() - Array of values ordered by FormFields hierarchy in target document (top to bottom)
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Sub FillFormFields(WordDoc As Word.Document, Values() As String)
    On Error GoTo ErrHandler
    Dim ff As Word.FormField
    Dim i As Integer
    For Each ff In WordDoc.FormFields
        ff.Result = Values(i)
        i = i + 1
    Next ff
    
ErrHandler:
    Set WordDoc = Nothing
End Sub
'/**
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Sub FillContentControls()
    
End Sub
'/**
' * Print out a Word document on default printer (modeless)
' * @param {Word.Document} WordDoc
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Sub PrintDocument(WordDoc As Word.Document)
    WordDoc.PrintOut
End Sub
'/**
' * @param {Word.Document} WordDoc
' * @param {Word.Application} WordApp
' * @param {Boolean} SaveChanges
' * @param {Boolean} QuitWordApp
' * @requires Microsoft Word 14.0 Object Library
' *
' */
Public Sub CloseDocument(WordDoc As Word.Document, _
                         WordApp As Word.Application, _
                         Optional SaveChanges As Boolean = False, _
                         Optional QuitWordApp As Boolean = True)
    Call WordDoc.Close(SaveChanges)
    If Not QuitWordApp = False Then Call WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
End Sub
