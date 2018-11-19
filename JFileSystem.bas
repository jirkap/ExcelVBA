Attribute VB_Name = "JFileSystem"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft VBScript Regular Expressions 5.5
' * @requires Microsoft Scripting Runtime
' *
' */
Option Explicit
Option Private Module
'/**
' * Browse for file and return its path
' * @param {String} Filter - Filetype filter, show all files by default
' * @param {String} Title - Window title
' * @param {String} ValueWhenCancelled
' * @param {Boolean} MultipleSelect - False by default
' * @returns {Variant} File path or False when cancelled
' * @todo Add StartDir param
' *
' */
Public Function BrowseForFile(Optional Filter As String = JConst.strFilterBrowseAllFiles, _
                              Optional Title As String = "Browse for file...", _
                              Optional ValueWhenCancelled As Variant, _
                              Optional MultipleSelect As Boolean = False) As Variant
    Dim varFile As Variant: varFile = Application.GetOpenFilename(Filter, , Title, MultipleSelect)
    If LCase(varFile) = False Then ' Hit Cancel
        If Not IsMissing(ValueWhenCancelled) Then
            BrowseForFile = ValueWhenCancelled
        Else
            BrowseForFile = False
        End If
    Else
        BrowseForFile = varFile
    End If
End Function
'/**
' * Browse for folder and return its path
' * @param {Variant} StartDir - Directory to start browsing from
' * @returns {Variant} Selected folder path
' * @link http://www.ozgrid.com/forum/showthread.php?t=80008
' *
' */
Public Function BrowseForFolder(Optional StartDir As Variant = -1) As Variant
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Dim varPath As Variant
    With fd
        .Title = "Browse for folder..."
        .AllowMultiSelect = False
        If StartDir = -1 Then
            .InitialFileName = Application.DefaultFilePath
        Else
            If Right(StartDir, 1) <> "\" Then
                .InitialFileName = StartDir & "\"
            Else
                .InitialFileName = StartDir
            End If
        End If
        If .Show <> -1 Then GoTo ReturnFolder
        varPath = .SelectedItems(1)
    End With

ReturnFolder:
    BrowseForFolder = varPath
    Set fd = Nothing
End Function
'/**
' * Find matching files
' * @param {String} DirPath - Path to directory
' * @param {String} RegExp - Regular expression
' * @param {Boolean} SearchSubFolders - False by default
' * @param {Boolean} CaseSensitive - False by default
' * @returns {String()} Array of filepaths
' * @requires RecursiveFileSearch()
' * @requires Microsoft VBScript Regular Expressions 5.5
' * @link http://stackoverflow.com/questions/555776/list-files-of-certain-pattern-using-excel-vba
' * @todo FileSearch(0) is always an empty string, may be because arrays start at 0 and collections at 1?
' *
' * @example
' *
' * ' One item will be stored in TemplatePath(0)
' * Dim TemplatePath() As String: TemplatePath = FileSearch("C:\Folder", "QEF123")
' */
Public Function FileSearch(ByVal DirPath As String, _
                           ByVal RegExp As String, _
                           Optional ByVal SearchSubFolders As Boolean = False, _
                           Optional ByVal CaseSensitive As Boolean = False) As String()
    Dim objFs As Object: Set objFs = CreateObject("Scripting.FileSystemObject")
    Dim objRegExp As Object: Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = RegExp
    objRegExp.IgnoreCase = CaseSensitive
    Dim colFiles As Collection: Set colFiles = New Collection
    Call RecursiveFileSearch(DirPath, objRegExp, colFiles, objFs, SearchSubFolders)
    Dim arrFound() As String
    ReDim arrFound(colFiles.Count)
    Dim varFile As Variant, i As Integer
    For Each varFile In colFiles
        arrFound(i) = varFile
        i = i + 1
    Next
    Set objFs = Nothing
    Set objRegExp = Nothing
    FileSearch = arrFound
End Function
'/**
' * Recursive function that populates Matches Collection using ByRef
' * @private
' * @param {String} DirPath - Path to directory
' * @param {Object} RegExp
' * @param {Collection} Matches
' * @param {Object} FSObj
' * @param {Boolean} SearchSubFolders - False by default
' * @returns {Collection}
' *
' */
Private Function RecursiveFileSearch(ByVal DirPath As String, _
                                     ByRef RegExp As Object, _
                                     ByRef Matches As Collection, _
                                     ByRef FSObj As Object, _
                                     Optional SearchSubFolders As Boolean = False)
    Dim objFolder As Object: Set objFolder = FSObj.GetFolder(DirPath)
    Dim objFile As Object
    For Each objFile In objFolder.Files
        If RegExp.Test(objFile) Then
            Call Matches.Add(objFile)
        End If
    Next
    If SearchSubFolders = True Then
        Dim objSubFolders As Object: Set objSubFolders = objFolder.objSubFolders
        Dim varSubFolder As Variant
        For Each varSubFolder In objSubFolders ' Loop through each of the subfolders recursively
            Call RecursiveFileSearch(varSubFolder, RegExp, Matches, FSObj)
        Next
        Set objSubFolders = Nothing
    End If
    Set objFolder = Nothing
    Set objFile = Nothing
End Function
'/**
' * @param {String} DirPath - Trailing slash is not required
' * @returns {Boolean}
' *
' */
Public Function DirExists(DirPath As String) As Boolean
    On Error Resume Next
    DirExists = (GetAttr(DirPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function
'/**
' * @param {String} Path
' * @returns {String}
' * @requires Microsoft Scripting Runtime
' * @link http://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path
' *
' */
Public Function GetFilenameFromPath(ByVal Path As String) As String
    Dim objFs As New FileSystemObject
    GetFilenameFromPath = objFs.GetFileName(Path)
End Function
' /**
' * Retrieve all files in directory and return them as Collection
' * @param {String} DirPath
' * @param {String} Wildcard – Accepts only ? or *.
' * @param {VbFileAttribute} FileAttribute
' * @returns {Collection}
' *
' */
Public Function GetFilesInDirectory(DirPath As String, _
                                    Optional Wildcard As String = "", _
                                    Optional FileAttribute As VbFileAttribute = vbNormal) As Collection
    Dim varFile As Variant, col As New Collection
    DirPath = IIf(Right$(DirPath, 1) <> "\", DirPath & "\", DirPath) ' Add "\" to DirPath if missing
    varFile = Dir(DirPath & Wildcard, FileAttribute)
    While (Not varFile = vbNullString)
        Call col.Add(DirPath & varFile, varFile)
        varFile = Dir
    Wend
    Set GetFilesInDirectory = col
End Function
'/**
' * Return last updated file in folder
' * @param {String} DirPath - Trailing slash is not required
' * @param {String} Filter - "*.ext", * and ? wildcards allowed (e.g. "*.xl*" for all Excel files)
' * @returns {Variant} "Filename.ext" | 0 (Dir is empty) | -1 (Dir does not exist)
' * @requires DirExists()
' *
' */
Public Function GetNewestFile(ByVal DirPath As String, _
                              Optional ByVal Filter As String = "*.*") As Variant
    Dim strFile, strNewestFile As String, datNewestDate As Date
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
    If DirExists(DirPath) = False Then
        Call MsgBox("Directory " & DirPath & " does not exist.", vbOKOnly + vbCritical, JConst.strMsgTitleError)
        GetNewestFile = -1 ' Directory does not exist
    Else
        strFile = Dir(DirPath & Filter, vbNormal)
        If strFile <> "" Then
            strNewestFile = strFile
            datNewestDate = FileDateTime(DirPath & strFile)
            Do While strFile <> ""
                If FileDateTime(DirPath & strFile) > datNewestDate Then
                     strNewestFile = strFile
                     datNewestDate = FileDateTime(DirPath & strFile)
                 End If
                 strFile = Dir
            Loop
            GetNewestFile = strNewestFile ' Return filename
        Else
            Call MsgBox("Directory " & DirPath & " does not contain any files matching extension" & Filter, vbOKOnly + vbInformation, JConst.strMsgTitleInformation)
            GetNewestFile = 0 ' Directory is empty
        End If
    End If
End Function
'/**
' * Open file/hyperlink with default Application
' * @param {String} Path
' * @link https://msdn.microsoft.com/en-us/library/d5fk67ky(v=vs.84).aspx
' * @todo Opens user's default directory without any error if the Path is invalid
' *
' */
Public Sub OpenWithDefaultApplication(ByVal Path As String)
    If Not Path = vbNullString Then
        Dim wshShell As Object: Set wshShell = CreateObject("WScript.Shell")
        Call wshShell.Run("explorer.exe " & Path)
        Set wshShell = Nothing
    Else
       Call MsgBox(JConst.strMsgBodyEmptyPath, vbOKOnly + vbInformation, JConst.strMsgTitleInformation)
       Exit Sub
    End If
End Sub

