Attribute VB_Name = "JListObject"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Find in ListObject
' * @param {ListObject} Table
' * @param {String} Query
' * @param {Boolean} CaseSensitive
' * @param {Boolean} FullMatch
' * @param {Boolean} SearchByColumns
' * @returns {Collection} Collection of Range.Address
' *
' */
Public Function FindAll(ByRef Table As ListObject, _
                        ByVal Query As String, _
                        Optional ByRef CaseSensitive As Boolean = False, _
                        Optional ByRef FullMatch As Boolean = False, _
                        Optional ByRef SearchByColumns As Boolean = False) As Collection
    Dim lngSearchOrder As Long: lngSearchOrder = xlByRows
    If SearchByColumns = True Then lngSearchOrder = xlByColumns
    Dim lngLookAt As Long: lngLookAt = xlPart
    If FullMatch = True Then lngLookAt = xlWhole
    Dim rngAfter As Range: Set rngAfter = Table.DataBodyRange(1, 1)
    Dim rngFind As Range: Set rngFind = Table.DataBodyRange.Find(Query, rngAfter, , lngLookAt, lngSearchOrder)
    Dim colStore As New Collection
    If Not rngFind Is Nothing Then
        Dim blnStop As Boolean
        Call colStore.Add(rngFind.Address) ' Save 1st occurence
        Do
            Set rngFind = Table.DataBodyRange.FindNext(rngFind)
            If rngFind.Address = colStore(1) Then Exit Do ' .FindNext loops endlessly
            Call colStore.Add(rngFind.Address) ' Save all remaining occurences
        Loop Until rngFind.Address = colStore(1)
    End If
    Set FindAll = colStore
End Function
'/**
' * Get range occupied by ListObject
' * @param {Worksheet} Sheet - Worksheet object containing Excel Table
' * @param {String} ListObj
' * @returns {String} Range.Address
' *
' */
Public Function GetRange(Sheet As Worksheet, ListObj As String) As String
    GetRange = Sheet.ListObjects(ListObj).DataBodyRange.Address
End Function
'/**
' * Vertically expand a ListObject. ListObjects consist of .HeadRowRange (if enabled) and one line in .DataBodyRange (always)
' * Table's placement coordinates cannot be changed via .Resize() method
' * @param {ListObject} ListObj
' * @param {Integer} Rows
' *
' */
Public Sub AddRows(ListObj As ListObject, Rows As Integer)
    Dim lngExistingCols As Long: lngExistingCols = ListObj.Range.Columns.Count
    With ListObj
        Call .Resize(Range(.Range.Cells(1, 1), .Range.Cells(Rows, lngExistingCols)))
    End With
End Sub
'/**
' * Find String in ListObject and hide rows with no match
' * @param {ListObject} ListObj
' * @param {String} Query
' * @param {Boolean} CaseSensitive
' * @param {Boolean} FullMatch
' * @requires FindAll()
' *
' */
Public Sub FulltextFilter(ByRef ListObj As ListObject, _
                          ByVal Query As String, _
                          Optional ByRef CaseSensitive As Boolean = False, _
                          Optional ByRef FullMatch As Boolean = False)
    Application.ScreenUpdating = False
    If Not Len(Query) = 0 Then ' Covers both vbNullString and ""
        Dim colMatches As New Collection
        Set colMatches = FindAll(ListObj, Query, CaseSensitive, FullMatch)
        Dim rngRow As Range
        Dim varKeep As Variant
        Dim blnHide As Boolean
        For Each rngRow In ListObj.DataBodyRange.Rows
            blnHide = True
            For Each varKeep In colMatches
                If Range(varKeep).Row = rngRow.Row Then
                    blnHide = False
                    Exit For
                End If
            Next varKeep
            If blnHide = True Then
                rngRow.EntireRow.Hidden = True
            Else
                rngRow.EntireRow.Hidden = False
            End If
        Next rngRow
        Application.ScreenUpdating = True
    End If
End Sub
'/**
' * Unhide all rows in ListObject
' * @param {ListObject} ListObj
' *
' */
Public Sub UnhideAllRows(ByRef ListObj As ListObject)
    Application.ScreenUpdating = False
    Dim rngRow As Variant
    For Each rngRow In ListObj.DataBodyRange.Rows
        rngRow.EntireRow.Hidden = False
    Next rngRow
    Application.ScreenUpdating = True
End Sub
'/**
' * @param {ListObject} ListObj
' *
' */
Public Sub TriggerAutoFilter(ListObj As ListObject)
    Call ListObj.Range.AutoFilter
End Sub
'/**
' * Resize ListObject to Range and paste its contents
' * @param {ListObject} ListObj
' * @param {Range} SrcRange
' * @param {Integer} Rows
' * @param {Integer} Columns
' * @param {Boolean} ContainsHeaders
' *
' */
Public Sub ResizeToAndPasteRange(ByRef ListObj As ListObject, _
                                 ByRef SrcRange As Range, _
                                 Optional ByVal Rows As Integer = 0, _
                                 Optional ByVal Columns As Integer = 0, _
                                 Optional ByVal ContainsHeaders As Boolean = True)
    With ListObj.DataBodyRange
        Call .ClearContents
        If ContainsHeaders = True Then
            If Not IsMissing(Rows) Then Rows = Rows + 1
        End If
        Call ListObj.Resize(ListObj.Range.Resize(Rows, Columns))
        .Value = SrcRange.Value
    End With
End Sub
