Attribute VB_Name = "JListView"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft Windows Common Controls 6.0 (SP6)
' *
' */
Option Explicit
Option Private Module
'/**
' * Populate ListView control with tabular data (ignores empty rows)
' * @param {ListView} LstView
' * @param {Range} Headers
' * @param {Range} Data
' * @param {Integer} ColWidth
' * @param {Boolean} AllowColReorder - Allow user to reorder columns
' * @param {String} ViewType - lvwReport (default) | lvwList | lvwIcon | lvwSmallIcon
' * @requires Microsoft Windows Common Controls 6.0 (SP6)
' *
' * @example
' *
' * Source table design
' * +--------------+--------------+--------------+
' * | ColumnHeader | ColumnHeader | ColumnHeader | > Headers range
' * +==============+==============+==============+
' * | ListItem     | SubItem      | SubItem      | > Data range
' * +--------------+--------------+--------------+
' * | ListItem     | SubItem      | SubItem      |
' * +--------------+--------------+--------------+
' *
' */
Public Sub InsertTabularData(ByRef LstView As ListView, _
                             ByRef Headers As Range, _
                             ByRef Data As Range, _
                             Optional ColWidth As Integer = 100, _
                             Optional ByVal AllowColReorder As Boolean = True, _
                             Optional ByVal ViewType As String = lvwReport)
     With LstView
        .View = ViewType
        .AllowColumnReorder = AllowColReorder
        Call .ColumnHeaders.Clear
        Dim varCell As Range
        Dim clhHeader As ColumnHeader
        For Each varCell In Headers
            Set clhHeader = .ColumnHeaders.Add(, , varCell) ' Add headers
            clhHeader.Width = ColWidth
        Next varCell
        Dim lviItem As ListItem
        For Each varCell In Data
            If Not varCell = vbNullString And varCell.Column = 1 Then
                Set lviItem = .ListItems.Add(, , varCell) ' Populate 1st column
                Dim i As Integer
                For i = 1 To Data.Columns.Count
                    lviItem.ListSubItems.Add , , varCell.Offset(0, i).Value ' Populate remaining columns
                Next i
            End If
        Next varCell
    End With
End Sub
'/**
' * Sort ListView by clicking on header
' * @param {ListView} LstView
' * @param {MSComctlLib.ColumnHeader} ColumnHeader - Pass from event handler
' *
' * @example
' *
' * Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' *     Call JListView.ClickHeaderToSort(ListView1, ColumnHeader)
' * End Sub
' *
' */
Public Sub ClickHeaderToSort(ByRef LstView As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LstView
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

