Attribute VB_Name = "JComboBox"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires
' *
' */
Option Explicit
Option Private Module
'/**
' * Populate ComboBox with Array elements (ignores CmbBox.Locked = True)
' * @param {MSForms.ComboBox} CmbBox - Sheet.cmbName
' * @param {Variant} Arr - Array of values
' * @param {Integer} DefaultListIndex – Empty by default
' *
' * Fixed: Replaced DefaultValue with DefaultListIndex (.Value cannot be set when ComboBox.Style = fmStyleDropDownList)
' * Fixed: ComboBox to MSForms.ComboBox, .Clear
' *
' */
Public Sub PopulateWithArray(ByVal CmbBox As MSForms.ComboBox, _
                             ByVal Arr As Variant, _
                             Optional DefaultListIndex As Integer = -1)
    With CmbBox
        Call .Clear
        .List = Arr
        .ListIndex = DefaultListIndex
    End With
End Sub
'/**
' * Populate ComboBox with unique values in Range
' * @param {Control} CmbBox
' * @param {Range} SrcRange
' * @param {Integer} SpecificColumnOnly - Add only values from specific column in range
' *
' */
Public Sub PopulateWithUniquesInRange(ByRef CmbBox As Control, _
                                      ByVal SrcRange As Range, _
                                      Optional ByRef SpecificColumnOnly As Integer)
    Dim rngCell As Range
    Dim colTemp As New Collection
    On Error Resume Next
    For Each rngCell In SrcRange
        If Not IsMissing(SpecificColumnOnly) Then
            If rngCell.Column = SpecificColumnOnly Then
                Call colTemp.Add(rngCell.Value, rngCell.Value) ' Key makes it unique!
            End If
        Else
            Call colTemp.Add(rngCell.Value, rngCell.Value)
        End If
    Next rngCell
    On Error GoTo 0
    Dim varItem As Variant
    For Each varItem In colTemp
        Call CmbBox.AddItem(varItem)
    Next varItem
End Sub
'/**
' * Find Value in ComboBox.List and return its ListIndex
' * @param {ComboBox} CmbBox - List1.cmbName
' * @param {Variant} Value - Value to find
' * @returns {Integer}
' *
' */
Public Function GetListIndex(CmbBox As ComboBox, Value As Variant) As Integer
    Dim i, intIndex As Integer
    For i = 0 To CmbBox.ListCount - 1
        If Value = CmbBox.List(i) Then intIndex = i
    Next i
    GetListIndex = intIndex
End Function
