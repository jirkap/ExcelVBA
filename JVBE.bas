Attribute VB_Name = "JVBE"
'/**
' * @author pokorny.jirka@gmail.com
' * @requires Microsoft Visual Basic for Application Extensibility Library 5.3
' * @todo Standard Class modules (e.g. ThisWorkbook) can be exported too as *.cls files
' *
' */
Option Explicit
Option Private Module
'/**
' * Export all VBA project modules
' * @param {String} DestinationDir – Trailing slash is not required
' * @requires Microsoft Visual Basic for Application Extensibility Library 5.3
' * @requires GetComponentExtension()
' * @link https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
' *
' */
Public Sub ExportAllModules(DestinationDir As String)
    Dim cmp As VBComponent
    ' Dim strVersion As String: strVersion = CStr(Format(Now, "_yyyy-mm-dd_hh-mm-ss"))
    If Right(DestinationDir, 1) <> "\" Then DestinationDir = DestinationDir & "\"
    Debug.Print vbNewLine
    Debug.Print "Exporting VBA source files"
    Debug.Print "   Destination    : " & DestinationDir
    Debug.Print "   Files          : " & Application.VBE.ActiveVBProject.VBComponents.Count
    For Each cmp In Application.VBE.ActiveVBProject.VBComponents
        If cmp.Type = vbext_ct_ClassModule Or cmp.Type = vbext_ct_StdModule Then
            Debug.Print "   File    : " & cmp.Name & GetComponentExtension(cmp.Type) ' & strVersion &
            Call cmp.Export(DestinationDir & cmp.Name & GetComponentExtension(cmp.Type)) ' & strVersion &
        End If
    Next
End Sub
'/**
' * @private
' * @param {vbext_ComponentType} VBEComponentType
' * @requires Microsoft Visual Basic for Application Extensibility Library 5.3
' * @link https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
' *
' */
Private Function GetComponentExtension(VBEComponentType As vbext_ComponentType) As String
    Select Case VBEComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            GetComponentExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            GetComponentExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            GetComponentExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            GetComponentExtension = vbNullString
    End Select
End Function
'/**
' * Remove all modules except JVBE
' * @private
' * @requires Microsoft Visual Basic for Application Extensibility Library 5.3
' * @link https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
' *
' */
Private Sub RemoveAllModules()
    Dim project As VBProject, cmp As VBComponent
    Set project = Application.VBE.ActiveVBProject
    For Each cmp In project.VBComponents
        If Not cmp.Name = "JVBE" And (cmp.Type = vbext_ct_ClassModule Or cmp.Type = vbext_ct_StdModule) Then
            Call project.VBComponents.Remove(cmp)
        End If
    Next
End Sub



