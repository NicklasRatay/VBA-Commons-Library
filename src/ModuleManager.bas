Attribute VB_Name = "ModuleManager"
Option Explicit

Private Sub ExportAll()
' Exports all classes, forms and modules of this VBProject into a "dist" folder inside the parent folder of this workbook

    Dim component As Object
    Dim components As Object
    Dim ext As String
    Dim path As String
    
    On Error GoTo TrustCenterIssue
        Set components = ThisWorkbook.VBProject.VBComponents
    On Error GoTo 0
    
    path = CreateExportDir
    
    For Each component In components
        Select Case component.Type
            Case 1
                ext = ".bas"
            Case 2
                ext = ".cls"
            Case 3
                ext = ".frm"
            Case Else
                GoTo NextIteration
        End Select
        component.Export path & "\" & component.Name & ext
NextIteration:
    Next component
    
    Exit Sub
    
TrustCenterIssue:
        
    MsgBox "Programmatic access is not trusted!" & vbNewLine & _
        "Go to:" & vbNewLine & vbNewLine & _
        "1. File" & vbNewLine & _
        "2. Options" & vbNewLine & _
        "3. Trust Center " & vbNewLine & _
        "4. Trust Center Settings..." & vbNewLine & _
        "5. Macro Settings" & vbNewLine & vbNewLine & _
        "Check ""Trust access to the VBA project object model"".", _
        vbCritical, "Error"
    
End Sub

Private Function CreateExportDir() As String
' Creates folder for export

    Dim fso As Object
    Dim path As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    path = ThisWorkbook.path & "\dist"
    
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    
    Set fso = Nothing
    
    CreateExportDir = path
    
End Function
