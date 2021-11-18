VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Offers general-purpose functionality for file system activites

Private fso As Object

' Sets reference to file system object
Private Sub Class_Initialize()

    Set fso = CreateObject("Scripting.FileSystemObject")
    
End Sub

' Releases memory
Private Sub Class_Terminate()

    Set fso = Nothing
    
End Sub

' Returns <True> when the directory or file specified by <Path> exists
Public Function Exists(ByVal Path As String) As Boolean

    If fso.FolderExists(Path) Then
        Exists = True
    ElseIf fso.FileExists(Path) Then
        Exists = True
    End If
    
End Function
