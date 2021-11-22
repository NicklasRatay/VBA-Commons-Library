VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileManager"
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

' Creates the directory specified by <path> and all needed parent directories and returns <True> if successful
Public Function CreateDirectory(ByVal path As String) As Boolean

    Dim arr() As String, tempArr() As String
    Dim i As Integer
    Dim tempStr As String
    Dim flag As Boolean
    
    path = Replace(path, "\", "/")
    arr = Split(path, "/")
    
    For i = LBound(arr) To UBound(arr)
        tempArr = arr
        ReDim Preserve tempArr(i)
        tempStr = Join(tempArr, "/")
        If Not Exists(tempStr) Then
            On Error GoTo Failed
                fso.CreateFolder tempStr
            On Error GoTo 0
        End If
    Next i
    
    CreateDirectory = True
    
    Exit Function
    
Failed:

    CreateDirectory = False
    
End Function

' Returns <True> when the directory or file specified by <Path> exists
Public Function Exists(ByVal path As String) As Boolean

    If fso.FolderExists(path) Then
        Exists = True
    ElseIf fso.FileExists(path) Then
        Exists = True
    End If
    
End Function

' Returns the name of the directory specified by <path>
Public Function GetName(ByVal path As String) As String

    Dim arr() As String
    
    path = Replace(path, "\", "/")
    arr = Split(path, "/")
    
    GetName = arr(UBound(arr))
    
End Function
