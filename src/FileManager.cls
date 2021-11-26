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

' For traversing file system
Enum DirectoryType
    FilesAndFolders
    JustFiles
    JustFolders
End Enum

' Sets reference to file system object
Private Sub Class_Initialize()

    Set fso = CreateObject("Scripting.FileSystemObject")
    
End Sub

' Releases memory
Private Sub Class_Terminate()

    Set fso = Nothing
    
End Sub

' Copies <sourcePath> to <targetPath> and overwrites if <overwrite> is set to <True> (copies recursively when used on folder)
Public Sub Copy(sourcePath As String, targetPath As String, Optional overwrite As Boolean = False)

    Dim arr() As String
    
    sourcePath = GetPath(sourcePath)
    targetPath = GetPath(targetPath)
    
    arr = Split(targetPath, "\")
    ReDim Preserve arr(UBound(arr) - 1)
    
    If fso.FolderExists(sourcePath) Then
        If CreateDirectory(Join(arr, "\")) Then
            fso.CopyFolder sourcePath, targetPath, overwrite
            Copy = True
        End If
    ElseIf fso.FileExists(sourcePath) Then
        If CreateDirectory(Join(arr, "\")) Then
            fso.CopyFile sourcePath, targetPath, overwrite
            Copy = True
        End If
    End If
    
End Sub

' Creates the directory specified by <path> and all needed parent directories and returns <True> if successful
Public Function CreateDirectory(ByVal path As String) As Boolean

    Dim arr() As String, tempArr() As String
    Dim i As Integer
    Dim tempStr As String
    Dim flag As Boolean
    
    arr = Split(GetPath(path), "\")
    
    For i = LBound(arr) To UBound(arr)
        tempArr = arr
        ReDim Preserve tempArr(i)
        tempStr = Join(tempArr, "\")
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

    path = GetPath(path)
    
    If fso.FolderExists(path) Then
        Exists = True
    ElseIf fso.FileExists(path) Then
        Exists = True
    End If
    
End Function

' Returns the parent directory of the specified <path>
Public Function GetParent(ByVal path As String) As String

    Dim arr() As String
    
    path = GetPath(path)
    
    arr = Split(path, "\")
    ReDim Preserve arr(UBound(arr) - 1)
    
    GetParent = Join(arr, "\")
    
End Function

' Returns a correctly formatted path with "." or ".." at the beginning replaced by the current or parent directory respectively
Private Function GetPath(ByVal path As String) As String

    Dim current As String
    Dim arr() As String
    
    ' So it does not matter what separator is used
    current = Replace(ThisWorkbook.path, "/", "\")
    path = Replace(path, "/", "\")
    
    ' Check for relative path
    If path Like "..*" Then ' Parent directory
        arr = Split(current, "\")
        ReDim Preserve arr(UBound(arr) - 1)
        path = Join(arr, "\") & Mid(path, 3)
    ElseIf path Like ".*" Then ' Current working directory
        path = current & Mid(path, 2)
    End If
    
    If path Like "*\" Then
        path = Left(path, Len(path) - 1)
    End If
    
    GetPath = path
    
End Function

' Returns the name of the directory specified by <path>
Public Function GetName(ByVal path As String) As String

    Dim arr() As String
    
    arr = Split(GetPath(path), "\")
    GetName = arr(UBound(arr))
    
End Function

' Returns an array of all subdirectories of the specified <path> including files if desired
Public Function GetSubDirectories(ByVal path As String, Optional ByVal dirType As DirectoryType = FilesAndFolders) As String()

    Dim arr() As String
    Dim folder As Object, f As Object
    Dim i As Integer
    
    path = GetPath(path)
    
    Set folder = fso.GetFolder(path)
    
    Select Case dirType
        Case FilesAndFolders
            For Each f In folder.SubFolders ' Folders
                ReDim Preserve arr(i)
                arr(i) = f.path
                i = i + 1
            Next f
            For Each f In folder.Files ' Files
                ReDim Preserve arr(i)
                arr(i) = f.path
                i = i + 1
            Next f
        Case JustFiles
            For Each f In folder.Files ' Files
                ReDim Preserve arr(i)
                arr(i) = f.path
                i = i + 1
            Next f
        Case JustFolders
            For Each f In folder.SubFolders ' Folders
                ReDim Preserve arr(i)
                arr(i) = f.path
                i = i + 1
            Next f
    End Select
    
    GetSubDirectories = arr
    
End Function
