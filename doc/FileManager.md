# Description
This class provides utility functions for dealing with the file system and common IO-Operations. Among other things it is an interface for the `Scripting.FileSystemObject`.

It does not matter if `\` or `/` is used as separator when specifying paths. By using `.` or `..` at the beginning the current working directory or its parent directory can be referenced respectively. This applies to all methods.

Valid string representations of a path:
 - "C:\\Windows\\explorer.exe"
 - "./resourcesInCurrent/testImage.jpg"
 - ".\.\\resourcesInParent\\testImage.jpg"
 - "."
 - ".\."

This class makes use of `ThisWorkbook.Path` to get the current working directory or its parent directory. To use this class inside other applications like *Word* or *Access* every occurrence of `ThisWorkbook.Path` in this class has to be replaced accordingly (`Ctrl + H`).
  - [FileManager](https://github.com/NicklasRatay/VBA-Library/tree/main/src/FileManager.cls)
# Methods
 - [Copy](#copy)
 - [CreateDirectory](#createdirectory)
 - [Exists](#exists)
 - [GetParent](#getparent)
 - [GetName](#getname)
 - [GetSubDirectories](#getsubdirectories)
## Copy
Copies the file or folder specified by `sourcePath` to `targetPath`. Folders are copied recursively. Creates all nonexistent directories of `targetPath`. Overwrites existing files if `overwrite` is set to `True`. If not specified it is set to `False`. An error is thrown when it is set to `False` and `targetPath` already exists.
 - Parameters
	 - `sourcePath` As `String` and `ByVal`
	 - `targetPath` As `String` and `ByVal`
	 - `overwrite` As `Boolean` and `ByVal` with default of `False`
 - Returns
	 - Nothing

Example Code:
```vba
Dim f As New FileManager

f.Copy ThisWorkbook.FullName, ".\New Directory\Test.xlsm", False
```
## CreateDirectory
Creates the directory specified by `path` and all needed parent directories. Returns `True` if the operation was successful.
 - Parameters
	 - `path` As `String` and `ByVal`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim f As New FileManager

Debug.Print f.CreateDirectory(".\FileManager\New Directory")
' True
```
## Exists
Returns `True` if the file or folder specified by `path` exists.
 - Parameters
	 - `path` As `String` and `ByVal`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim f As New FileManager

Debug.Print f.Exists("C:\Windows\explorer.exe")
' True
```
## GetParent
Returns the path of the parent directory of the directory specified by `path`. 
 - Parameters
	 - `path` As `String` and `ByVal`
 - Returns
	 - `String`

Example Code:
```vba
Dim f As New FileManager

Debug.Print f.GetParent("C:\Windows\explorer.exe")
' C:\Windows
```
## GetName
Returns the name of the directory specified by `path`. Includes the file extension.
 - Parameters
	 - `path` As `String` and `ByVal`
 - Returns
	 - `String`

Example Code:
```vba
Dim f As New FileManager

Debug.Print f.GetName("C:\Windows\explorer.exe")
' explorer.exe
```
## GetSubDirectories
Returns an array of the sub directories of the folder specified by `path`. `dirType` can be set to one of the following to control what is listed:
 - FilesAndFolders
 - JustFiles
 - JustFolders

This method is shallow and does not list sub directories of sub directories.
 - Parameters
	 - `path` As `String` and `ByVal`
	 - `dirType` As `DirectoryType` and `ByVal` with default of `FilesAndFolders`
 - Returns
	 - `String()`

Example Code:
```vba
Dim f As New FileManager
Dim arr() As String
Dim i As Integer

arr = f.GetSubDirectories("C:\Users\Default", FilesAndFolders)

For i = LBound(arr) To UBound(arr)
	Debug.Print arr(i)
Next i
' C:\Users\Default\AppData
' C:\Users\Default\Application Data
' C:\Users\Default\Cookies
' C:\Users\Default\Desktop
' C:\Users\Default\Documents
' C:\Users\Default\Downloads
' ...
```
