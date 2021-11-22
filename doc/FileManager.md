# Description
This class provides utility functions for dealing with the file system and common IO-Operations. Among other things it is an interface for the `Scripting.FileSystemObject`. It does not matter if "/" or "\\" is used as separator for specifying paths. This applies to all methods.
  - [FileManager](https://github.com/NicklasRatay/VBA-Library/tree/main/src/FileManager.cls)
# Methods
 - [CreateDirectory](#createdirectory)
 - [Exists](#exists)
 - [GetName](#getname)
## CreateDirectory
Creates the directory specified by `path` and all needed parent directories. Returns `True` if the operation was successful.
 - Parameters
	 - `path` As `String` and `ByVal`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim f As New FileManager

Debug.Print f.CreateDirectory("C:/IOManager Test/New Directory")
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

Debug.Print f.Exists("C:/Windows/explorer.exe")
' True
```
## GetName
Returns the name of the file or folder specified by `path`. Includes the file extension.
 - Parameters
	 - `path` As `String` and `ByVal`
 - Returns
	 - `String`

Example Code:
```vba
Dim f As New FileManager

Debug.Print f.GetName("C:/Windows/explorer.exe")
' explorer.exe
```
