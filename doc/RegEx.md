# Description
This module provides convenient access to basic functions using regular expressions. Serves as an interface for the build-in *"VBScript.RegExp"*.
  - [RegEx](https://github.com/NicklasRatay/VBA-Library/tree/main/src/RegEx.bas)
# Methods
 - [GetMatches](#getmatches)
 - [IsMatch](#ismatch)
 - [ReplaceAll](#replaceall)
## GetMatches
Returns an array containing all parts of `expression` that match the `pattern`. If there are no matches an empty array is returned. If `ignoreCase` is not specified it is set to `False`.
 - Parameters
	 - `expression` As `String` and `ByVal`
	 - `pattern` As `String` and `ByVal`
	 - `ignoreCase` As `Boolean` and `ByVal` with default of `False`
 - Returns
	 - `String()`

Example Code:
```vba
Dim expression As String, pattern As String
Dim matches() As String
Dim i As Integer

expression = "This is a test message!"
pattern = "\w*e\w*" ' A word containing an "e"

matches = RegEx.GetMatches(expression, pattern, True)

For i = LBound(matches) To UBound(matches)
	Debug.Print i & ": " & matches(i)
Next i
' 0: test
' 1: message
```
## IsMatch
Returns `True` if `expression` contains at least one occurrence of `pattern`. If `ignoreCase` is not specified it is set to `False`.
 - Parameters
	 - `expression` As `String` and `ByVal`
	 - `pattern` As `String` and `ByVal`
	 - `ignoreCase` As `Boolean` and `ByVal` with default of `False`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim expression As String, pattern As String
Dim matches() As String
Dim i As Integer

expression = "This is a test message!"
pattern = "^Th" ' Begins with "Th"

Debug.Print RegEx.IsMatch(expression, pattern, True)
' True
```
## ReplaceAll
Replaces all occurrences of `pattern` in `expression` with `replacement` and returns the result as a `String`. If `ignoreCase` is not specified it is set to `False`.
 - Parameters
	 - `expression` As `String` and `ByVal`
	 - `pattern` As `String` and `ByVal`
	 - `replacement` As `String` and `ByVal`
	 - `ignoreCase` As `Boolean` and `ByVal` with default of `False`
 - Returns
	 - `String`

Example Code:
```vba
Dim expression As String, pattern As String
Dim i As Integer

expression = "This is a test message!"
pattern = "\sis\s" ' Letters "is" isolated by white space characters

Debug.Print RegEx.ReplaceAll(expression, pattern, " was ", True)
' This was a test message!
```
