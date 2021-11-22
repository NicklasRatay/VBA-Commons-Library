Attribute VB_Name = "RegEx"
Option Explicit

' Returns all matches of <pattern> in <expression> or an array with an empty string as the only element if there is none
Public Function GetMatches(ByVal expression As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False) As String()

    Dim RegEx As Object, tempMatches As Object, match As Object
    Dim matches() As String
    Dim i As Integer
    
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.pattern = pattern
    RegEx.Global = True ' To find all occurences
    RegEx.ignoreCase = ignoreCase
    
    Set tempMatches = RegEx.Execute(expression) ' List of all matches as an object
    
    ' Converts object to array of strings
    For Each match In tempMatches
        ReDim Preserve matches(i)
        matches(i) = match
        i = i + 1
    Next match
    
    If i = 0 Then ReDim matches(0) ' To prevent type mismatch compile error
    
    Set RegEx = Nothing
    
    GetMatches = matches
    
End Function

' Returns <True> if <expression> contains at least one occurence of <pattern>
Public Function IsMatch(ByVal expression As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False) As Boolean

    Dim RegEx As Object
    
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.pattern = pattern
    RegEx.ignoreCase = ignoreCase
    
    IsMatch = RegEx.test(expression)
    
    Set RegEx = Nothing
    
End Function

' Replaces all occurences of <pattern> in <expression> with <replacement>
Public Function ReplaceAll(ByVal expression As String, ByVal pattern As String, ByVal replacement As String, Optional ByVal ignoreCase As Boolean = False) As String

    Dim RegEx As Object
    
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.pattern = pattern
    RegEx.Global = True ' So all occurences are replaced
    RegEx.ignoreCase = ignoreCase
    
    ReplaceAll = RegEx.Replace(expression, replacement)
    
    Set RegEx = Nothing
    
End Function
