Attribute VB_Name = "RegEx"
Option Explicit

' Replaces all occurences of <Pattern> in <Expression> with <Replacement>
Public Function ReplaceAll(ByVal Expression As String, ByVal Pattern As String, ByVal Replacement As String) As String

    Dim regEx As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = Pattern
    regEx.Global = True ' So all occurences are replaced
    
    ReplaceAll = regEx.Replace(Expression, Replacement)
    
End Function

' Returns <True> if <Expression> contains at least one occurence of <Pattern>
Public Function IsMatch(ByVal Expression As String, ByVal Pattern As String) As Boolean

    Dim regEx As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = Pattern
    
    IsMatch = regEx.test(Expression)
    
End Function

' Returns all matches of <Pattern> in <Expression> or an array with an empty string as the only element if there is none
Public Function GetMatches(ByVal Expression As String, ByVal Pattern As String) As String()

    Dim regEx As Object, tempMatches As Object, match As Object
    Dim matches() As String
    Dim i As Integer
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = Pattern
    regEx.Global = True ' To find all occurences
    
    Set tempMatches = regEx.Execute(Expression) ' List of all matches as an object
    
    ' Converts object to array of strings
    For Each match In tempMatches
        ReDim Preserve matches(i)
        matches(i) = match
        i = i + 1
    Next match
    
    If i = 0 Then ReDim matches(0) ' To prevent type mismatch compile error
    
    GetMatches = matches
    
End Function

