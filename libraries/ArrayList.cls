VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Elements() As Variant ' Array that stores all items

Public Sub Add(ByRef Item As Variant, Optional ByVal Index As Long = -1)
' Adds an <Item> to this array (at <Index> if specified)

    Dim i As Long, size As Long
    
    size = Me.GetSize
    ReDim Preserve Elements(size)
    
    If Index = -1 Then
        Index = size
    End If
    
    ' Reposition items after <Index>
    For i = size To Index + 1 Step -1
        Elements(i) = Elements(i - 1)
    Next i
    
    Elements(Index) = Item ' Set value at <Index> to <Item>
    
End Sub

Public Sub AddAll(ByVal ArrayList As ArrayList, Optional ByVal Index As Long = -1)
' Adds all items of <ArrayList> to this array

    Dim i As Long, ownSize As Long, otherSize As Long
    
    ownSize = Me.GetSize
    otherSize = ArrayList.GetSize
    
    If Index = -1 Then
        Index = ownSize
    End If
    
    ReDim Preserve Elements(ownSize + otherSize - 1)
    
    ' Reposition items after <Index>
    For i = ownSize + otherSize - 1 To Index + otherSize Step -1
        Elements(i) = Elements(i - otherSize)
    Next i
    
    ' Inserts items of <ArrayList>
    For i = 0 To otherSize - 1
        Elements(i + Index) = ArrayList.GetItem(i)
    Next i
    
End Sub

Public Sub Clear()
' Removes all items of this array

    Dim arr() As Variant
    
    Me.SetItems arr
    
End Sub

Public Function Clone() As ArrayList
' Copies this ArrayList by value

    Dim arr As New ArrayList
    
    arr.SetItems Me.ToArray
    
    Set Clone = arr
    
End Function

Public Function Contains(ByVal Item As Variant) As Boolean
' Returns <True> when <Item> occurs at least once inside this array

    Dim i As Long
    
    If Me.GetSize = 0 Then
        Exit Sub
    End If
    
    For i = 0 To Me.GetSize
        If Elements(i) = Item Then
            Contains = True
            Exit Function
        End If
    Next i
    
End Function

Public Function GetItem(ByVal Index As Long) As Variant
' Returns the item at <Index> of this array

    GetItem = Elements(Index)
    
End Function

Public Function GetSize() As Long
' Returns the count of items in this array

    On Error GoTo EmptyArray
        GetSize = UBound(Elements) + 1
        Exit Function
    On Error GoTo 0
    
EmptyArray:

    GetSize = 0
    
End Function

Public Function GetSubArrayList(Optional ByVal StartIndex As Long = 0, Optional ByVal EndIndex As Long = -1) As ArrayList
' Returns the part from <StartIndex> (inclusive) to <EndIndex> (exclusive) of this array

    Dim arr As New ArrayList
    Dim i As Long
    
    If EndIndex = -1 Then
        EndIndex = Me.GetSize - 1
    End If
    
    ' Extract items
    For i = StartIndex To EndIndex - 1
        arr.Add Me.GetItem(i)
    Next i
    
    Set GetSubArrayList = arr
    
End Function

Public Function IndexOf(ByVal Item As Variant) As Long
' Returns the index of the first occurence of <Item> in this array or -1 if this array does not contain <Item>

    Dim i As Long
    
    If Me.GetSize = 0 Then
        IndexOf = -1
        Exit Function
    End If
    
    For i = 0 To Me.GetSize - 1
        If Me.GetItem(i) = Item Then
            IndexOf = i
            Exit Function
        End If
    Next i
    
    IndexOf = -1
    
End Function

Public Sub Remove(Optional ByVal Index As Long = -1)
' Removes item at <Index> (removes last item when no <Index> specified)

    Dim i As Long, size As Long
    
    size = Me.GetSize
    
    If size = 0 Then
        Exit Sub
    End If
    
    If Index = -1 Then Index = size - 1
    
    ' Reposition items after <Index>
    For i = Index To size - 2
        Elements(i) = Elements(i + 1)
    Next i
    
    ' Resize array
    If size = 0 Then
        Dim arr() As Variant
        Me.SetItems arr
    Else
        ReDim Preserve Elements(size - 2)
    End If
    
End Sub

Public Function RemoveAll(ByVal List As ArrayList) As Boolean
' Removes all items that are contained in <List> and returns <True> when at least one item has been removed

    Dim i As Long, j As Long
    
    If Me.GetSize = 0 Or List.GetSize = 0 Then
        Exit Sub
    End If
    
    For i = Me.GetSize - 1 To 0 Step -1
        For j = List.GetSize - 1 To 0 Step -1
            If Me.GetItem(i) = List.GetItem(j) Then
                Me.Remove i
                RemoveAll = True
                GoTo NextIteration
            End If
        Next j
NextIteration:
    Next i
    
End Function

Public Function RemoveDuplicates() As Boolean
' Removes all duplicate items and returns <True> when at least one item has been removed

    Dim duplicates As New ArrayList
    Dim i As Long, j As Long
    
    If Me.GetSize = 0 Then
        Exit Sub
    End If
    
    For i = 0 To Me.GetSize - 1
        For j = i + 1 To Me.GetSize - 1
            If Not duplicates.Contains(j) Then
                If Me.GetItem(i) = Me.GetItem(j) Then
                    duplicates.Add j
                End If
            End If
        Next j
    Next i
    
    For i = 0 To duplicates.GetSize - 1
        Me.Remove duplicates.GetItem(i)
        RemoveDuplicates = True
    Next i
    
End Function

Public Sub RemoveRange(Optional ByVal StartIndex As Long = 0, Optional ByVal EndIndex As Long = -1)
' Removes all items from <StartIndex> (inclusive) to <EndIndex> (exclusive) from this array

    Dim i As Long, size As Long, offset As Long, newSize As Long
    
    size = Me.GetSize
    offset = EndIndex - StartIndex
    
    ' Reposition items after <StartIndex>
    For i = size - 1 To EndIndex Step -1
        Elements(i - offset) = Elements(i)
    Next i
    
    ' Resize array
    newSize = size - offset - 1
    If newSize >= 0 Then
        ReDim Preserve Elements(newSize)
    Else
        Dim arr() As Variant
        Me.SetItems arr
    End If
    
End Sub

Public Function RetainAll(ByVal List As ArrayList) As Boolean
' Removes all items that are not contained in <List> and returns <True> when at least one item has been removed
    
    Dim i As Long, j As Long
    
    If Me.GetSize = 0 Or List.GetSize = 0 Then
        Exit Sub
    End If
    
    For i = Me.GetSize - 1 To 0 Step -1
        For j = List.GetSize - 1 To 0 Step -1
            If Me.GetItem(i) <> List.GetItem(j) Then
                Me.Remove i
                RetainAll = True
                GoTo NextIteration
            End If
        Next j
NextIteration:
    Next i
    
End Function

Public Sub Reverse()
' Reverses the order of the items in this array

    Dim arr() As Variant
    Dim i As Long, size As Long
    
    size = Me.GetSize
    
    If size = 0 Then
        Exit Sub
    End If
    
    ReDim arr(size - 1)
    
    For i = 0 To size - 1
        arr(i) = Me.GetItem(size - 1 - i)
    Next i
    
    Me.SetItems arr
    
End Sub

Public Sub SetItem(ByRef Item As Variant, Optional ByVal Index As Long = -1)
' Replaces the item at <Index> with <Item> (last item if <Index> is not specified)

    If Index = -1 Then
        Index = Me.GetSize - 1
    End If
    
    Elements(Index) = Item
    
End Sub

Public Sub SetItems(Items() As Variant)
' Sets this array to <Items>

    Elements = Items
    
End Sub

Public Sub Sort(Optional ByVal Ascending As Boolean = True)
' Sorts the items of this array

    Dim temp As Variant
    Dim i As Long, j As Long
    
    If Me.GetSize <= 1 Then
        Exit Sub
    End If
    
    If Ascending Then
        For i = 0 To Me.GetSize - 2
            For j = i + 1 To Me.GetSize - 1
                If Me.GetItem(i) > Me.GetItem(j) Then
                    temp = Me.GetItem(i)
                    Me.SetItem Me.GetItem(j), i
                    Me.SetItem temp, j
                End If
            Next j
        Next i
    Else
        For i = 0 To Me.GetSize - 2
            For j = i + 1 To Me.GetSize - 1
                If Me.GetItem(i) < Me.GetItem(j) Then
                    temp = Me.GetItem(i)
                    Me.SetItem Me.GetItem(j), i
                    Me.SetItem temp, j
                End If
            Next j
        Next i
    End If
    
End Sub

Public Function ToArray() As Variant()
' Returns the array

    ToArray = Elements
    
End Function
