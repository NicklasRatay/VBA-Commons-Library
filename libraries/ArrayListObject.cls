VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayListObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Changing the data type of <m_Items> to a specific class does not improve performance

Private m_Items() As Object ' Array that stores all items

Public Sub Add(ByRef item As Object, Optional ByVal index As Long = -1)
' Adds an <Item> to this array (at <Index> if specified)

    Dim i As Long, size As Long
    
    size = Me.GetSize
    ReDim Preserve m_Items(size)
    
    If index = -1 Then
        index = size
    End If
    
    ' Reposition items after <Index>
    For i = size To index + 1 Step -1
        Set m_Items(i) = m_Items(i - 1)
    Next i
    
    Set m_Items(index) = item ' Set value at <Index> to <Item>
    
End Sub

Public Sub AddAll(ByVal list As ArrayListObject, Optional ByVal index As Long = -1)
' Adds all items of <ArrayListObject> to this array

    Dim i As Long, ownSize As Long, otherSize As Long
    
    ownSize = Me.GetSize
    otherSize = ArrayListObject.GetSize
    
    If index = -1 Then
        index = ownSize
    End If
    
    ReDim Preserve m_Items(ownSize + otherSize - 1)
    
    ' Reposition items after <Index>
    For i = ownSize + otherSize - 1 To index + otherSize Step -1
        Set m_Items(i) = m_Items(i - otherSize)
    Next i
    
    ' Inserts items of <ArrayListObject>
    For i = 0 To otherSize - 1
        Set m_Items(i + index) = ArrayListObject.GetItem(i)
    Next i
    
End Sub

Public Sub Clear()
' Removes all items of this array

    Dim arr() As Object
    
    Me.SetItems arr
    
End Sub

Public Function Clone() As ArrayListObject
' Copies this ArrayListObject by value

    Dim arr As New ArrayListObject
    
    arr.SetItems Me.ToArray
    
    Set Clone = arr
    
End Function

Public Function Contains(ByVal item As Object) As Boolean
' Returns <True> when <Item> occurs at least once inside this array

    Dim i As Long
    
    If Me.GetSize = 0 Then
        Exit Function
    End If
    
    For i = 0 To Me.GetSize
        If m_Items(i) = item Then
            Contains = True
            Exit Function
        End If
    Next i
    
End Function

Public Function GetItem(ByVal index As Long) As Object
' Returns the item at <Index> of this array

    GetItem = m_Items(index)
    
End Function

Public Function GetSize() As Long
' Returns the count of items in this array

    On Error GoTo EmptyArray
        GetSize = UBound(m_Items) + 1
        Exit Function
    On Error GoTo 0
    
EmptyArray:

    GetSize = 0
    
End Function

Public Function GetSubArrayListObject(Optional ByVal startIndex As Long = 0, Optional ByVal endIndex As Long = -1) As ArrayListObject
' Returns the part from <StartIndex> (inclusive) to <EndIndex> (exclusive) of this array

    Dim arr As New ArrayListObject
    Dim i As Long
    
    If endIndex = -1 Then
        endIndex = Me.GetSize - 1
    End If
    
    ' Extract items
    For i = startIndex To endIndex - 1
        arr.Add Me.GetItem(i)
    Next i
    
    Set GetSubArrayListObject = arr
    
End Function

Public Function IndexOf(ByVal item As Object) As Long
' Returns the index of the first occurence of <Item> in this array or -1 if this array does not contain <Item>

    Dim i As Long
    
    If Me.GetSize = 0 Then
        IndexOf = -1
        Exit Function
    End If
    
    For i = 0 To Me.GetSize - 1
        If Me.GetItem(i) = item Then
            IndexOf = i
            Exit Function
        End If
    Next i
    
    IndexOf = -1
    
End Function

Public Sub Remove(Optional ByVal index As Long = -1)
' Removes item at <Index> (removes last item when no <Index> specified)

    Dim i As Long, size As Long
    
    size = Me.GetSize
    
    If size = 0 Then
        Exit Sub
    End If
    
    If index = -1 Then index = size - 1
    
    ' Reposition items after <Index>
    For i = index To size - 2
        Set m_Items(i) = m_Items(i + 1)
    Next i
    
    ' Resize array
    If size = 0 Then
        Dim arr() As Object
        Me.SetItems arr
    Else
        ReDim Preserve m_Items(size - 2)
    End If
    
End Sub

Public Function RemoveAll(ByVal list As ArrayListObject) As Boolean
' Removes all items that are contained in <List> and returns <True> when at least one item has been removed

    Dim i As Long, j As Long
    
    If Me.GetSize = 0 Or list.GetSize = 0 Then
        Exit Function
    End If
    
    For i = Me.GetSize - 1 To 0 Step -1
        For j = list.GetSize - 1 To 0 Step -1
            If Me.GetItem(i) = list.GetItem(j) Then
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

    Dim duplicates As New ArrayListObject
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

Public Sub RemoveRange(Optional ByVal startIndex As Long = 0, Optional ByVal endIndex As Long = -1)
' Removes all items from <StartIndex> (inclusive) to <EndIndex> (exclusive) from this array

    Dim i As Long, size As Long, offset As Long, newSize As Long
    
    size = Me.GetSize
    offset = endIndex - startIndex
    
    ' Reposition items after <StartIndex>
    For i = size - 1 To endIndex Step -1
        Set m_Items(i - offset) = m_Items(i)
    Next i
    
    ' Resize array
    newSize = size - offset - 1
    If newSize >= 0 Then
        ReDim Preserve m_Items(newSize)
    Else
        Dim arr() As Object
        Me.SetItems arr
    End If
    
End Sub

Public Function RetainAll(ByVal list As ArrayListObject) As Boolean
' Removes all items that are not contained in <List> and returns <True> when at least one item has been removed
    
    Dim i As Long, j As Long
    
    If Me.GetSize = 0 Or list.GetSize = 0 Then
        Exit Function
    End If
    
    For i = Me.GetSize - 1 To 0 Step -1
        For j = list.GetSize - 1 To 0 Step -1
            If Me.GetItem(i) <> list.GetItem(j) Then
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

    Dim arr() As Object
    Dim i As Long, size As Long
    
    size = Me.GetSize
    
    If size = 0 Then
        Exit Sub
    End If
    
    ReDim arr(size - 1)
    
    For i = 0 To size - 1
        Set arr(i) = Me.GetItem(size - 1 - i)
    Next i
    
    Me.SetItems arr
    
End Sub

Public Sub SetItem(ByRef item As Object, Optional ByVal index As Long = -1)
' Replaces the item at <Index> with <Item> (last item if <Index> is not specified)

    If index = -1 Then
        index = Me.GetSize - 1
    End If
    
    Set m_Items(index) = item
    
End Sub

Public Sub SetItems(items() As Object)
' Sets this array to <Items>

    Set m_Items = items
    
End Sub

Public Sub Sort(Optional ByVal ascending As Boolean = True)
' Sorts the items of this array

    Dim temp As Object
    Dim i As Long, j As Long
    
    If Me.GetSize <= 1 Then
        Exit Sub
    End If
    
    If ascending Then
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

Public Function ToArray() As Object()
' Returns the array

    ToArray = m_Items
    
End Function