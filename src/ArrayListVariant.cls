VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayListVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Offers general-purpose functionality for array handling
' Variant impacts performances so the data type of <m_Items> should be changed (replace in whole class using ctrl + h)
' Changing it to object does not work because assigning an object to a variable requires the keyword <set>
' There is an extra class for that purpose

Private m_Items() As Object ' Array that stores all items

' Adds an <item> to this array (at <index> if specified)
Public Sub Add(ByRef item As Object, Optional ByVal index As Long = -1)

    Dim i As Long, size As Long
    
    size = Me.GetSize
    ReDim Preserve m_Items(size)
    
    If index = -1 Then
        index = size
    End If
    
    ' Reposition items after <Index>
    For i = size To index + 1 Step -1
        m_Items(i) = m_Items(i - 1)
    Next i
    
    m_Items(index) = item ' Set value at <index> to <item>
    
End Sub

' Adds all items of <ArrayListObject> to this array
Public Sub AddAll(ByVal list As ArrayListObject, Optional ByVal index As Long = -1)

    Dim i As Long, ownSize As Long, otherSize As Long
    
    ownSize = Me.GetSize
    otherSize = ArrayListObject.GetSize
    
    If index = -1 Then
        index = ownSize
    End If
    
    ReDim Preserve m_Items(ownSize + otherSize - 1)
    
    ' Reposition items after <Index>
    For i = ownSize + otherSize - 1 To index + otherSize Step -1
        m_Items(i) = m_Items(i - otherSize)
    Next i
    
    ' Inserts items of <ArrayListObject>
    For i = 0 To otherSize - 1
        m_Items(i + index) = ArrayListObject.GetItem(i)
    Next i
    
End Sub

' Removes all items of this array
Public Sub Clear()

    Dim arr() As Object
    
    Me.SetItems arr
    
End Sub

' Copies this ArrayListObject by value
Public Function Clone() As ArrayListObject

    Dim arr As New ArrayListObject
    
    arr.SetItems Me.ToArray
    
    Clone = arr
    
End Function

' Returns <True> when <item> occurs at least once inside this array
Public Function Contains(ByVal item As Object) As Boolean

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

' Returns the item at <index> of this array
Public Function GetItem(ByVal index As Long) As Object

    GetItem = m_Items(index)
    
End Function

' Returns the count of items in this array
Public Function GetSize() As Long

    On Error GoTo EmptyArray
        GetSize = UBound(m_Items) + 1
        Exit Function
    On Error GoTo 0
    
EmptyArray:

    GetSize = 0
    
End Function

' Returns the part from <startIndex> (inclusive) to <endIndex> (exclusive) of this array
Public Function GetSubArrayListObject(Optional ByVal startIndex As Long = 0, Optional ByVal endIndex As Long = -1) As ArrayListObject

    Dim arr As New ArrayListObject
    Dim i As Long
    
    If endIndex = -1 Then
        endIndex = Me.GetSize - 1
    End If
    
    ' Extract items
    For i = startIndex To endIndex - 1
        arr.Add Me.GetItem(i)
    Next i
    
    GetSubArrayListObject = arr
    
End Function

' Returns the index of the first occurence of <item> in this array or -1 if this array does not contain <item>
Public Function IndexOf(ByVal item As Object) As Long

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

' Removes item at <index> (removes last item when no <index> specified)
Public Sub Remove(Optional ByVal index As Long = -1)

    Dim i As Long, size As Long
    
    size = Me.GetSize
    
    If size = 0 Then
        Exit Sub
    End If
    
    If index = -1 Then index = size - 1
    
    ' Reposition items after <Index>
    For i = index To size - 2
        m_Items(i) = m_Items(i + 1)
    Next i
    
    ' Resize array
    If size = 0 Then
        Dim arr() As Object
        Me.SetItems arr
    Else
        ReDim Preserve m_Items(size - 2)
    End If
    
End Sub

' Removes all items that are contained in <list> and returns <True> when at least one item has been removed
Public Function RemoveAll(ByVal list As ArrayListObject) As Boolean

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

' Removes all duplicate items and returns <True> when at least one item has been removed
Public Function RemoveDuplicates() As Boolean

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

' Removes all items from <startIndex> (inclusive) to <endIndex> (exclusive) from this array
Public Sub RemoveRange(Optional ByVal startIndex As Long = 0, Optional ByVal endIndex As Long = -1)

    Dim i As Long, size As Long, offset As Long, newSize As Long
    
    size = Me.GetSize
    offset = endIndex - startIndex
    
    ' Reposition items after <StartIndex>
    For i = size - 1 To endIndex Step -1
        m_Items(i - offset) = m_Items(i)
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

' Removes all items that are not contained in <list> and returns <True> when at least one item has been removed
Public Function RetainAll(ByVal list As ArrayListObject) As Boolean
    
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

' Reverses the order of the items in this array
Public Sub Reverse()

    Dim arr() As Object
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

' Replaces the item at <index> with <item> (last item if <index> is not specified)
Public Sub SetItem(ByRef item As Object, Optional ByVal index As Long = -1)

    If index = -1 Then
        index = Me.GetSize - 1
    End If
    
    m_Items(index) = item
    
End Sub

' Sets the intern array to <items>
Public Sub SetItems(items() As Object)

    m_Items = items
    
End Sub

' Sorts the items of this array
Public Sub Sort(Optional ByVal ascending As Boolean = True)

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

' Returns the intern array
Public Function ToArray() As Object()

    ToArray = m_Items
    
End Function
