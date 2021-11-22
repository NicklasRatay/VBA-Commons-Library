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
' Offers general-purpose functionality for array handling
' Changing the data type of <m_Items> to a specific class does not improve performance

Private m_Items() As Object ' Array that stores all items

' Adds an <item> to this array (at <index> if specified)
Public Sub Add(ByRef item As Object, Optional ByVal index As Long = -1)

    Dim i As Long, Size As Long
    
    Size = Me.Size
    ReDim Preserve m_Items(Size)
    
    If index = -1 Then
        index = Size
    End If
    
    ' Reposition items after <Index>
    For i = Size To index + 1 Step -1
        Set m_Items(i) = m_Items(i - 1)
    Next i
    
    Set m_Items(index) = item ' Set value at <index> to <item>
    
End Sub

' Adds all items of <ArrayListObject> to this array
Public Sub AddAll(ByVal list As ArrayListObject, Optional ByVal index As Long = -1)

    Dim i As Long, ownSize As Long, otherSize As Long
    
    ownSize = Me.Size
    otherSize = list.Size
    
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
        Set m_Items(i + index) = list.GetItem(i)
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
    
    Set Clone = arr
    
End Function

' Returns <True> when <item> occurs at least once inside this array
Public Function Contains(ByVal item As Object) As Boolean

    Dim i As Long, Size As Long
    
    If Size = 0 Then
        Exit Function
    End If
    
    For i = 0 To Size - 1
        If m_Items(i) = item Then
            Contains = True
            Exit Function
        End If
    Next i
    
End Function

' Returns the item at <index> of this array
Public Function GetItem(ByVal index As Long) As Object

    Set GetItem = m_Items(index)
    
End Function

' Returns the index of the first occurrence of <item> in this array or -1 if this array does not contain <item>
Public Function IndexOf(ByVal item As Object) As Long

    Dim i As Long
    
    If Me.Size = 0 Then
        IndexOf = -1
        Exit Function
    End If
    
    For i = 0 To Me.Size - 1
        If Me.GetItem(i) = item Then
            IndexOf = i
            Exit Function
        End If
    Next i
    
    IndexOf = -1
    
End Function

' Removes and returns item at <index> (removes last item when no <index> specified)
Public Function Remove(Optional ByVal index As Long = -1) As Object

    Dim i As Long, Size As Long
    
    Size = Me.Size
    
    If Size = 0 Then
        Exit Function
    End If
    
    If index = -1 Then index = Size - 1
    
    Set Remove = m_Items(i)
    
    ' Reposition items after <Index>
    For i = index To Size - 2
        Set m_Items(i) = m_Items(i + 1)
    Next i
    
    ' Resize array
    If Size <= 1 Then
        Dim arr() As Object
        Me.SetItems arr
    Else
        ReDim Preserve m_Items(Size - 2)
    End If
    
End Function

' Removes all items that are contained in <list> and returns <True> when at least one item has been removed
Public Function RemoveAll(ByVal list As ArrayListObject) As Boolean

    Dim i As Long, j As Long
    
    If Me.Size = 0 Or list.Size = 0 Then
        Exit Function
    End If
    
    For i = Me.Size - 1 To 0 Step -1
        For j = list.Size - 1 To 0 Step -1
            If Me.GetItem(i) Is list.GetItem(j) Then
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

    Dim duplicates() As Long
    Dim i As Long, j As Long, dupCount As Long
    
    If Me.Size = 0 Then
        Exit Function
    End If
    
    For i = 0 To Me.Size - 3
        For j = i + 1 To Me.Size - 1
            If Not duplicates.Contains(j) Then
                If Me.GetItem(i) Is Me.GetItem(j) Then
                    ReDim Preserve duplicates(dupCount)
                    duplicates(dupCount) = j
                    dupCount = dupCount + 1
                End If
            End If
        Next j
    Next i
    
    For i = 0 To UBound(duplicates)
        Me.Remove duplicates.GetItem(i)
        RemoveDuplicates = True
    Next i
    
End Function

' Removes all items from <startIndex> (inclusive) to <endIndex> (exclusive) from this array
Public Function RemoveRange(Optional ByVal startIndex As Long = 0, Optional ByVal endIndex As Long = -1) As ArrayListObject

    Dim i As Long, Size As Long, offset As Long, newSize As Long
    
    Size = Me.Size
    
    If endIndex = -1 Then
        endIndex = Size
    End If
    
    Set RemoveRange = Me.SubList(startIndex, endIndex)
    
    offset = endIndex - startIndex
    
    ' Reposition items after <StartIndex>
    For i = Size - 1 To endIndex Step -1
        Set m_Items(i - offset) = m_Items(i)
    Next i
    
    ' Resize array
    newSize = Size - offset - 1
    If newSize >= 0 Then
        ReDim Preserve m_Items(newSize)
    Else
        Dim arr() As Object
        Me.SetItems arr
    End If
    
End Function

' Removes all items that are not contained in <list> and returns <True> when at least one item has been removed
Public Function RetainAll(ByVal list As ArrayListObject) As Boolean
    
    Dim i As Long, j As Long
    Dim flag As Boolean
    
    If Me.Size = 0 Then
        Exit Function
    ElseIf list.Size = 0 Then
        Me.Clear
        RetainAll = True
        Exit Function
    End If
    
    For i = Me.Size - 1 To 0 Step -1
        flag = False
        For j = list.Size - 1 To 0 Step -1
            If Me.GetItem(i) Is list.GetItem(j) Then
                flag = True
                RetainAll = True
            End If
        Next j
        If Not flag Then
            Me.Remove i
        End If
    Next i
    
End Function

' Reverses the order of the items in this array
Public Sub Reverse()

    Dim arr() As Object
    Dim i As Long, Size As Long
    
    Size = Me.Size
    
    If Size = 0 Then
        Exit Sub
    End If
    
    ReDim arr(Size - 1)
    
    For i = 0 To Size - 1
        Set arr(i) = Me.GetItem(Size - 1 - i)
    Next i
    
    Me.SetItems arr
    
End Sub

' Replaces the item at <index> with <item> (last item if <index> is not specified)
Public Sub SetItem(ByRef item As Object, Optional ByVal index As Long = -1)

    If index = -1 Then
        index = Me.Size - 1
    End If
    
    Set m_Items(index) = item
    
End Sub

' Sets the intern array to <items>
Public Sub SetItems(items() As Object)

    m_Items = items
    
End Sub

' Returns the count of items in this array
Public Function Size() As Long

    On Error GoTo EmptyArray
        Size = UBound(m_Items) + 1
        Exit Function
    On Error GoTo 0
    
EmptyArray:

    Size = 0
    
End Function

' Sorts the items of this array
Public Sub Sort(Optional ByVal ascending As Boolean = True)

    Dim temp As Object
    Dim i As Long, j As Long
    
    If Me.Size <= 1 Then
        Exit Sub
    End If
    
    If ascending Then
        For i = 0 To Me.Size - 2
            For j = i + 1 To Me.Size - 1
                If Me.GetItem(i) > Me.GetItem(j) Then
                    temp = Me.GetItem(i)
                    Me.SetItem Me.GetItem(j), i
                    Me.SetItem temp, j
                End If
            Next j
        Next i
    Else
        For i = 0 To Me.Size - 2
            For j = i + 1 To Me.Size - 1
                If Me.GetItem(i) < Me.GetItem(j) Then
                    temp = Me.GetItem(i)
                    Me.SetItem Me.GetItem(j), i
                    Me.SetItem temp, j
                End If
            Next j
        Next i
    End If
    
End Sub

' Returns the part from <startIndex> (inclusive) to <endIndex> (exclusive) of this array
Public Function SubList(Optional ByVal startIndex As Long = 0, Optional ByVal endIndex As Long = -1) As ArrayListObject

    Dim arr As New ArrayListObject
    Dim i As Long
    
    If endIndex = -1 Then
        endIndex = Me.Size - 1
    End If
    
    ' Extract items
    For i = startIndex To endIndex - 1
        arr.Add Me.GetItem(i)
    Next i
    
    Set SubList = arr
    
End Function

' Returns the intern array
Public Function ToArray() As Object()

    ToArray = m_Items
    
End Function
