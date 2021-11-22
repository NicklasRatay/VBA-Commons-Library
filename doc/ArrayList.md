# Description
There are two ArrayList classes due to the fact that the compiler requires the use of slightly different syntax when dealing with objects instead of primitive data types. 
  - [ArrayListObject](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ArrayListObject.cls) for handling objects
  - [ArrayListVariant](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ArrayListVariant) for handling primitive data types

These classes provide quality of life features for dealing with arrays like automatic resizing and advanced functions like sorting or reversing the elements.
# Performance Notice
  - ArrayListObject: This class accepts every kind of object and changing the data type of the internal array to a specific class does not improve performance. Therefore it can be used as it is for all types of objects without impacting runtime.   

  - ArrayListVariant: The data type `Variant` can accept every primitive data type but this flexibility comes with a bad impact on performance.
It is advised to change the data type of the internal array of this class if heavy workload is expected. This can be done by simply replacing the keyword `Variant` within the whole class module with the needed data type (Shortcut: `ctrl + h`) and renaming the class accordingly. An *ArrayList* class of type `String` for example should be named *ArrayListString*.

# Methods
 - [Add](#add)
 - [AddAll](#addall)
 - [Clear](#clear)
 - [Clone](#clone)
 - [Contains](#contains)
 - [GetItem](#getitem)
 - [IndexOf](#indexof)
 - [PrintItems](#printitems)
 - [Remove](#remove)
 - [RemoveAll](#removeall)
 - [RemoveDuplicates](#removeduplicates)
 - [RemoveRange](#removerange)
 - [RetainAll](#retainall)
 - [Reverse](#reverse)
 - [SetItem](#setitem)
 - [SetItems](#setitems)
 - [Size](#size)
 - [Sort](#sort)
 - [SubList](#sublist)
 - [ToArray](#toarray)
## Add
This adds an item to this *ArrayList*. If no `index` is specified it is added at the end.
 - Parameters
	 - `item` As `<Type>` and `ByRef`
	 - `index` As `Long` and `ByVal`
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3", 0

arr.PrintItems
' 0: Item3
' 1: Item1
' 2: Item2
```
## AddAll
This adds all items of a second *ArrayList* to this *ArrayList*. If no `index` is specified they are added at the end.
 - Parameters
	 - `list` As `ArrayList<Type>` and `ByVal`
	 - `index` As `Long` and `ByVal`
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2 As New ArrayListVariant

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"

arr2.AddAll arr1

arr2.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3
```
## Clear
Removes all items from this *ArrayList*.
 - Parameters
	 - None
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"

arr.Clear

arr.PrintItems
' Empty
```
## Clone
Returns an exact copy of this *ArrayList*. The clone is a new instance of this class so changes to the original are not applied to the clone and vice versa. When used with `ArrayListObject` the items of both *ArrayLists* point to the same instances.
 - Parameters
	 - None
 - Returns
	 - `ArrayList<Type>`

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2 As ArrayListVariant

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"

Set arr2 = arr1.Clone

arr1.Add "Item4" ' Not done to arr2

arr2.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3
```
## Contains
Returns `True` when the specified `item` is equal to an item inside this *ArrayList*.
 - Parameters
	 - `item` As `<Type>` and `ByVal`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

Debug.Print arr.Contains("Item2")
' True
```
## GetItem
Returns the item at the specified `index`.
 - Parameters
	 - `index` As `Long` and `ByVal`
 - Returns
	 - `<Type>`

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

arr.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3

Debug.Print arr.GetItem(1)
' Item2
```
## IndexOf
Returns the index of the first occurrence of the specified `item` inside this *ArrayList*. Returns -1 if not contained at all.
 - Parameters
	 - `item` As `<Type>` and `ByVal`
 - Returns
	 - `Long`

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

arr.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3

Debug.Print arr.IndexOf("Item3")
' 2

Debug.Print arr.IndexOf("Item4")
' -1
```
## PrintItems
Prints all items with their corresponding index to the console.
This method is only available for the *ArrayListVariant* class and its derivatives created by the user because only primitive data types are automatically converted to strings in VBA.
 - Parameters
	 - None
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

arr.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3
```
## Remove
Returns and then deletes an item from this *ArrayList*. If no `index` is specified the last item is removed.
 - Parameters
	 - `index` As `Long` and `ByVal`
 - Returns
	 - `<Type>`

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

arr.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3

Debug.Print arr.Remove(1)
' Item2

arr.PrintItems
' 0: Item1
' 2: Item3
```
## RemoveAll
Deletes all items from this *ArrayList* that equal an item of `list`. Returns `True` if at least one item has been removed.
 - Parameters
	 - `list` As `ArrayList<Type>` and `ByVal`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2 As New ArrayListVariant

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"
arr1.Add "Item4"
arr1.Add "Item5"

arr2.Add "Item2"
arr2.Add "Item5"

Debug.Print arr1.RemoveAll(arr2)

arr1.PrintItems
' 0: Item1
' 1: Item3
' 2: Item4
```
## RemoveDuplicates
Deletes all duplicate items so every item in this *ArrayList* occurs only once. Returns `True` if at least one item has been removed.
 - Parameters
	 - None
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item2"
arr.Add "Item3"
arr.Add "Item4"
arr.Add "Item4"
arr.Add "Item4"

Debug.Print arr.RemoveDuplicates
' True

arr.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3
' 3: Item4
```
## RemoveRange
Deletes and returns all items from `startIndex` (inclusive) to `endIndex` (exclusive) of this *ArrayList*. If no `startIndex` is specified it is set to 0. If no `endIndex` is specified it is set to the last index of this *ArrayList*.
 - Parameters
	 - `startIndex` As `Long` and `ByVal`
	 - `endIndex` As `Long` and `ByVal`
 - Returns
	 - `ArrayList<Type>`

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2 As ArrayListVariant

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"
arr1.Add "Item4"
arr1.Add "Item5"

Set arr2 = arr1.RemoveRange(1, 3)

arr1.PrintItems
' 0: Item1
' 1: Item4
' 2: Item5

arr2.PrintItems
' 0: Item2
' 1: Item3
```
## RetainAll
Deletes all items from this *ArrayList* that do not equal an item of `list`. Returns `True` if at least one item has been removed.
 - Parameters
	 - `list` As `ArrayList<Type>` and `ByVal`
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2 As New ArrayListVariant

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"
arr1.Add "Item4"
arr1.Add "Item5"

arr2.Add "Item2"
arr2.Add "Item5"

Debug.Print arr1.RetainAll(arr2)

arr1.PrintItems
' 0: Item2
' 1: Item5
```
## Reverse
Reverses the order of the items of this *ArrayList*.
 - Parameters
	 - None
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"
arr.Add "Item4"
arr.Add "Item5"

arr.Reverse

arr.PrintItems
' 0: Item5
' 1: Item4
' 2: Item3
' 3: Item2
' 4: Item1
```
## SetItem
Replaces the item at `index` with `item`. If no `index` is specified the last item is replaced.
 - Parameters
	 - `item` As `<Type>` and `ByRef`
	 - `index` As `Long` and `ByVal`
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

arr.SetItem "Item4", 1

arr.PrintItems
' 0: Item1
' 1: Item4
' 2: Item3
```
## SetItems
Converts the specified `items` array to an *ArrayList* by setting the items of this *ArrayList* to the array's items. 
 - Parameters
	 - `items` As `<Type>()` and `ByVal`
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2() As Variant

arr2 = Array("Item1", "Item2", "Item3")

arr1.SetItems arr2

arr1.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3
```
## Size
Returns the count of items inside this *ArrayList*.
 - Parameters
	 - None
 - Returns
	 - `Long`

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "Item1"
arr.Add "Item2"
arr.Add "Item3"

arr.PrintItems
'0: Item1
'1: Item2
'2: Item3

Debug.Print arr.Size
'3
```
## Sort
Sorts the items of this *ArrayList* using the build-in comparison functionality of VBA. Strings are sorted alphabetically while numbers are sorted by their value for example. If `ascending` is set to `True` the items are sorted in ascending order. If set to `False` they are sorted in descending order. If `ascending` is not specified it is set to `True`.
This method is only available for the *ArrayListVariant* class and its derivatives created by the user because only primitive data types can be relationally compared in VBA.
 - Parameters
	 - None
 - Returns
	 - Nothing

Example Code 1:
```vba
Dim arr As New ArrayListVariant

arr.Add 420
arr.Add 1337
arr.Add 666
arr.Add 69
arr.Add 404

arr.Sort True

arr.PrintItems
' 0: 69
' 1: 404
' 2: 420
' 3: 666
' 4: 1337
```
Example Code 2:
```vba
Dim arr As New ArrayListVariant

arr.Add "Cherry"
arr.Add "orange"
arr.Add "Strawberry"
arr.Add "apple"
arr.Add "Banana"

arr.Sort True

arr.PrintItems
' 0: Banana
' 1: Cherry
' 2: Strawberry
' 3: apple
' 4: orange
```
## SubList
Returns a new *ArrayList* containing only the items of this *ArrayList* from `startIndex` (inclusive) to `endIndex` (exclusive).
 - Parameters
	 - None
 - Returns
	 - `ArrayList<Type>`

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2 As ArrayListVariant

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"
arr1.Add "Item4"
arr1.Add "Item5"

arr1.PrintItems
' 0: Item1
' 1: Item2
' 2: Item3
' 3: Item4
' 4: Item5

Set arr2 = arr1.SubList(2, 4)

arr2.PrintItems
' 0: Item3
' 1: Item4
```
## ToArray
Converts this *ArrayList* to a normal array.
 - Parameters
	 - None
 - Returns
	 - `<Type>()`

Example Code:
```vba
Dim arr1 As New ArrayListVariant
Dim arr2() As Variant
Dim i As Integer

arr1.Add "Item1"
arr1.Add "Item2"
arr1.Add "Item3"

arr2 = arr1.ToArray

For i = LBound(arr2) To UBound(arr2)
	Debug.Print i & ": " & arr2(i)
Next i
' 0: Item1
' 1: Item2
' 2: Item3
```
