# ArrayList
This class provides quality of life features for dealing with arrays like automatic resizing and advanced functions like sorting or reversing the elements.

There are actually two ArrayList classes due to the fact that the compiler requires the keyword `Set` for assigning an object to a variable while it must not be used for assigning primitive data type values.
  - [ArrayListObject](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ArrayListObject.cls) for handling objects
  - [ArrayListVariant](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ArrayListVariant) for handling primitive data types

## Performance Notice
  - ArrayListObject: This class accepts every kind of object and changing the data type of the internal array to a specific class does not improve performance. Therefore it can be used as it is for all types of objects without impacting runtime.

  - ArrayListVariant: The data type `Variant` can accept every primitive data type but this flexibility comes with a bad impact on performance.
It is advised to change the data type of the internal array of this class if heavy workload is expected. This can be done by simply replacing the keyword `Variant` within the whole class module with the needed data type (Shortcut: `ctrl + h`) and renaming the class accordingly. An array of type `String` for example should be named *ArrayListString*.

## Methods
### Add
This adds an element to this *ArrayList*. If no `index` is specified it is added at the end.
 - Parameters
	 - `item` As `<Type>` and `ByRef`
	 - `index` As `Long` and `ByVal`
 - Returns
	 - Nothing

Example Code:
```vba
Dim arr As New ArrayListVariant

arr.Add "First insertion"
arr.Add "Second insertion"
arr.Add "Insertion at the start", 0
```
### AddAll
This adds all elements of a second *ArrayList* to this *ArrayList*. If no `index` is specified they are added at the end.
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
```
### Clear
Removes all elements from this *ArrayList*.
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
```
### Remaining methods will be added in the future
