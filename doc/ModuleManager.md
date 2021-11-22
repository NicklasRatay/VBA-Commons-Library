# Description
This class provides a method to easily export all modules, classes and forms of a VBA-project. To use the method it simply has to be run from the Visual Basic Editor using `F5` while the selection-cursor is inside the method code.
  - [ModuleManager](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ModuleManager.bas)
# Methods
 - [ExportAll](#exportall)
## ExportAll
Exports all modules, classes and forms of the VBA-project this method is run in. Creates a `.\dist` directory if not already existent to store the exports in. Special modules like the workbook module for example are ignored.
 - Parameters
	 - None
 - Returns
	 - Nothing
