# ArrayList
There are two ArrayList classes due to the fact that the compiler requires the keyword `Set` for assigning an object to a variable while it must not be used for assigning primitive data type values. 
  - [ArrayListObject](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ArrayListObject.cls) for handling objects
  - [ArrayListVariant](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ArrayListVariant) for handling primitive data types

## Performance Notice
  - ArrayListObject
This class accepts every kind of object and changing the data type of the internal array to a specific class does not improve performance. Therefore it can be used as it is for all types of objects without impacting runtime.
  - ArrayListVariant
The data type `Variant` can accept every primitive data type but this flexibility comes with a bad impact on performance.
It is advised to change the data type of the internal array of this class if heavy workload is expected. This can be done by simply replacing the keyword `Variant` within the whole class module with the needed data type (Shortcut: `ctrl + h`).

## Methods

