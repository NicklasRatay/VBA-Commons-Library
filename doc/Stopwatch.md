﻿# Description
This class provides functionality to measure time in a more convenient way and is mainly intended for debugging. For example it can be used to measure the time a certain block of code needs to be executed and print the result to the console.
All methods can be used while the *Stopwatch* is running or not.
  - [Stopwatch](https://github.com/NicklasRatay/VBA-Library/tree/main/src/Stopwatch.cls)
# Methods
 - [GetElapsedTime](#getelapsedtime)
 - [IsRunning](#isrunning)
 - [Start](#start)
 - [Pause](#pause)
 - [PrintElapsedTime](#printelapsedtime)
## GetElapsedTime
Returns the total time this *Stopwatch* has been running in seconds.
 - Parameters
	 - None
 - Returns
	 - `Single`

Example Code:
```vba
Dim watch As New Stopwatch

watch.Start

' Some code that is measured

Debug.Print watch.GetElapsedTime
' #.###s
```
## IsRunning
Returns `True` if this *Stopwatch* is currently running.
 - Parameters
	 - None
 - Returns
	 - `Boolean`

Example Code:
```vba
Dim watch As New Stopwatch

Debug.Print watch.IsRunning
' False

watch.Start

Debug.Print watch.IsRunning
' True

watch.Pause

Debug.Print watch.IsRunning
' False
```
## Start
Starts this *Stopwatch*. Can be used after this *Stopwatch* has been paused. Nothing happens if this *Stopwatch* is already running.
 - Parameters
	 - None
 - Returns
	 - Nothing

Example Code:
```vba
Dim watch As New Stopwatch

watch.Start
```
## Pause
Stops this *Stopwatch*. Nothing happens if this *Stopwatch* is already paused.
 - Parameters
	 - None
 - Returns
	 - Nothing

Example Code:
```vba
Dim watch As New Stopwatch

watch.Start

' Some code that is measured

watch.Pause

' Some code that is not measured

watch.Start

' Some code that is measured
```
## PrintElapsedTime
Prints the total elapsed time of this *Stopwatch* in seconds to the console including a `message`. If no `message` is specified it just prints the time.
 - Parameters
	 - `message` As `String` and `ByVal` with default of `""`
 - Returns
	 - Nothing

Example Code:
```vba
Dim watch As New Stopwatch

watch.Start

' Some code that is measured

watch.PrintElapsedTime "Code has been executed."
' #.###s | Code has been executed.
```
