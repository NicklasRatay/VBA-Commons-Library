# Description
This form provides easy access to a stand-alone progress bar that is useful for displaying progress of time consuming loops for example.
  - [ProgressBar](https://github.com/NicklasRatay/VBA-Library/tree/main/src/ProgressBar.frm)
# Methods
 - [Update](#update)
## Update
Calculates percentage of completion and repaints the form accordingly using a custom `message` as well. If `message` is not specified no text is displayed above the progress bar.
 - Parameters
	 - `current` As `Long` and `ByVal`
	 - `max` As `Long` and `ByVal`
	 - `message` As `String` and `ByVal` with default of `""`
 - Returns
	 - Nothing

Example Code:
```vba
Dim i As Integer, max As Integer

max = 10000

ProgressBar.Show

For i = 0 To max
	' Some code
	ProgressBar.Update i, max, "Operation " & i & " from " & max
Next i

Unload ProgressBar
```
