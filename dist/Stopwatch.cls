VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Measures elapsed time in seconds
' Can be started after being stopped (e.g. to exclude lines of code from the time measurement)

Private m_LastStartTime As Single
Private m_Elapsedtime As Single
Private m_IsRunning As Boolean

' Returns the actual time this timer was running
Public Function GetElapsedTime() As Single

    If m_IsRunning Then
        GetElapsedTime = m_Elapsedtime + Timer - m_LastStartTime
    Else
        GetElapsedTime = m_Elapsedtime
    End If
    
End Function

' Returns <True> if this stopwatch is running
Public Function IsRunning() As Boolean

    IsRunning = m_IsRunning
    
End Function

' Starts the stopwatch
Public Sub Start()

    If Not m_IsRunning Then
        m_LastStartTime = Timer
        m_IsRunning = True
    End If
    
End Sub

' Stores the current time in <m_LastStartTime>
Public Sub Pause()

    If m_IsRunning Then
        m_Elapsedtime = m_Elapsedtime + Timer - m_LastStartTime
        m_IsRunning = False
    End If
    
End Sub

' Prints the current <m_Elapsedtime> to the console
Public Sub PrintElapsedtime(Optional ByVal message As String = "")

    Dim temp As Single
    
    If m_IsRunning Then
        temp = m_Elapsedtime + Timer - m_LastStartTime
    Else
        temp = m_Elapsedtime
    End If
    
    If message = "" Then
        Debug.Print Format(temp, "0.000"" s""")
    Else
        Debug.Print Format(temp, "0.000"" s""") & " | " & message
    End If
    
End Sub
