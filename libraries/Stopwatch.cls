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

Private m_StartTime As Single
Private m_Elapsedtime As Single
Private m_IsRunning As Boolean

Public Function GetElapsedTime() As Single
' Returns the actual time this timer was Elapsedning

    GetElapsedTime = m_Elapsedtime
    
End Function

Public Function GetStartTime() As Single
' Returns the stored <m_StartTime>

    GetStartTime = m_StartTime
    
End Function

Public Sub Start()
' Starts the stopwatch

    If Not m_IsRunning Then
        m_StartTime = Timer
        m_IsRunning = True
    End If
    
End Sub

Public Sub Pause()
' Stores the current time in <m_StartTime>

    If m_IsRunning Then
        m_Elapsedtime = m_Elapsedtime + Timer - m_StartTime
        m_IsRunning = False
    End If
    
End Sub

Public Sub PrintElapsedtime()
' Prints the current <m_Elapsedtime> to the console

    Dim temp As Single
    
    If m_IsRunning Then
        temp = m_Elapsedtime + Timer - m_StartTime
    Else
        temp = m_Elapsedtime
    End If
    
    Debug.Print Format(temp, "0.000"" s""")
    
End Sub
