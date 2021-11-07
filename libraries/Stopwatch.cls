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

Private m_StartTime As Single
Private m_RunTime As Single
Private m_IsRunning As Boolean

Public Function GetRunTime() As Single
' Returns the actual time this timer was running

    GetRunTime = m_RunTime
    
End Function

Public Function GetStartTime() As Single
' Returns the stored <m_StartTime>

    GetStartTime = m_StartTime
    
End Function

Public Sub SetStartTime(Value As Single)
' Sets the <m_StartTime>

    m_StartTime = Value
    
End Sub

Public Sub Start()
' Starts the stopwatch

    If Not m_IsRunning Then
        m_StartTime = Timer
        m_IsRunning = True
    End If
    
End Sub

Public Sub Pause()
' Stores the current time in <m_EndTime> and <m_StartTime>

    If m_IsRunning Then
        m_RunTime = m_RunTime + Timer - m_StartTime
        m_IsRunning = False
    End If
    
End Sub

Public Sub PrintRunTime()
' Prints the current run time to the console

    Dim temp As Single
    
    If m_IsRunning Then
        temp = m_RunTime + Timer - m_StartTime
    Else
        temp = m_RunTime
    End If
    
    Debug.Print Format(temp, "0.000"" s""")
    
End Sub
