VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_StartTime As Date
Private m_EndTime As Date
Private m_RunTime As Date
Private m_IsRunning As Boolean

Public Function GetEndTime() As Date
' Returns the stored <m_EndTime>

    GetEndTime = m_EndTime
    
End Function

Public Function GetRunTime() As Date
' Returns the actual time this timer was running

    GetRunTime = m_RunTime
    
End Function

Public Function GetStartTime() As Date
' Returns the stored <m_StartTime>

    GetStartTime = m_StartTime
    
End Function

Public Sub SetEndTime(Value As Date)
' Sets the <m_EndTime>

    m_EndTime = Value
    
End Sub

Public Sub SetStartTime(Value As Date)
' Sets the <m_StartTime>

    m_StartTime = Value
    
End Sub

Public Sub StartTimer()
' Starts the timer

    If Not m_IsRunning Then
        If CDec(m_StartTime) = 0 Then
            m_StartTime = Time
            m_StartTime = m_StartTime
        Else
            m_StartTime = Time
        End If
        
        m_IsRunning = True
    End If
    
End Sub

Public Sub StopTimer()
' Stores the current time in <m_EndTime> and <m_StartTime>

    If m_IsRunning Then
        m_EndTime = Time
        m_RunTime = m_RunTime + m_EndTime - m_StartTime
        m_IsRunning = False
    End If
    
End Sub

Public Sub PrintRunTime()
' Prints the current run time to the console

    Dim temp As Date
    Dim hours As Integer, minutes As Integer
    Dim seconds As Double
    
    If m_IsRunning Then
        temp = m_RunTime + Time - m_StartTime
    Else
        temp = m_RunTime
    End If
    
    Debug.Print Format(temp, "hh""h ""mm""m ""ss""s""") & " | " & CDec(temp) * 100000 & "s"
    
End Sub
