VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Please wait..."
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PROGRESS_BAR_WIDTH As Integer = 168

' Updates progress bar
Public Sub UpdateProgressBar(ByVal CurrentValue As Long, ByVal MaxValue As Long)

    Me.lblFront.Width = PROGRESS_BAR_WIDTH * (CurrentValue / MaxValue)
    Me.lblPercentage.Caption = Round(100 * (CurrentValue / MaxValue)) & " %"
    Me.Repaint

End Sub
