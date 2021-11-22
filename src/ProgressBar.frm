VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Please wait..."
   ClientHeight    =   840
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
Public Sub Update(ByVal current As Long, ByVal max As Long, Optional ByVal message As String = "")

    lblFront.Width = PROGRESS_BAR_WIDTH * (current / max)
    lblPercentage.Caption = Round(100 * (current / max)) & " %"
    lblMessage.Caption = message
    Repaint

End Sub
