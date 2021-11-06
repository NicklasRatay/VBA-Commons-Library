VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBar 
   Caption         =   "Bitte warten"
   ClientHeight    =   615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3735
   OleObjectBlob   =   "frmProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PROGRESS_BAR_WIDTH As Integer = 168

Public Sub UpdateProgressBar(ByVal CurrentValue As Long, MaxValue As Long)
' Updates progress bar

    Me.lblFront.Width = PROGRESS_BAR_WIDTH * (CurrentValue / MaxValue)
    Me.lblPercentage.Caption = Round(100 * (CurrentValue / MaxValue)) & " %"
    Me.Repaint

End Sub
