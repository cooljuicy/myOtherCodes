VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   11760
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   6480
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5280
      Top             =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Private Sub Timer1_Timer()
Dim a As Form
Set a = New Form1
a.Show vbmodel
End Sub

Private Sub Timer2_Timer()
SetCursorPos 500, 50
End Sub
