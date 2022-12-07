VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4260
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cb 
      Caption         =   "¿ªËã"
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox inpp 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label sz 
      Caption         =   "HI"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim geshu, x As Single
Dim a As Single
Private Sub cb_Click()
sz.Caption = "working on it"
a = inpp.Text
geshu = 1
x = 3

Do Until geshu = a
For i = 3 To Sqr(x) Step 2
If x Mod i = 0 Then
Exit For
End If
DoEvents
Next
If i > Sqr(x) Then
geshu = geshu + 1
End If
x = x + 2
DoEvents
Loop

sz.Caption = x - 2
End Sub

Private Sub Form_Load()
sz.ForeColor = RGB(0, 0, 255)
sz.FontSize = 30
inpp.ForeColor = RGB(255, 0, 0)
inpp.FontSize = 30
cb.FontSize = 11
End Sub
