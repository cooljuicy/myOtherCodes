VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   4605
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton go 
      Caption         =   "go"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c As Double
Private Sub Form_Load()
Text1.Text = ""
Text1.Text = ""
Text1.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub go_Click()
a = Text1.Text
b = Text2.Text
If a < b Then
a = Text2.Text
b = Text2.Text
End If
Do Until a Mod b = 0
c = a Mod b
a = b
b = c
Loop
c = a / b
Text3.Text = c
End Sub
