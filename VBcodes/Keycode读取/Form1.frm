VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   11430
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Label1.Caption = KeyCode
End Sub

Private Sub Form_Load()
Label1.FontSize = 160
Label1.ForeColor = RGB(255, 50, 170)
End Sub
