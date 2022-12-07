VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   10530
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Hello World!"
      Height          =   1575
      Left            =   6480
      TabIndex        =   0
      Top             =   3000
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Print "Hello World!"
End Sub

