VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   11445
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   855
      Left            =   8040
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   855
      Left            =   8040
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   1095
      Left            =   3840
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   975
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Single
Private Sub Command1_Click()
a = Text1.Text
b = Text2.Text
Label1.Caption = a / 6 & "e" & b - 1
Label2.Caption = a & "e" & b
Label3.Caption = a * 6 & "e" & b + 1
Label4.Caption = a * 4 & "e" & b + 1
Label5.Caption = a * 6 * 14 & "e" & b + 1
Label6.Caption = a * 6 * 24 & "e" & b + 1
End Sub
