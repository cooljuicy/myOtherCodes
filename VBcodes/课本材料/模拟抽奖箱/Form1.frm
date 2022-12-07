VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5940
   ScaleWidth      =   11505
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton c 
      Caption         =   "开始"
      Height          =   1095
      Left            =   8520
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox t1 
      Height          =   855
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox t1 
      Height          =   855
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox t1 
      Height          =   855
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label daa 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c_Click()
Randomize
daa.Caption = t1(Int(Rnd * 3))
End Sub

Private Sub Form_Load()
daa.Caption = eeeeeeee
For i = 0 To 2
t1(i).FontSize = 30
Next
c.FontSize = 30
daa.FontSize = 30
End Sub
