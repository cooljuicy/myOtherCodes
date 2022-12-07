VERSION 5.00
Begin VB.Form ck 
   Caption         =   "Í¼Æ¬"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton e 
      Caption         =   "ÍË³ö"
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton d 
      Caption         =   "D"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton c 
      Caption         =   "C"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton b 
      Caption         =   "B"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton a 
      Caption         =   "A"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image picshow 
      Height          =   2775
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7560
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "ck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
picshow.Picture = LoadPicture("C:\Users\liuchenyue\Desktop\°®±à³ÌµÄÎÒ\Í¼Æ¬\·ÊÇò.jpg")
End Sub

Private Sub e_Click()
End
End Sub
