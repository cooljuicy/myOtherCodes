VERSION 5.00
Begin VB.Form w 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton q 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "w"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
q.Visible = False
End Sub

Private Sub Form_DblClick()
q.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
w.BackColor = vbRed
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
w.BackColor = vbGreen
End Sub

Private Sub Form_Load()
w.BackColor = vbBlue
w.Caption = "捉迷藏"
q.Caption = "单击消失，双击出现"
End Sub
