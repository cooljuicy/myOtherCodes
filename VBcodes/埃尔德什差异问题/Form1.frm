VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   11445
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbcayi As Single
Dim dqcayi As Single
Dim i As Single
Private Sub Form_Load()
mbcayi = InputBox("请输入需要的差异")
dqcayi = 0
i = 0
Do Until Abs(dqcayi) = mbcayi
Randomize
If Int(Rnd * 2) = 0 Then
dqcayi = dqcayi + 1
Else
dqcayi = dqcayi - 1
End If
i = i + 1
Loop
Label1.FontSize = 30
Label1.Caption = "已用步数：" & i & "    和理论偏差" & Abs(i - mbcayi ^ 2) * 100 / mbcayi ^ 2 & "%"
End Sub
