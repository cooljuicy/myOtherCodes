VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   11070
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yb As Integer
Dim jishu As Integer
Dim beforex As Integer
Dim pjs As Double
Dim shiyancs As Double
Dim maxjs As Integer
Dim dabiaocs As Double
Dim meizucs As Double
Dim yaoqui As Double
Private Sub Form_Load()
shiyancs = InputBox("请输入组数", "请输入组数")
meizucs = InputBox("每组次数", "每组次数")
yaoqui = InputBox("要求连续值", "要求连续值")
dabiaocs = 0
For a = 1 To shiyancs
yb = 2
maxjs = 0
For i = 1 To meizucs
beforex = yb
Randomize
yb = (Int(Rnd * 2))
If beforex = yb Then
jishu = jishu + 1
If jishu > maxjs Then
maxjs = jishu
End If
Else
jishu = 1
End If
Next i
If maxjs >= yaoqui Then
dabiaocs = dabiaocs + 1
pjs = pjs + maxjs
End If
Next a
pjs = pjs / shiyancs
Label1.FontSize = 15
Label1.Caption = "平均最大值：" & pjs & "    " & "达标次数：" & dabiaocs & "    " & "达标比例：" & dabiaocs * 100 / shiyancs & "%"
End Sub
