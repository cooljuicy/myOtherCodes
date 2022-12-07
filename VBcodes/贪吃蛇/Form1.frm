VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Timer1 As Timer
Attribute Timer1.VB_VarHelpID = -1
Private WithEvents Label1 As Label
Attribute Label1.VB_VarHelpID = -1
Dim GFangXiang As Boolean
Dim HWB As Single
Dim She() As ShenTi
Dim X As Long, Y As Long
Dim ZhuangTai(23, 23) As Long

Private Type ShenTi
F As Long
X As Long
Y As Long
End Type

'按键反应 ←↑↓→
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim C As Long
If KeyCode = 27 Then End 'ESC退出
If KeyCode = 32 Then
If Timer1.Enabled = True Then '空格暂停
Timer1.Enabled = False
Label1.Visible = True
Else '空格开始
Timer1.Enabled = True
Label1.Visible = False
End If
End If
C = UBound(She)
If GFangXiang = True Then Exit Sub
Select Case KeyCode
Case 37 '←
If She(C).F = 2 Then Exit Sub
She(C).F = 0
GFangXiang = True
Case 38 '↑
If She(C).F = 3 Then Exit Sub
She(C).F = 1
GFangXiang = True
Case 39 '↑
If She(C).F = 0 Then Exit Sub
She(C).F = 2
GFangXiang = True
Case 40 '→
If She(C).F = 1 Then Exit Sub
She(C).F = 3
GFangXiang = True
End Select
End Sub

Private Sub Form_Load()
Me.AutoRedraw = True
Me.BackColor = &HC000&
Me.FillColor = 255
Me.FillStyle = 0
Me.WindowState = 2
Set Timer1 = Controls.Add("VB.Timer", "Timer1")
Set Label1 = Controls.Add("VB.Label", "Label1")
Label1.AutoSize = True
Label1.BackStyle = 0
Label1 = "暂停"
Label1.ForeColor = RGB(255, 255, 0)
Label1.FontSize = 50
ChuShiHua '初始化
End Sub

Private Sub Form_Resize()
On Error GoTo 1:
With Me
If .WindowState <> 1 Then
.Cls
.ScaleMode = 3
HWB = .ScaleHeight / .ScaleWidth
.ScaleWidth = 24
.ScaleHeight = 24
Label1.Move (Me.ScaleWidth - Label1.Width) / 2, (Me.ScaleHeight - Label1.Height) / 2
HuaTu
Me.Line (X, Y)-(X + 1, Y + 1), RGB(255, 255, 0), BF
End If
End With
1:
End Sub

Private Sub Timer1_Timer()
Dim C As Long, I As Long
On Error GoTo 2:
QingChu '清图
C = UBound(She)
Select Case She(C).F
Case 0
If ZhuangTai(She(C).X - 1, She(C).Y) = 2 Then
C = C + 1
ReDim Preserve She(C)
She(C).F = She(C - 1).F
She(C).X = She(C - 1).X - 1
She(C).Y = She(C - 1).Y
ChanShengShiWu
GoTo 1:
ElseIf ZhuangTai(She(C).X - 1, She(C).Y) = 1 Then
GoTo 2:
End If
Case 1
If ZhuangTai(She(C).X, She(C).Y - 1) = 2 Then
C = C + 1
ReDim Preserve She(C)
She(C).F = She(C - 1).F
She(C).X = She(C - 1).X
She(C).Y = She(C - 1).Y - 1
ChanShengShiWu
GoTo 1:
ElseIf ZhuangTai(She(C).X, She(C).Y - 1) = 1 Then
GoTo 2:
End If
Case 2
If ZhuangTai(She(C).X + 1, She(C).Y) = 2 Then
C = C + 1
ReDim Preserve She(C)
She(C).F = She(C - 1).F
She(C).X = She(C - 1).X + 1
She(C).Y = She(C - 1).Y
ChanShengShiWu
GoTo 1:
ElseIf ZhuangTai(She(C).X + 1, She(C).Y) = 1 Then
GoTo 2:
End If
Case 3
If ZhuangTai(She(C).X, She(C).Y + 1) = 2 Then
C = C + 1
ReDim Preserve She(C)
She(C).F = She(C - 1).F
She(C).X = She(C - 1).X
She(C).Y = She(C - 1).Y + 1
ChanShengShiWu
GoTo 1:
ElseIf ZhuangTai(She(C).X, She(C).Y + 1) = 1 Then
GoTo 2:
End If
End Select
ZhuangTai(She(0).X, She(0).Y) = 0
For I = 0 To C
Select Case She(I).F
Case 0
She(I).X = She(I).X - 1
Case 1
She(I).Y = She(I).Y - 1
Case 2
She(I).X = She(I).X + 1
Case 3
She(I).Y = She(I).Y + 1
End Select
Next
TiaoZheng
1:
GFangXiang = False
ZhuangTai(She(C).X, She(C).Y) = 1
HuaTu
Exit Sub
2: '游戏结束
If MsgBox("得分：" & UBound(She) - 2 & "分 " & vbCrLf & "游戏结束，点“是”重新开始游戏，点“否”", vbYesNo, "贪吃蛇") = vbYes Then
ChuShiHua
Else
End
End If
End Sub

'初始化
Private Sub ChuShiHua()
Me.Cls
Timer1.Enabled = True
Timer1.Interval = 50
Erase ZhuangTai
ReDim She(2)
She(0).F = 2
She(0).X = 9
She(0).Y = 11
ZhuangTai(9, 11) = 1
She(1).F = 2
She(1).X = 10
She(1).Y = 11
ZhuangTai(10, 11) = 1
She(2).F = 2
She(2).X = 11
She(2).Y = 11
ZhuangTai(11, 11) = 1
HuaTu '画图
ChanShengShiWu
End Sub

'清图
Private Sub QingChu()
Dim I As Long
For I = 0 To UBound(She)
Me.Line (She(I).X, She(I).Y)-(She(I).X + 1, She(I).Y + 1), Me.BackColor, BF
Next
End Sub

'画图 蛇
Private Sub HuaTu()
Dim I As Long
For I = 0 To UBound(She)
Me.Circle (She(I).X + 0.5, She(I).Y + 0.5), 0.49, RGB(255, 255, 0), , , HWB
Next
End Sub

Private Sub TiaoZheng()
Dim I As Long
For I = 0 To UBound(She) - 1
She(I).F = She(I + 1).F
Next
End Sub

'随机产生食物
Private Sub ChanShengShiWu()
Randomize Timer
1:
X = Int(Rnd * 24)
Y = Int(Rnd * 24)
If ZhuangTai(X, Y) > 0 Then GoTo 1:
ZhuangTai(X, Y) = 2
Me.Line (X, Y)-(X + 1, Y + 1), RGB(255, 255, 0), BF
End Sub
