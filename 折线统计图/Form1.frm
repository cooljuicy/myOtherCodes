VERSION 5.00
Begin VB.Form biao 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11910
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timer 
      Left            =   2640
      Top             =   1440
   End
   Begin VB.ListBox datalist 
      Height          =   4560
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton choosecb 
      Height          =   1095
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton changecb 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label showdata 
      Height          =   1215
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "biao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inpdata() As Double
Dim answer As Variant
Dim geshu As Double
Dim tsstate As Integer
Dim refill, helpreplace, helpinsert As Double
Dim a, b As Variant
Dim max, min As Double
Dim x1, y1, x2, y2 As Double
Private Sub changecb_Click()
If changecb.Caption = "开始" Then
changecb.Visible = False
datalist.Visible = True
geshu = 0
answer = "abc"
biao.Visible = True
Do Until answer = "ok"
answer = "abc"
Do Until IsNumeric(answer) = True Or answer = "ok"
answer = InputBox("请输入需要统计的数据", "请输入需要统计的数据")
Loop
If Not answer = "ok" Then
geshu = geshu + 1
ReDim Preserve inpdata(1 To geshu)
inpdata(geshu) = answer
datalist.AddItem ("(" & geshu & ")" & answer)
showdata.Caption = answer
End If
Loop
tsstate = 1
showdata.Caption = "开始检查吧！"
choosecb.Visible = True
changecb.FontSize = 10
changecb.Caption = "一个神秘的按钮"
changecb.Visible = True
Else
If tsstate < 6 Then
tsstate = tsstate + 1
Else
tsstate = 1
End If
Select Case tsstate
Case 1
choosecb.Caption = "替换"
Case 2
choosecb.Caption = "插入"
Case 3
choosecb.Caption = "删除"
Case 4
choosecb.Caption = "开始画图"
Case 5
choosecb.Caption = "清屏"
Case 6
choosecb.Caption = "画多条线"
End Select
End If
End Sub

Private Sub choosecb_Click()
If geshu <> 0 Then
answer = geshu + 2
Select Case tsstate

Case 1
Do Until IsNumeric(answer) = True And answer <= geshu And answer > 0
answer = InputBox("你想替换那一项？")
Loop
helpreplace = answer
answer = "abc"
Do Until IsNumeric(answer) = True
answer = InputBox("替换成什么？")
Loop
inpdata(helpreplace) = answer

Case 2
Do Until IsNumeric(answer) = True And answer <= geshu + 1 And answer > 0
answer = InputBox("插入到哪？（小数或较大的数，如：插入到3和4之间,为3和4间的小数，3.1，3.78，4等）？")
Loop
helpinsert = Abs(Int(answer * (-1)))
a = "abc"
Do Until IsNumeric(a) = True
a = InputBox("插入什么？")
Loop
ReDim Preserve inpdata(1 To geshu + 1)
For i = helpinsert To geshu + 1
b = inpdata(i)
inpdata(i) = a
a = b
Next
geshu = geshu + 1

Case 3
Do Until IsNumeric(answer) = True And answer <= geshu And answer > 0
answer = InputBox("删掉哪一项？")
Loop
a = answer
For i = answer To geshu - 1
inpdata(i) = inpdata(i + 1)
Next
ReDim Preserve inpdata(1 To geshu - 1)
geshu = geshu - 1

Case 4
If MsgBox("开始吗", vbYesNo, "折线统计图") = vbYes Then
changecb.Visible = False
choosecb.Visible = False
datalist.Visible = False
showdata.Visible = False
checkMM
biao.Scale (0, 1000)-(1000, 0)
biao.DrawWidth = 10
biao.Line (30, 950)-(30, 60)
biao.Line (30, 60)-(980, 60)
biao.Line (980, 60)-(980, 950)
biao.Line (980, 950)-(30, 950)
timer.Enabled = True
timer.Interval = 0
drawlines
End If

Case 5
If MsgBox("清空所画的所有线条吗", vbYesNo, "折线统计图") = vbYes Then
biao.Cls
End If

Case 6
duoxian
showdata.Caption = "村长还没有研发出这个功能，请静待下一个版本"

End Select
If tsstate = 1 Or 2 Or 3 Then
xuigailiebiao
End If
End If
End Sub

Private Sub Form_Load()
timer.Enabled = False
biao.AutoRedraw = True
choosecb.Visible = False
datalist.Visible = False
changecb.Visible = True
showdata.Visible = True
datalist.Clear
datalist.FontSize = 20
choosecb.FontSize = 20
changecb.FontSize = 20
showdata.FontSize = 20
showdata.ForeColor = RGB(255, 100, 50)
changecb.Caption = "开始"
choosecb.Caption = "替换"
showdata.Caption = "请点击“开始”"
End Sub
Public Sub xuigailiebiao()
datalist.Clear
refill = 0
Do Until refill = geshu
refill = refill + 1
datalist.AddItem ("(" & refill & ")" & inpdata(refill))
Loop
End Sub

Public Sub duoxian()

End Sub

Public Sub checkMM()
max = inpdata(1)
min = inpdata(1)
For i = 1 To geshu
If inpdata(i) > max Then
max = inpdata(i)
End If
If inpdata(i) < min Then
min = inpdata(i)
End If
Next
End Sub

Public Sub drawlines()
a = 910 / (geshu - 1)
b = 810 / (max - min)
For i = 1 To geshu - 1
x1 = 50 + a * (i - 1)
y1 = b * (inpdata(i) - min) + 100
x2 = x1 + a
y2 = b * (inpdata(i + 1) - min) + 100
biao.Line (x1, y1)-(x2, y2)
Next
End Sub
