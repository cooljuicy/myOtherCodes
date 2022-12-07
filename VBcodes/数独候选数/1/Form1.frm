VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   11625
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox list 
      Height          =   6360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zb(1 To 9, 1 To 9) As String
Dim ggesz(1 To 9, 1 To 9) As String
Dim gge As String
Dim an As String
Dim now As Integer
Dim canpc As Boolean
Dim addtx, addtd As String
Dim hanggs As Integer
Dim aj, ai As Integer

Private Sub Form_Load()
chushi
tianjia

hangges = 1
addtd = ""
For j = 9 To 1 Step -1 '1
For i = 1 To 9 Step 1
aj = j
ai = i
addtx = ""
If zb(i, j) = 0 Then
For n = 1 To 9 '2
canpc = False
checkpaichu

If canpc = False Then
addtx = addtx & n
End If

Next '2

addtx = addtx & Space(9 - Len(addtx))
gaihanggs

Else
addtx = Space(9)
gaihanggs
End If
list.AddItem (addtx)
Next '1
Next
End Sub

Public Sub chushi()
Form1.Show
list.FontSize = 30
an = ""
gge = ""
For i = 1 To 3
gge = gge & "111222333"
Next
For i = 1 To 3
gge = gge & "444555666"
Next
For i = 1 To 3
gge = gge & "777888999"
Next
now = 1
For j = 9 To 1 Step -1
For i = 1 To 9 Step 1
ggesz(i, j) = Mid(gge, now, 1)
now = now + 1
Next
Next
End Sub

Public Sub tianjia()
Do Until Len(an) = 81 And IsNumeric(an) = True
an = InputBox("输入数独")
Loop
now = 1
For j = 9 To 1 Step -1
For i = 1 To 9 Step 1
zb(i, j) = Mid(an, now, 1)
now = now + 1
Next
Next
End Sub

Public Sub checkpaichu()
For m = 1 To 9
If zb(m, aj) = n Then
canpc = True
End If
Next
For m = 1 To 9
If zb(ai, m) = n Then
canpc = True
End If
Next
For a = 1 To 9
For b = 1 To 9
If ggesz(a, b) = ggesz(ai, aj) And zb(a, b) = n Then
canpc = True
End If
Next
Next
End Sub

Public Sub gaihanggs()
If hanggs <> 3 Then
addtd = addtd & addtx & "#"
hanggs = hanggs + 1
Else
addtd = addtd & addtx
hanggs = 1
End If
End Sub
