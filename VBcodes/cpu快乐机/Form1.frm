VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4260
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim geshu, x As Single
Dim a As Single
a = 10000000
geshu = 1
x = 3
Do Until geshu = a
For i = 3 To x Step 2
If x Mod i = 0 Then
Exit For
End If
DoEvents
Next
If i = x Then
geshu = geshu + 1
End If
x = x + 2
DoEvents
Loop
End Sub
