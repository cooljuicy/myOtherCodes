VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7590
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cb 
      Caption         =   "开始"
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox text 
      Height          =   975
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.ListBox list 
      Height          =   5640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cb_Click()
list.Clear
a = text.text
list.AddItem (a)
If a Mod 2 = 0 Then
b = Len(a) / 2
Else
b = Len(a - 1) / 2
End If
b = Int(b)

For i = 1 To b
If Mid(a, i, 1) <> Mid(a, Len(a) + 1 - i, 1) Then
i = 0
Exit For
End If
Next

list.AddItem (i)
list.AddItem (b)
If i - 1 = b Then
list.AddItem ("a")
Else
list.AddItem ("b")
End If

End Sub
