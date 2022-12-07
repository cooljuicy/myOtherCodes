VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7995
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton start 
      Caption         =   "GO"
      Height          =   975
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox text 
      Height          =   1095
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.ListBox list 
      Height          =   5640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Variant
Private Sub Form_Load()
start.FontSize = 35
text.FontSize = 35
list.FontSize = 25
text.ForeColor = RGB(255, 0, 0)
list.ForeColor = RGB(0, 0, 255)
text.text = ""
list.Clear
End Sub

Private Sub start_Click()
list.Clear
a = text.text
b = text.text
Do Until a - (2 * Fix(a / 2)) <> 0
a = a / 2
list.AddItem (2)
Loop

For i = 3 To Sqr(b) Step 2

If a = 1 Then
Exit For
End If

For j = 2 To Sqr(i) Step 2
If i Mod j = 0 Then
Exit For
End If
DoEvents
Next

If j > Sqr(i) Then
Do Until a - (i * Fix(a / i)) <> 0
a = a / i
list.AddItem (i)
Loop

For j = 2 To Sqr(a) Step 2
If a - (j * Fix(a / j)) = 0 Then
Exit For
End If
DoEvents
Next
If j > Sqr(a) And a <> 1 Then
list.AddItem (a)
Exit For
End If
End If

DoEvents
Next













If list.ListCount = 0 Then
list.AddItem (b)
End If
End Sub
