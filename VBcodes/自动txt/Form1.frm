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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim filename As String
filename = App.Path & "\" & Text1 & ".txt" '�������ļ���ʽ.
If Dir(filename) = "" Then '�ж��ļ��Ƿ���ڣ������ھʹ��������ھͲ�����
Open filename For Append As #1
Close #1
MsgBox "�����ɹ���", vbInformation, "�ɹ���"
Else
MsgBox "�ļ��Ѵ���.", vbInformation, "�Ѵ������ļ�."
End If
End Sub
