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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Dim x As Integer, y As Integer
Dim Buffer As Long, hBitmap As Long, Desktop As Long, hScreen As Long, ScreenBuffer As Long
Private Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long)

Private Sub Form_Load()
    Me.Hide

    Desktop = GetWindowDC(GetDesktopWindow())
    hBitmap = CreateCompatibleDC(Desktop)
    hScreen = CreateCompatibleDC(Desktop)
    Buffer = CreateCompatibleBitmap(Desktop, 32, 32)
    ScreenBuffer = CreateCompatibleBitmap(Desktop, Screen.Width / 15, Screen.Height / 15)
    SelectObject hBitmap, Buffer
    SelectObject hScreen, ScreenBuffer
    BitBlt hScreen, 0, 0, Screen.Width / 15, Screen.Height / 15, Desktop, 0, 0, SRCCOPY

    For i = 0 To 1E+17
        y = (Screen.Height / 15) * Rnd
        x = (Screen.Width / 15) * Rnd
        BitBlt hBitmap, 0, 0, 32, 32, Desktop, x, y, SRCCOPY
        BitBlt Desktop, x + (1 - 2 * Rnd), y + (1 - 2 * Rnd), 32, 32, hBitmap, 0, 0, SRCCOPY
        DoEvents
    Next i
End Sub
