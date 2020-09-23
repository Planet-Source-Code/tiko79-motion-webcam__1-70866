Attribute VB_Name = "Trans"
Option Explicit
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 30 Or Perc > 255 Then
  MakeTransparent = 30
Else
  Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 30
End If
End Function





