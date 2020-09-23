Attribute VB_Name = "Module2"
Option Explicit
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public Const CONNECT As Long = 1034
Public Const DISCONNECT As Long = 1035
Public Const GET_FRAME As Long = 1084
Public Const COPY As Long = 1054
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public sForm2 As Boolean
Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Public New_IP As String
Public m As Integer
Function capDlgVideoFormat(ByVal hCapWnd As Long) As Boolean
   capDlgVideoFormat = SendMessage(hCapWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
End Function
Function capDlgVideoSource(ByVal hCapWnd As Long) As Boolean
   capDlgVideoSource = SendMessage(hCapWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Function

