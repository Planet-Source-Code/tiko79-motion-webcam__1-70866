VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "WebCam Motion Detector"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1_Con 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   3600
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   6000
      TabIndex        =   7
      ToolTipText     =   "ancho"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "ancho"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   1680
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6525
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Nivel de Movimiento"
            TextSave        =   "Nivel de Movimiento"
            Object.ToolTipText     =   "Nivel de Movimiento"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "Form1.frx":72D8
            Object.ToolTipText     =   "Estado de la deteccion"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Inten"
            TextSave        =   "Inten"
            Object.ToolTipText     =   "Nivel de intensidad de rastreo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Toler"
            TextSave        =   "Toler"
            Object.ToolTipText     =   "Nivel de tolerancia al cambio"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "15:36"
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "10/11/2009"
            Object.ToolTipText     =   "Fecha"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DrawWidth       =   5
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock Estate 
      Left            =   0
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Transfer 
      Left            =   0
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu munOp 
         Caption         =   "&Opciones"
      End
      Begin VB.Menu mnuUP 
         Caption         =   "Siempre &Visible"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNone2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrigen 
         Caption         =   "Ori&gen"
      End
      Begin VB.Menu mnuFormato 
         Caption         =   "&Formato"
      End
      Begin VB.Menu mnuNone0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "Sonido"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMotion 
         Caption         =   "Detectar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuZones 
         Caption         =   "Selector de Zonas"
      End
      Begin VB.Menu mnuVis 
         Caption         =   "Visualizador"
      End
      Begin VB.Menu mnuNone1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConection 
         Caption         =   "Servidor"
      End
      Begin VB.Menu mnuNone3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAb 
         Caption         =   "&Acerca de ..."
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mCapHwnd As Long
Private Reg As CRegSettings
Dim fso As New FileSystemObject
'
Dim POn() As Boolean
'
Dim inten As Integer, i As Integer, j As Integer, R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer, Tolerance As Integer, RealMov As Integer
Dim iRes As Integer, iVal As Integer, n As Integer, nAlarm As Integer
'
Dim P() As Long, Ri As Long, Wo As Long, RealRi As Long, c As Long, c2 As Long, LastTime As Long
'
Dim TppX As Single, TppY As Single
'
Dim sTime As String, sDate As String, str1 As String, sName As String
Private Const chunk = 8000
Dim na As Integer, nn As Integer

Option Explicit
Private Sub Form_Load()
Set Reg = New CRegSettings
Reg.Company = "TIKO®\Alberto Fedriani"
Reg.AppName = "WebCam Motion"
Picture1.Width = 640 * Screen.TwipsPerPixelX
Picture1.Height = 480 * Screen.TwipsPerPixelY
inten = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Intensidad", "5")
Tolerance = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Tolerancia", "60")
Me.Top = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Y", "0")
Me.Left = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "X", "0")
Me.WindowState = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "WindowState", "2")
mnuSound.Checked = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Sound", "True")
mnuMotion.Checked = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Motion", "True")
If mnuMotion.Checked = True Then
    mnuSound.Enabled = True
    mnuZones.Enabled = True
Else
    mnuSound.Enabled = False
    mnuZones.Enabled = False
End If
Trans.MakeTransparent Me.hWnd, Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Trans", "200")
TppX = Screen.TwipsPerPixelX
TppY = Screen.TwipsPerPixelY
ReDim POn(640 / inten, 480 / inten)
ReDim P(640 / inten, 480 / inten)
PBar1.Value = 100
PBar1.Height = 250
PBar1.Width = StatusBar1.Panels(1).Width
sForm2 = False
'
na = nn = 0
If Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Active", "0") = 0 Then
    mnuConection.Checked = False
Else
    mnuConection.Checked = True
End If

'
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Form_Resize
STARTCAM
End Sub
Private Sub Form_Resize()
Image1.Height = Me.Height - StatusBar1.Height - 375
Image1.Width = Me.Width
PBar1.Left = 10
PBar1.Top = Image1.Height + 25
End Sub
Private Sub Form_Unload(Cancel As Integer)
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "WindowState", Me.WindowState
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "X", Me.Left
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Y", Me.Top
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Sound", mnuSound.Checked
Reg.DeleteSetting "HKEY_CURRENT_USER", "Options", "sZone"
End Sub
Private Sub Image1_DblClick()
If Me.WindowState = 0 Then
    Me.WindowState = 2
Else
    Me.WindowState = 0
End If
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuPop
ElseIf Button = 4 Then
    If Label1.visible = False Then
        Label1.visible = True
        Label1.Caption = iVal & " % de movimiento." & vbCrLf & _
                            "Movimiento real (" & RealRi & ")." & vbCrLf & _
                            "Intensidad de rastreo(" & inten & "), tolerancia (" & Tolerance & ")." & vbCrLf & _
                            "Fecha: " & sDate & " Hora: " & sTime
        Label1.Top = Y
        Label1.Left = X
        Timer2.Enabled = True
    Else
        Label1.visible = False
    End If
Else
End If
End Sub
Private Sub mnuAb_Click()
Me.WindowState = 0
mnuUP.Checked = False
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
About.Show vbModal, Me
End Sub
Private Sub MnuConection_Click()
Me.WindowState = 0
mnuUP.Checked = False
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Form5.Show vbModal, Me
End Sub
Private Sub mnuFormato_Click()
Call capDlgVideoFormat(mCapHwnd)
End Sub
Private Sub mnuMotion_Click()
If mnuMotion.Checked = True Then
    mnuMotion.Checked = False
    mnuSound.Enabled = False
    mnuZones.Enabled = False
Else
    mnuMotion.Checked = True
    mnuSound.Enabled = True
    mnuZones.Enabled = True
End If
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Motion", mnuMotion.Checked
End Sub
Private Sub mnuOrigen_Click()
Call capDlgVideoSource(mCapHwnd)
End Sub
Private Sub mnuSalir_Click()
STOPCAM
Unload Me
End Sub
Private Sub mnuSound_Click()
If mnuSound.Checked = True Then
    mnuSound.Checked = False
Else
    mnuSound.Checked = True
End If
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Sound", mnuSound.Checked
End Sub
Private Sub mnuUP_Click()
If mnuUP.Checked = True Then
    mnuUP.Checked = False
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Else
    mnuUP.Checked = True
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub
Private Sub mnuVis_Click()
On Error Resume Next
Me.WindowState = 0
mnuUP.Checked = False
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Form3.Show , Me
End Sub
Private Sub mnuZones_Click()
Me.WindowState = 0
mnuUP.Checked = False
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Form4.Show vbModal, Me
End Sub
Private Sub munOp_Click()
Me.WindowState = 0
mnuUP.Checked = False
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Form2.Show vbModal, Me
End Sub
Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Panel.Index = 2 Then
    If Timer1.Enabled = False Then
        STARTCAM
    Else
        STOPCAM
    End If
End If
End Sub
Private Sub Timer1_Timer()
On Error GoTo err_Found
If mnuConection.Checked = False Then
    Timer1_Con.Enabled = False
    Estate.Close

Else
    If Timer1_Con.Enabled = True Then
    Else
        Timer1_Con.Enabled = True
        Estate.LocalPort = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Port", "9999")
        Estate.Listen
    End If
End If
nAlarm = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Alarma", "10")
n = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "iOP", "2")
str1 = Left(Reg.GetSetting("HKEY_CURRENT_USER", "Options", "sZone", "00"), 1)
inten = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Intensidad", "5")
Tolerance = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Tolerancia", "60")
StatusBar1.Panels(3).Text = " Intensidad = " & inten
StatusBar1.Panels(4).Text = " Tolerancia = " & Tolerance
StatusBar1.Panels(5).Text = Time
StatusBar1.Panels(6).Text = Date
Picture1.DrawWidth = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Punteo", "5")
SendMessage mCapHwnd, GET_FRAME, 0, 0
SendMessage mCapHwnd, COPY, 0, 0
Picture1.picture = Clipboard.GetData(2)
Clipboard.Clear
PBar1.Value = 0
sTime = Time
sDate = Date
With Picture1
    .FontBold = False
    .FontItalic = False
    .FontStrikethru = False
    .FontUnderline = False
    .ForeColor = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Letracolor", &HFF00&)
End With
If Picture1.Width / Screen.TwipsPerPixelX = 644 Then
    Picture1.FontSize = 7
ElseIf Picture1.Width / Screen.TwipsPerPixelX = 356 Then
    Picture1.FontSize = 7
ElseIf Picture1.Width / Screen.TwipsPerPixelX = 324 Then
    Picture1.FontSize = 7
ElseIf Picture1.Width / Screen.TwipsPerPixelX = 180 Then
    Picture1.FontSize = 6
Else
    Picture1.FontSize = 5
End If
Picture1.CurrentX = 50
Picture1.CurrentY = 50
Picture1.Print sTime
Picture1.CurrentX = 50
Picture1.CurrentY = 200
Picture1.Print sDate
If mnuMotion.Checked = True Then
    If Reg.GetSetting("HKEY_CURRENT_USER", "Options", "AutoFrame", vbUnchecked) = vbChecked Then
        StatusBar1.Panels(2).picture = ImageList1.ListImages.Item(1).picture
    Else
        StatusBar1.Panels(2).picture = ImageList1.ListImages.Item(2).picture
    End If
    If str1 = "1" Then
        If List1.ListCount = m Then 'todas seleccionadas
            Motion
        Else
            MotionB
        End If
    Else
        Motion
    End If
Else
    StatusBar1.Panels(2).picture = ImageList1.ListImages.Item(3).picture
End If
Image1 = Picture1.Image
err_Found:
If Err.Number = 521 Then
    Beep
    Resume Next
ElseIf Err.Number = 0 Then
Else
    Beep
    iRes = MsgBox(Err.Description & vbCr & "Presiones SI para ignorar, NO para salir", vbInformation + vbYesNo, "Error Nº= " & Err.Number)
    If iRes = 6 Then
        Resume Next
    Else
        STOPCAM
        Unload Me
    End If
End If
End Sub
Private Sub Motion()
On Error GoTo err_Found
Dim nMax As Integer
Dim iTPPX As Integer, iTPPY As Integer, iTPPX1 As Integer, iTPPY1 As Integer, iTPPX2  As Integer, iTPPY2 As Integer
Ri = 0
Wo = 0
For i = 0 To 640 / inten - 1
    For j = 0 To 480 / inten - 1
        PartA
    Next j
Next i
RealRi = 0
For i = 1 To 640 / inten - 2
    For j = 1 To 480 / inten - 2
        PartB
    Next j
Next i
PartC
err_Found:
If Err.Number = 0 Then
ElseIf Err.Number = 380 Then 'imagen inicial
    Beep
    Resume Next
ElseIf Err.Number = 6 Then
    Resume Next
Else
    Beep
    iRes = MsgBox(Err.Description & vbCr & "Presiones SI para ignorar, NO para salir", vbInformation + vbYesNo, "Error Nº= " & Err.Number)
    If iRes = 6 Then
        Resume Next
    Else
        STOPCAM
        Unload Me
        End
    End If
End If
End Sub
Private Sub PartA()
c = Picture1.Point(i * inten * Screen.TwipsPerPixelX, j * inten * Screen.TwipsPerPixelY)
c2 = P(i, j)
If Different(c, c2) Then
    P(i, j) = Picture1.Point(i * inten * Screen.TwipsPerPixelX, j * inten * Screen.TwipsPerPixelY)
    If n = 0 Or n = 2 Then Picture1.PSet (i * inten * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY * j * inten), vbRed
    Wo = Wo + 1
    POn(i, j) = False
Else
    POn(i, j) = True
    Ri = Ri + 1
End If
End Sub
Private Sub PartB()
If POn(i, j) = False Then
    If POn(i, j + 1) = False Then
        If POn(i, j - 1) = False Then
            If POn(i + 1, j) = False Then
                If POn(i - 1, j) = False Then
                    If n = 1 Or n = 2 Then Picture1.PSet (i * inten * TppX, j * inten * TppY), vbGreen
                End If
            End If
        End If
    End If
End If
End Sub
Private Sub PartC()
'aki alarma
iVal = Int(Wo / (Ri + Wo) * 100)
If sForm2 = True Then
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & " iVal" & iVal & " .... Real Mov " & RealRi
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
Else
End If
PBar1.Value = iVal
If iVal >= nAlarm Then
        With Picture1
            .FontBold = False
            .FontItalic = False
            .FontStrikethru = False
            .FontUnderline = False
            .ForeColor = vbRed
            .FontSize = 7
        End With
        Picture1.CurrentX = Picture1.Width - 1000
        Picture1.CurrentY = 50
        Picture1.Print "iVal= " & iVal
        If mnuSound.Checked = True Then PlayWaveRes "ALARMA"
        If fso.FolderExists(App.path & "\Recorded") = False Then fso.CreateFolder App.path & "\Recorded"
        If Reg.GetSetting("HKEY_CURRENT_USER", "Options", "AutoFrame", vbUnchecked) = vbChecked Then
            sName = App.path & "\Recorded\img" & "_" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & "_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date) & ".bmp"
            SavePicture Image1.picture, sName
            If fso.FileExists(sName) = True Then
                Open sName For Append As #1
                    Print #1, "##MARCADEAGUA##"
                    Print #1, "#NombreFichero: " & sName
                    Print #1, "#Dia: " & Date
                    Print #1, "#Hora: " & Time
                    Print #1, "#App Name: " & App.EXEName
                    Print #1, "#App Version: " & App.Major & "." & App.Minor & "." & App.Revision
                    Print #1, "#Intensidad: " & inten
                    Print #1, "#Tolerancia: " & Tolerance
                    Print #1, "#Alarma: " & nAlarm
                    Print #1, "#iVal: " & iVal
                    Print #1, "#Tamaño Fichero: " & LOF(1) + 25
                Close #1
            End If
        Else
        End If
Else
End If
End Sub
Private Sub MotionB()
On Error GoTo err_Found
Dim nMax As Integer
Dim iTPPX As Integer, iTPPY As Integer, iTPPX1 As Integer, iTPPY1 As Integer, iTPPX2  As Integer, iTPPY2 As Integer
Ri = 0
Wo = 0
If List1.ListCount <= 0 Then Exit Sub
For i = 0 To 640 / inten - 1
    iTPPX = Val(i * inten * TppX)
    For nMax = 1 To List1.ListCount
        iTPPX1 = Val(List3.List(nMax - 1))
        iTPPX2 = iTPPX1 + 600
        If iTPPX >= iTPPX1 And iTPPX <= iTPPX2 Then
            For j = 0 To 480 / inten - 1
                iTPPY = Val(j * inten * TppY)
                    iTPPY1 = Val(List1.List(nMax - 1))
                    iTPPY2 = iTPPY1 + 600
                    If iTPPY >= iTPPY1 And iTPPY <= iTPPY2 Then
                            PartA
                    End If
            Next j
        End If
    Next nMax
Next i
RealRi = 0
For i = 1 To 640 / inten - 2
    iTPPX = Val(i * inten * TppX)
    For nMax = 1 To List1.ListCount
        iTPPX1 = Val(List3.List(nMax - 1))
        iTPPX2 = iTPPX1 + 600
        If iTPPX >= iTPPX1 And iTPPX <= iTPPX2 Then
            For j = 1 To 480 / inten - 2
                iTPPY = Val(j * inten * TppY)
                    iTPPY1 = Val(List1.List(nMax - 1))
                    iTPPY2 = iTPPY1 + 600
                    If iTPPY >= iTPPY1 And iTPPY <= iTPPY2 Then
                            PartB
                    End If
            Next j
        End If
    Next nMax
Next i
PartC
err_Found:
If Err.Number = 0 Then
ElseIf Err.Number = 380 Then 'imagen inicial
    Beep
    Resume Next
ElseIf Err.Number = 6 Then
    Resume Next
Else
    Beep
    iRes = MsgBox(Err.Description & vbCr & "Presiones SI para ignorar, NO para salir", vbInformation + vbYesNo, "Error Nº= " & Err.Number)
    If iRes = 6 Then
        Resume Next
    Else
        STOPCAM
        Unload Me
        End
    End If
End If
End Sub
Private Function Different(ByVal c As Long, ByVal c1 As Long) As Boolean
R = c Mod 256
G = (c \ 256) Mod 256
B = (c \ 256 \ 256) Mod 256
R2 = c2 Mod 256
G2 = (c2 \ 256) Mod 256
B2 = (c2 \ 256 \ 256) Mod 256
Different = (Sqr((R - R2) * (R - R2) + (G - G2) * (G - G2) + (B - B2) * (B - B2)) > Tolerance)
End Function
Sub STOPCAM()
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
Timer1.Enabled = False
StatusBar1.Panels(2).picture = ImageList1.ListImages.Item(3).picture
End Sub
Sub STARTCAM()
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hWnd, 0)
DoEvents
SendMessage mCapHwnd, CONNECT, 0, 0
Timer1.Enabled = True
End Sub
Private Sub Timer2_Timer()
Label1.visible = False
Timer2.Enabled = False
End Sub
Private Sub Timer1_Con_Timer()
Dim Data As String
If Transfer.State = 7 Then
    na = 0
    nn = 0
    SavePicture Image1.picture, App.path & "\tmp.tmp"
    Transfer.SendData "Start"
    Close #1
    Open App.path & "\tmp.tmp" For Binary As #1
        Do While Not EOF(1)
            Data = Input(chunk, #1)
            On Error GoTo err_Found
            Transfer.SendData Data
            DoEvents
        Loop
        Close #1
        Transfer.SendData "EndFile"
Else
1    If na = 0 Then
        Estate.Close
        Estate.LocalPort = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Port", "9999")
        Estate.Listen
        Transfer.Close
        Transfer.LocalPort = Estate.LocalPort - 1
        Transfer.Listen
        na = 1
    Else
        nn = nn + 1
        If nn = 1000 Then
            nn = 0
            na = 0
        End If
    End If
End If
Exit Sub
err_Found:
If Err.Number = 400006 Then
    Estate.Close
    Transfer.Close
    n = 0
    GoTo 1
Else
    Resume Next
End If
End Sub
Private Sub Transfer_ConnectionRequest(ByVal requestID As Long)
    Transfer.Close
    Transfer.Accept requestID
End Sub
Private Sub Estate_ConnectionRequest(ByVal requestID As Long)
Dim nT As Integer, nI As Integer
nT = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Num_IP", "0")
For nI = 1 To nT
    If Estate.RemoteHostIP = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "IP" & nI, "") Then
        Estate.Close
        Estate.Accept requestID
        Estate.SendData "Pass"
        Exit Sub
    Else
    End If
Next nI
End Sub
Private Sub Estate_DataArrival(ByVal bytesTotal As Long)
Dim sResp As String
Estate.GetData sResp, vbString
If Left(sResp, 4) = "Pass" Then
    If Right(sResp, Len(sResp) - 4) = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Pass", "") Then
        Transfer.Close
        Transfer.LocalPort = Estate.LocalPort - 1
        Transfer.Listen
        Timer1_Con.Enabled = True
    Else
        Transfer.Close
        Transfer.LocalPort = Estate.LocalPort - 1
        Transfer.Listen
        Estate.Close
        Estate.LocalPort = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Port", "9999")
        Estate.Listen
    End If
Else
End If
If sResp = "Conected?" Then
    Estate.SendData "YES"
Else
End If
End Sub
Private Sub Estate_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = 10053 Then
    Estate.Close
    Estate.LocalPort = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Port", "9999")
    Estate.Listen
Else
End If
End Sub
Private Sub Transfer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = 10053 Then
    Transfer.Close
    Transfer.LocalPort = Estate.LocalPort - 1
    Transfer.Listen
Else
End If
End Sub

