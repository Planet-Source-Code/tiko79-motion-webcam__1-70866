VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Punteo"
      Height          =   1695
      Left            =   360
      TabIndex        =   12
      Top             =   1800
      Width           =   3615
      Begin VB.OptionButton Option1 
         Caption         =   "Ninguno"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rojo y Verde"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Value           =   -1  'True
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Verde"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rojo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   3135
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         ToolTipText     =   "Tamaño del punteo"
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Label Label3 
         Caption         =   "Tamaño"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mas datos (para ajustar valores)"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Guardar imagen en alarma"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   3127
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aplicar"
      Default         =   -1  'True
      Height          =   375
      Left            =   82
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Intensidad de Rastreo"
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   5
      TickFrequency   =   10
      Value           =   5
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   7680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Color de la Fecha/Hora"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Tolerancia al cambio"
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   200
      SelStart        =   60
      TickFrequency   =   10
      Value           =   60
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Nivel de Alarma"
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   10
      TickFrequency   =   10
      Value           =   10
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Min             =   30
      Max             =   255
      SelStart        =   200
      TickFrequency   =   10
      Value           =   200
   End
   Begin VB.Label Label5 
      Caption         =   "Transparencia"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Nivel de alarma"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Tolerancia al cambio"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Intensidad de Rastreo"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2160
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Reg As CRegSettings
Private Sub Check2_Click()
If Check2.Value = 1 Then
    Me.Width = 7395
    Text1.visible = True
Else
    Me.Width = 4395
    Text1.visible = False
End If
End Sub
Private Sub Command1_Click()
With Dialog1
    .DialogTitle = "Color de la Fecha/Hora"
    .ShowColor
End With
Shape1.FillColor = Dialog1.Color
End Sub
Private Sub Command2_Click()
Dim iOp As Integer
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "LetraColor", Shape1.FillColor
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Intensidad", Slider1.Value
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Tolerancia", Slider2.Value
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Punteo", Slider3.Value
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Alarma", Slider4.Value
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "AutoFrame", Check1.Value
If Option1(0).Value = True Then
    iOp = 0
ElseIf Option1(1).Value = True Then
    iOp = 1
ElseIf Option1(2).Value = True Then
    iOp = 2
ElseIf Option1(3).Value = True Then
    iOp = 3
Else
    iOp = 2
End If
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "iOP", iOp
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "Trans", Slider5.Value
Trans.MakeTransparent Me.hWnd, Slider5.Value
Trans.MakeTransparent Form1.hWnd, Slider5.Value
End Sub
Private Sub Command3_Click()
Check2.Value = vbUnchecked
Me.Width = 7545
Text1.visible = False
Unload Me
End Sub
Private Sub Form_Load()
Dim n As Integer
Set Reg = New CRegSettings
Reg.Company = "TIKO®\Alberto Fedriani"
Reg.AppName = "WebCam Motion"
Shape1.FillColor = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Letracolor", &HFF00&)
Slider1.Value = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Intensidad", "5")
Slider2.Value = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Tolerancia", "60")
Slider3.Value = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Punteo", "5")
Slider4.Value = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Alarma", "10")
Check1.Value = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "AutoFrame", vbUnchecked)
n = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "iOP", "2")
Option1(n).Value = True
Slider5.Value = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "Trans", "200")
Trans.MakeTransparent Me.hWnd, Slider5.Value
Trans.MakeTransparent Form1.hWnd, Slider5.Value
Check2.Value = vbUnchecked
Me.Width = 4470
Text1.visible = False
sForm2 = True
Label1.Caption = "Intensidad de Rastreo ( " & Slider1.Value & " )"
Label2.Caption = "Tolerancia al cambio ( " & Slider2.Value & " )"
Label3.Caption = "Tamaño ( " & Slider3.Value & " )"
Label4.Caption = "Nivel de alarma ( " & Slider4.Value & "% )"
Label5.Caption = "Transparencia ( " & Slider5.Value & " )"
End Sub
Private Sub Form_Unload(Cancel As Integer)
sForm2 = False
End Sub
Private Sub Slider1_Change()
Label1.Caption = "Intensidad de Rastreo ( " & Slider1.Value & " )"
End Sub
Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Intensidad de Rastreo ( " & Slider1.Value & " )"
End Sub
Private Sub Slider2_Change()
Label2.Caption = "Tolerancia al cambio ( " & Slider2.Value & " )"
End Sub
Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Tolerancia al cambio ( " & Slider2.Value & " )"
End Sub
Private Sub Slider3_Change()
Label3.Caption = "Tamaño ( " & Slider3.Value & " )"
End Sub
Private Sub Slider3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = "Tamaño ( " & Slider3.Value & " )"
End Sub
Private Sub Slider4_Change()
Label4.Caption = "Nivel de alarma ( " & Slider4.Value & "% )"
End Sub
Private Sub Slider4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Nivel de alarma ( " & Slider4.Value & "% )"
End Sub
Private Sub Slider5_Change()
Label5.Caption = "Transparencia ( " & Slider5.Value & " )"
End Sub
Private Sub Slider5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Transparencia ( " & Slider5.Value & " )"
End Sub

