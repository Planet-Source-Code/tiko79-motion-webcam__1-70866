VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WebCap Motion Detector - Acerca de ..."
   ClientHeight    =   1380
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   3780
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3120
   End
   Begin VB.CommandButton Ok_b 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3309
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1800
      Picture         =   "About.frx":0000
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   8
      Left            =   1200
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   600
      Picture         =   "About.frx":030A
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   0
      Picture         =   "About.frx":0BD4
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   3000
      Picture         =   "About.frx":149E
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   2400
      Picture         =   "About.frx":1D68
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   1800
      Picture         =   "About.frx":2632
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "About.frx":2EFC
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "About.frx":37C6
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "About.frx":4090
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "TÏ|{0'oi®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Send me an e-mail !!"
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   90
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":495A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Motion  es una marca registrada de              .  All Right reserved. CopyRight 2007 - 2027."
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Aplicacion para vigilar con tu webcam y tener un registro de todos los movimientos."
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer
Dim toggle As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
toggle = 0
End Sub
Private Sub Label24_Click()
ShellExecute Me.hWnd, "Open", "mailto:bihotz.izoztua@live.fr", "", "", 1
End Sub
Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
toggle = 1
End Sub
Private Sub Ok_b_Click()
On Error GoTo ErrorFound
About.Hide
ErrorFound:
If Err.Number <> 0 Then
Else
End If
End Sub
Private Sub Timer1_Timer()
If toggle = 1 Then Run_Mouse
End Sub
Private Sub Run_Mouse()
x = x + 1: If x = 9 Then x = 0
Label24.MouseIcon = Image2(x).Picture
End Sub
