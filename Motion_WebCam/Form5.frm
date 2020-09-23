VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones de Conexion"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Establecer Clave de Acceso a Ip's Autorizadadas"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Puerto de Conexión"
      Height          =   855
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   1695
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro IP"
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar seleccion"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   3960
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dirección IP Autorizada"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Activar servidor"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Desactivar servidor"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Reg As CRegSettings
Option Explicit
Dim sPass As String
Private Sub Command1_Click()
New_IP = ""
Form6.Show vbModal, Me
If New_IP <> "0.0.0.0" Then
    ListView1.ListItems.Add , , New_IP
Else
End If
End Sub
Private Sub Command2_Click()
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End Sub
Private Sub Command3_Click()
sPass = InputBox("Introduce la clave de acceso, para las conexiones", "Clave", sPass)
End Sub
Private Sub Command4_Click()
Dim n As Integer
If Option3.Value = True Then
    Reg.SaveSetting "HKEY_CURRENT_USER", "Server", "Active", "0"
Else
    Reg.SaveSetting "HKEY_CURRENT_USER", "Server", "Active", "1"
End If
Reg.SaveSetting "HKEY_CURRENT_USER", "Server", "Port", Text1.Text
Reg.SaveSetting "HKEY_CURRENT_USER", "Server", "Pass", sPass
n = 0
Reg.SaveSetting "HKEY_CURRENT_USER", "Server", "Num_IP", ListView1.ListItems.Count
For n = 1 To ListView1.ListItems.Count
    Reg.SaveSetting "HKEY_CURRENT_USER", "Server", "IP" & n, ListView1.ListItems.Item(n).Text
Next n
If Option3.Value = True Then
    Form1.mnuConection.Checked = False
Else
    Form1.mnuConection.Checked = True
End If
End Sub
Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Form_Load()
Dim nI As Integer, nT As Integer
Set Reg = New CRegSettings
Reg.Company = "TIKO®\Alberto Fedriani"
Reg.AppName = "WebCam Motion"
ListView1.ColumnHeaders.Item(1).Width = 3405
sPass = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Pass", "")
Text1.Text = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Port", "")
nT = Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Num_IP", "0")
For nI = 1 To nT
    ListView1.ListItems.Add , , Reg.GetSetting("HKEY_CURRENT_USER", "Server", "IP" & nI, "")
Next nI
If Reg.GetSetting("HKEY_CURRENT_USER", "Server", "Active", "0") = 0 Then
    Option3.Value = True
Else
    Option1.Value = True
End If
End Sub
Private Sub Option1_Click()
Frame3.Enabled = True
Frame2.Enabled = True
Command3.Enabled = True
End Sub
Private Sub Option3_GotFocus()
Frame3.Enabled = False
Frame2.Enabled = False
Command3.Enabled = False
End Sub
Private Sub Text1_LostFocus()
Dim sInt As Long
On Error GoTo err_Found
sInt = Text1.Text
Text1.Text = sInt
If sInt <= 0 Or sInt > 65535 Then GoTo 1
Winsock1.Close
Winsock1.LocalPort = Text1.Text
Winsock1.Listen
Winsock1.Close
Winsock1.LocalPort = Text1.Text - 1
Winsock1.Listen
Winsock1.Close
Exit Sub
err_Found:
If Err.Number = 13 Then
1    MsgBox "El puerto no es correcto. Los valores del puerto deben estar entre 1 y 65535.", vbExclamation + vbOKOnly, "Error ..."
    Text1.Text = ""
ElseIf Err.Number = 10048 Then
    MsgBox "El puerto ya está ocupado, intenta con otro.", vbExclamation + vbOKOnly, "Error ..."
    Winsock1.Close
    Text1.Text = ""
Else
    Debug.Print Err.Number & "    " & Err.Description
End If
End Sub
