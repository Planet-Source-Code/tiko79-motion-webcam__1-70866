VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Cliente 
   Caption         =   "Cliente Imagenes"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   Icon            =   "Cliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Imagen"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock Estate 
      Left            =   4920
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Transfer 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "192.168.7.101"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Puerto"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "No Conectdo"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Estado actual:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Conectar a"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   120
      Picture         =   "Cliente.frx":2532
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer
Private Sub Command1_Click()
On Error GoTo err_found
If Command1.Caption = "Conectar" Then
    Estate.Connect Text1.Text, Text2.Text
    Timer1.Enabled = True
    Command1.Caption = "Desconectar"
Else
    Estate.Close
    Transfer.Close
    If Estate.State <> 7 And Transfer.State <> 7 Then
        Command1.Caption = "Conectar"
        Label3.Caption = "No Conectado"
    End If
End If
err_found:
If Err.Number = 40018 Then
    MsgBox "Revisa los datos !!!!, no se conectará.", vbCritical + vbOKOnly, "Error ..."
Else
'    Debug.Print Err.Number & "      " & Err.Description
End If
End Sub
Private Sub Command2_Click()
SavePicture Image1.Picture, App.Path & "\Cliente_" & Day(Now) & "_" & Month(Now) & "_" & Year(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_" & ".bmp"
End Sub
Private Sub Command3_Click()
Transfer.Close
Estate.Close
Unload Me
End Sub
Private Sub Form_Load()
Dim fso, fFile
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.Path & "\tmp.tmp") = True Then
    Set fFile = fso.GetFile(App.Path & "\tmp.tmp")
    fFile.Delete
    Open App.Path & "\tmp.tmp" For Binary As #1
    Close #1
End If
Timer1.Enabled = False
Label3.Caption = "No Conectado"
n = 0
End Sub
Private Sub Form_Resize()
Label1.Top = Me.Height - 1200
Text2.Top = Me.Height - 1200
Command2.Top = Me.Height - 1200
Label11.Top = Me.Height - 1200
Text1.Top = Me.Height - 1200
Label2.Top = Me.Height - 800
Label3.Top = Me.Height - 800
Command1.Top = Me.Height - 800
Command3.Top = Me.Height - 800
Image1.Height = Me.Height - 1500
Image1.Width = Me.Width - 350
End Sub
Private Sub Image1_DblClick()
If Me.WindowState = 0 Then
    Me.WindowState = 2
Else
    Me.WindowState = 0
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then
   KeyAscii = 0
   Beep
End If
End Sub
Private Sub Text2_LostFocus()
If Text2.Text < 1 Or Text2.Text > 65535 Then
    Text2.Text = ""
    MsgBox "El puerto debe tener un valor entre 1 y 65535.", vbExclamation + vbOKOnly, "Error ..."
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Estate.SendData "Conected?"
End Sub
Private Sub Estate_ConnectionRequest(ByVal requestID As Long)
Estate.Close
Estate.Accept requestID
End Sub
Private Sub Estate_DataArrival(ByVal bytesTotal As Long)
Dim sResp As String, sIn As String
Estate.GetData sResp, vbString
If sResp = "YES" Then
    Label3.Caption = "Conectado"
    If Transfer.State <> 7 Then
        Transfer.Close
        Transfer.Connect Text1.Text, Text2.Text - 1
    Else
    End If
ElseIf sResp = "Pass" Then
    sIn = InputBox("Indica la clave", "Password")
    Estate.SendData "Pass" & sIn
Else
    Label3.Caption = "No Conectado"
End If
End Sub
Private Sub Estate_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Debug.Print Number & " " & Description
If Number = 10053 Then
    Label3.Caption = "No Conectado"
    Estate.Close
    Estate.Connect Text1.Text, Text2.Text
Else
End If
End Sub
Private Sub Transfer_ConnectionRequest(ByVal requestID As Long)
Transfer.Close
Transfer.Accept requestID
End Sub
Private Sub Transfer_DataArrival(ByVal bytesTotal As Long)
Dim data As String, data2 As String, data3 As String
Transfer.GetData data, vbString
'Debug.Print data
data2 = Right(data, 7)
data3 = Left(data, 5)
If data3 = "Start" Then
    Close #1
    Open App.Path & "\tmp.tmp" For Binary As #1
    data3 = ""
    data = Right(data, Len(data) - 5)
End If
Select Case data2
Case "EndFile"
    Put #1, , Left(data, Len(data) - 7)
    Close #1
    On Error GoTo err_found
    Image1.Picture = LoadPicture(App.Path & "\tmp.tmp")
Case Else
    Put #1, , data
End Select
Exit Sub
err_found:
'Debug.Print Err.Number & " " & Err.Description
If Err.Number = 481 Then
    Beep
    Resume Next
Else
    Beep
End If
End Sub
Private Sub Transfer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Debug.Print Number & " " & Description
If Number = 10053 Then
    Transfer.Close
    Transfer.Connect Text1.Text, Text2.Text - 1
Else
End If
End Sub

