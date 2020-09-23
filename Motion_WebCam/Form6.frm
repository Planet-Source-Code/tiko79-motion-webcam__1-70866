VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Direccion IP"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "000"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "000"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   600
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "000"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "000"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim str As Long
Dim str1 As String
Dim n As Integer
On Error GoTo err_Found
For n = 0 To 3
    str = Text1(n).Text
    Text1(n).Text = str
    If Text1(n).Text < 0 Or Text1(n).Text > 255 Then GoTo 1
Next n
str1 = Text1(0).Text & "." & Text1(1).Text & "." & Text1(2).Text & "." & Text1(3).Text
If str1 <> "0.0.0.0" Then
    New_IP = str1
Else
    MsgBox "Ip nula, no se guardan cambios!", vbExclamation + vbOKOnly, "Error ..."
    New_IP = "0.0.0.0"
End If
Unload Me
err_Found:
If Err.Number = 13 Then
1    MsgBox "La Ip no es correcta. Los valores deben estar entre 0 y 255.", vbExclamation + vbOKOnly, "Error ..."
    Text1(n).Text = ""
    Text1(n).SetFocus
Else
    Debug.Print Err.Number & "    " & Err.Description
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
Text1_GotFocus (0)
End Sub
Private Sub Text1_Change(Index As Integer)
If Index >= 0 And Index <= 2 Then
    If Len(Text1(Index).Text) = 3 Then
        Text1(Index + 1).SetFocus
    Else
    End If
Else
    If Len(Text1(3).Text) = 3 Then
        Command1.SetFocus
    End If
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = 3
End Sub
