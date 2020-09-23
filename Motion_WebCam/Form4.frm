VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selector de zona"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12645
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List5 
      Height          =   5130
      Left            =   3960
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   5130
      Left            =   960
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   5130
      Left            =   2040
      TabIndex        =   10
      ToolTipText     =   "ancho"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List4 
      Height          =   5130
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "ancho"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   7275
      ItemData        =   "Form4.frx":0000
      Left            =   9720
      List            =   "Form4.frx":0002
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seleccionar Todo"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deseleccionar Todo"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   480
      Top             =   0
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mas Datos"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Activar Funcion"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   7300
      Left            =   50
      ScaleHeight     =   7245
      ScaleWidth      =   9540
      TabIndex        =   0
      Top             =   50
      Width           =   9600
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   855
         Index           =   0
         Left            =   8520
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Reg As CRegSettings
Dim Columnas As Integer, Filas As Integer, n As Integer, i As Integer
Dim STR As String
Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    Check1.Caption = "Desactivar Funcion"
Else
    Check1.Caption = "Activar Funcion"
    Form1.List1.Clear
    Form1.List2.Clear
    Form1.List3.Clear
    Form1.List4.Clear
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    Me.Width = 12735
Else
    Me.Width = 9825
End If
End Sub
Private Sub Command1_Click()
Dim n As Integer
For n = 1 To Shape1.Count - 1
    Shape1(n).FillStyle = 7
    List1.RemoveItem n - 1
    List1.AddItem n - 1 & " Deselec", n - 1
Next n
End Sub
Private Sub Command2_Click()
Dim n As Integer
For n = 1 To Shape1.Count - 1
    Shape1(n).FillStyle = 1
    List1.RemoveItem n - 1
    List1.AddItem "Fila :" & (Shape1(n).Top / 600) + 1 & " Columna: " & (Shape1(n).Left / 600) + 1, n - 1
Next n
End Sub
Private Sub Command3_Click()
Dim sA As String, sB As String, sC As String
Dim iTemp As String
Form1.List1.Clear
Form1.List2.Clear
Form1.List3.Clear
Form1.List4.Clear
'#,##.##,##.##, . . .  ...
'A,BB.CC,DD.EE.F, . . .  ...
'A= 1(Activo); 0(Inactivo) la Funcion Zone
'BB= Nº Total de Filas
'CC= Nº Total de Columnas
'DD= Nº Fila
'EE= Nº Columna
'F= 0 - 1 para esa celda
STR = ""
For n = 1 To m - 1
    If Shape1(n).FillStyle = 1 Then 'trans (activo)
        sA = "01"
    Else
        sA = "00"
    End If
    iTemp = (Shape1(n).Top / 600) + 1
    If Len(iTemp) = 1 Then
        sB = "0" & iTemp
    Else
        sB = iTemp
    End If
    iTemp = (Shape1(n).Left / 600) + 1
    If Len(iTemp) = 1 Then
        sC = "0" & iTemp
    Else
        sC = iTemp
    End If
    STR = STR & sB & "." & sC & "." & sA & ","
Next n
STR = Check1.Value & "," & Left(STR, Len(STR) - 1)
Reg.SaveSetting "HKEY_CURRENT_USER", "Options", "sZone", STR
For n = 1 To List1.ListCount
    If Right(List1.List(n - 1), 7) = "Deselec" Then
    Else
        List2.AddItem (Shape1(n).Top)
        List3.AddItem (Shape1(n).Top + 600)
        List4.AddItem (Shape1(n).Left)
        List5.AddItem (Shape1(n).Left + 600)

        Form1.List1.AddItem (Shape1(n).Top)
        Form1.List2.AddItem (Shape1(n).Top + 600)
        Form1.List3.AddItem (Shape1(n).Left)
        Form1.List4.AddItem (Shape1(n).Left + 600)
    End If
Next n
m = 0
m = List1.ListCount
Unload Me
End Sub
Private Sub Command4_Click()
Unload Me
End Sub
Private Sub Form_DblClick()
Dim filas1 As String, columnas1 As String, STR1 As String
STR = Reg.GetSetting("HKEY_CURRENT_USER", "Options", "sZone", "00")
If Left(STR, 1) = "0" Then
    Command1_Click
Else
    filas1 = Mid(STR, Len(STR) - 7, 2)
    columnas1 = Mid(STR, Len(STR) - 4, 2)
    m = Filas * Columnas
    i = 3
    For n = 1 To m
        STR1 = Mid(STR, (7 * n) + i, 1)
        If Mid(STR, (7 * n) + i, 1) = "0" Then
            Shape1(n).FillStyle = 7
            List1.RemoveItem n - 1
            List1.AddItem n - 1 & " Deselec", n - 1
        Else
            Shape1(n).FillStyle = 1
            List1.RemoveItem n - 1
            List1.AddItem "Fila :" & (Shape1(n).Top / 600) + 1 & " Columna: " & (Shape1(n).Left / 600) + 1, n - 1
        End If
        i = i + 2
    Next n
End If
End Sub
Private Sub Form_Load()
Set Reg = New CRegSettings
Reg.Company = "TIKO®\Alberto Fedriani"
Reg.AppName = "WebCam Motion"
Picture1.Height = Form1.Picture1.Height
Picture1.Width = Form1.Picture1.Width
Me.Width = 9825
Me.Caption = "Selector de Zonas. Resolucion : " & (Picture1.Width / Screen.TwipsPerPixelX) - 4 & " x " & (Picture1.Height / Screen.TwipsPerPixelY) - 4
m = 0
Fill_All
Form_Resize
End Sub
Private Sub Form_Resize()
Picture1.Top = (Check1.Top / 2) - (Picture1.Height / 2) - 50
Picture1.Left = (9825 / 2) - (Picture1.Width / 2) - 50
End Sub
Private Sub Timer1_Timer()
Picture1.picture = Form1.Image1.picture
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1(Index).FillStyle = 7 Then
    Label1(Index).ToolTipText = "Area NO Seleccionada"
Else
    Label1(Index).ToolTipText = "Area SI Seleccionada"
End If
End Sub
Private Sub Label1_Click(Index As Integer)
If Shape1(Index).FillStyle = 7 Then
    Shape1(Index).FillStyle = 1
    List1.RemoveItem Index - 1
    List1.AddItem "Fila :" & (Shape1(Index).Top / 600) + 1 & " Columna: " & (Shape1(Index).Left / 600) + 1, Index - 1
        List2.AddItem (Shape1(Index).Top)
        List3.AddItem (Shape1(Index).Top + 600)
        List4.AddItem (Shape1(Index).Left)
        List5.AddItem (Shape1(Index).Left + 600)
Else
    Shape1(Index).FillStyle = 7
    List1.RemoveItem Index - 1
    List1.AddItem Index - 1 & " Deselec", Index - 1
End If
End Sub
Private Function Fill_All()
Columnas = Picture1.Width / 600
Filas = Picture1.Height / 600
m = 1
List1.Clear
With Shape1(0)
    .Height = 600
    .Width = 600
    .Top = 0
    .Left = 0
    .FillStyle = 7
    .DrawMode = 13
End With
With Label1(0)
    .Height = 600
    .Width = 600
    .Top = 0
    .Left = 0
End With
For n = 0 To Filas - 1
    For i = 0 To Columnas - 1
            Load Shape1(m)
            Load Label1(m)
            Shape1(m).Top = n * 600
            Label1(m).Top = n * 600
            Shape1(m).Left = Shape1(m).Width * i
            Label1(m).Left = Label1(m).Width * i
            Shape1(m).visible = True
            Label1(m).visible = True
            List1.AddItem List1.ListCount & " Deselec"
            m = m + 1
    Next i
Next n
End Function
