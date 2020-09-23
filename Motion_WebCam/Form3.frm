VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFF3EF&
   Caption         =   "Picture Viewer"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12150
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   596
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3960
      Top             =   8520
   End
   Begin VB.FileListBox filemain 
      Height          =   1845
      Left            =   10200
      Pattern         =   "*.bmp;*.gif;*.jpg;*.jpeg;*.jpe;*.jfif"
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picin 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   10200
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   4
      Top             =   0
      Width           =   1875
   End
   Begin VB.PictureBox picout 
      BackColor       =   &H00FFF3EF&
      BorderStyle     =   0  'None
      DragIcon        =   "Form3.frx":01B6
      Height          =   7935
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   529
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   777
      TabIndex        =   1
      Top             =   120
      Width           =   11655
      Begin VB.Line Linmain 
         Index           =   3
         X1              =   128
         X2              =   632
         Y1              =   432
         Y2              =   432
      End
      Begin VB.Line Linmain 
         Index           =   2
         X1              =   648
         X2              =   648
         Y1              =   80
         Y2              =   416
      End
      Begin VB.Line Linmain 
         Index           =   1
         X1              =   112
         X2              =   112
         Y1              =   80
         Y2              =   416
      End
      Begin VB.Line Linmain 
         Index           =   0
         X1              =   128
         X2              =   632
         Y1              =   64
         Y2              =   64
      End
      Begin VB.Image imgimage 
         Height          =   5040
         Left            =   1920
         Top             =   1200
         Width           =   7560
      End
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11280
      Tag             =   "1"
      Top             =   7440
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   2640
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   8475
      Width           =   360
   End
   Begin MSComctlLib.ImageList imlmain 
      Left            =   11040
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   16
      MaskColor       =   16774127
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0FF8
            Key             =   "left"
            Object.Tag             =   "Previous Image (Left Arrow)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":134A
            Key             =   "right"
            Object.Tag             =   "Next Image (Right Arrow)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":169C
            Key             =   "bestfiton"
            Object.Tag             =   "Best Fit (Ctrl+B)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":19EE
            Key             =   "actualon"
            Object.Tag             =   "Actual Size (Ctrl+A)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":1D40
            Key             =   "slideshow"
            Object.Tag             =   "Start Slideshow (F11)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2092
            Key             =   "magnify"
            Object.Tag             =   "Zoom In (+)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":23E4
            Key             =   "minify"
            Object.Tag             =   "Zoom Out (-)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2736
            Key             =   "delete"
            Object.Tag             =   "Delete (Delete)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2A88
            Key             =   "editon"
            Object.Tag             =   "Open for editing (Ctrl+E)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2DDA
            Key             =   "play"
            Object.Tag             =   "Start the slideshow"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":312C
            Key             =   "pause"
            Object.Tag             =   "Pause the slideshow"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":347E
            Key             =   "stop"
            Object.Tag             =   "Stop the slideshow"
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar hscrmain 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.VScrollBar vscrmain 
      Height          =   7935
      Left            =   11760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgmain 
      Height          =   330
      Index           =   2
      Left            =   11640
      Picture         =   "Form3.frx":37D0
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgmain 
      Height          =   330
      Index           =   1
      Left            =   11640
      Picture         =   "Form3.frx":3E42
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgmain 
      Height          =   330
      Index           =   0
      Left            =   11640
      Picture         =   "Form3.frx":44B4
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Menu mnudelay 
      Caption         =   "Time"
      Visible         =   0   'False
      Begin VB.Menu mnutime 
         Caption         =   "0 seconds"
         Index           =   0
      End
      Begin VB.Menu mnutime 
         Caption         =   "1 second"
         Index           =   1
      End
      Begin VB.Menu mnutime 
         Caption         =   "2 seconds"
         Index           =   2
      End
      Begin VB.Menu mnutime 
         Caption         =   "5 seconds"
         Index           =   3
      End
      Begin VB.Menu mnutime 
         Caption         =   "10 seconds"
         Index           =   4
      End
      Begin VB.Menu mnutime 
         Caption         =   "20 seconds"
         Index           =   5
      End
      Begin VB.Menu mnutime 
         Caption         =   "30 seconds"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Reg As CRegSettings
Dim currentpicture As Long, currentstate As Long, origx As Long, origy As Long, imagename As String, lastop As Long, isinslideshowmode As Boolean, oldimage As String
Public Sub linemove(line2move As Line, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    line2move.x1 = x1
    line2move.x2 = x2
    line2move.y1 = y1
    line2move.y2 = y2
End Sub
Public Sub loadimage(Filename As String)
    On Error Resume Next
        imagename = LCase(Filename)
        imgimage.Stretch = False
        imgimage.picture = LoadPicture(Filename)
        imgimage.Stretch = True
        origx = imgimage.Width
        origy = imgimage.Height
        bestfit True
    Caption = Right(Filename, Len(Filename) - InStrRev(Filename, "\")) & " - Picture Viewer"
End Sub
Public Function ISaDIR(Filename As String) As Boolean
    On Error Resume Next
    If Len(Filename) > 0 Then ISaDIR = (GetAttr(Filename) And vbDirectory) = vbDirectory
End Function

Public Function ShellFile(hWnd As Long, strOperation As String, ByVal File As String, WindowStyle As VbAppWinStyle) As Long
'"Open, Print, Explore, Find, Edit, Play, 0&"
    ShellFile = ShellExecute(hWnd, strOperation, File, vbNullString, App.path, WindowStyle)
End Function

Public Sub bestfit(Optional force As Boolean)
    Dim tempx As Long, tempy As Long
    tempx = origx
    tempy = origy
    thumbsize tempx, tempy, picout.Width, picout.Height, force
    imgimage.Move (picout.Width / 2) - (tempx / 2), (picout.Height / 2) - (tempy / 2), tempx, tempy
    Form_Resize
End Sub
Public Sub actualsize()
    imgimage.Move (picout.Width / 2) - (origx / 2), (picout.Height / 2) - (origx / 2), origx, origy
    Form_Resize
End Sub
Public Sub magnify()
    imgimage.Move (picout.Width / 2) - (imgimage.Width * 0.6), (picout.Height / 2) - (imgimage.Height * 0.6), imgimage.Width * 1.2, imgimage.Height * 1.2
    Form_Resize
End Sub
Public Sub minify()
    imgimage.Move (picout.Width / 2) - (imgimage.Width * 0.4), (picout.Height / 2) - (imgimage.Height * 0.4), imgimage.Width * 0.8, imgimage.Height * 0.8
    Form_Resize
End Sub
Public Function findfileindex(ByVal Filename As String) As Long
    Dim temp As Long
    findfileindex = -1
    filemain.Refresh
    filemain.path = Left(Filename, InStrRev(Filename, "\"))
    Filename = LCase(Right(Filename, Len(Filename) - InStrRev(Filename, "\")))
    For temp = 0 To (filemain.ListCount - 1)
        If LCase(filemain.List(temp)) = Filename Then
            findfileindex = temp
            Exit For
        End If
    Next
End Function

Public Sub seektoimage(reference As Long, Optional currimage As Long = -1)
    Dim directory As String, temp As Long
    directory = Left(imagename, InStrRev(imagename, "\"))
    filemain.path = directory
    
    If currimage = -1 Then currimage = findfileindex(imagename) + reference Else currimage = currimage + reference
    If currimage = filemain.ListCount Then currimage = 0
    If currimage = -1 Then currimage = filemain.ListCount - 1
    
    If filemain.ListCount > 0 Then loadimage directory & filemain.List(currimage)
    
    Timer.Tag = reference
End Sub

Public Sub drawpicture(pbox As PictureBox, Optional State As Long = -1, Optional picture As String)
    pbox.Width = 24
    pbox.Height = 22
    If State = -1 Then
        pbox.picture = LoadPicture(Empty)
        pbox.backcolor = &HFFF3EF
    Else
        pbox.picture = imgmain(State).picture
    End If
    imlmain.ListImages.Item(picture).Draw pbox.hDC, IIf(State <> 1, 5, 4), IIf(State <> 1, 3, 4), imlTransparent
    pbox.ToolTipText = imlmain.ListImages.Item(picture).Tag
    pbox.Tag = picture
End Sub

Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37: picmain_Click 0 'left
    Case 39: picmain_Click 1 'right
    
    Case 65: picmain_Click 3 'ctrl+a actual size
    Case 66: picmain_Click 2 'ctrl+b best fit
    Case 69: picmain_Click 8 'ctrl+e edit
    
    Case 83, 112 'f1 slideshow
    Case 107, 187, 38: picmain_Click 5 '+ magnify
    Case 109, 189, 40: picmain_Click 6 '- minify
    
    Case 46: picmain_Click 7 'delete

End Select
End Sub

Private Sub Form_Load()
Set Reg = New CRegSettings
Reg.Company = "TIKO®\Alberto Fedriani"
Reg.AppName = "WebCam Motion"
    Dim spoth As Long, temps() As String
    InitOpen "Image" & Chr(0) & filemain.Pattern, "Load an image"
    temps = Split("left right bestfiton actualon slideshow magnify minify delete editon play pause left right stop", " ")
    picmain(0).backcolor = Me.backcolor
    For spoth = 0 To 13
        If spoth > 0 Then
            Load picmain(spoth)
            picmain(spoth).Left = picmain(spoth - 1).Left + picmain(spoth - 1).Width + 4
            picmain(spoth).Top = picmain(spoth - 1).Top
            picmain(spoth).visible = True
        End If
        If spoth = 9 Then
            Set picmain(9).Container = picin
            picmain(9).Left = picin.Width - ((picmain(9).Width + 1) * 5)
            picmain(9).Top = 0
        End If
        If spoth > 9 Then
            Set picmain(spoth).Container = picin
            picmain(spoth).Left = picmain(spoth - 1).Left + picmain(spoth - 1).Width + 1
            picmain(spoth).Top = picmain(9).Top
        End If
        drawpicture picmain(spoth), -1, temps(spoth)
    Next
    picin.visible = False
    currentpicture = -1
    
    WindowState = Reg.GetSetting("HKEY_CURRENT_USER", "Preview", "WindowState", WindowState)
    Width = Reg.GetSetting("HKEY_CURRENT_USER", "Preview", "Width", Width)
    Height = Reg.GetSetting("HKEY_CURRENT_USER", "Preview", "Height", Height)
    Top = Reg.GetSetting("HKEY_CURRENT_USER", "Preview", "Top", Top)
    Left = Reg.GetSetting("HKEY_CURRENT_USER", "Preview", "Left", Left)
    filemain.path = App.path & "\Recorded"
    If filemain.ListCount > 0 Then
        loadimage filemain.path & "\" & filemain.List(0)
    Else
        MsgBox "No hay imagenes para mostrar.", vbInformation + vbOKOnly, "Sin imagenes"
        Unload Me
    End If
    
'    If Command <> Empty Then
'        loadimage Command
'    Else
'        loadimage Open_File(Me.hWnd)
'    End If
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If currentpicture > -1 Then
    currentstate = -1
    drawpicture picmain(currentpicture), currentstate, picmain(currentpicture).Tag
    currentpicture = -1
End If
End Sub

Public Sub Form_Resize()
'GNDN
If WindowState <> vbMinimized Then
Dim tempx As Long, tempy As Long, temp As Long
tempx = Width / 15
tempy = Height / 15

hscrmain.visible = False
vscrmain.visible = False

If isinslideshowmode = False Then
picmain(0).Left = tempx / 2 - (picmain(0).Width * 9 + 4 * 8) / 2
picmain(0).Top = tempy - 61
For temp = 1 To 8
    picmain(temp).Left = picmain(temp - 1).Left + picmain(temp - 1).Width + 4
    picmain(temp).Top = picmain(temp - 1).Top
Next

picout.Width = tempx - 24
picout.Height = tempy - 80

If imgimage.Width > picout.Width Then
    hscrmain.visible = True
    picout.Height = picout.Height - hscrmain.Height
    hscrmain.Top = picout.Height + picout.Top
    hscrmain.Width = picout.Width - IIf(imgimage.Height > picout.Height, vscrmain.Width, 0)
    hscrmain.Max = imgimage.Width - picout.Width
    hscrmain.Value = hscrmain.Max / 2
    hscrmain.LargeChange = hscrmain.Max / 5
End If
If imgimage.Height > picout.Height Then
    vscrmain.visible = True
    picout.Width = picout.Width - vscrmain.Width
    vscrmain.Left = picout.Left + picout.Width
    vscrmain.Height = picout.Height
    vscrmain.Max = imgimage.Height - picout.Height
    vscrmain.Value = vscrmain.Max / 2
    vscrmain.LargeChange = vscrmain.Max / 5
End If

If imgimage.Height > picout.Height Then vscrmain.visible = True
Else
    picout.Width = tempx - picout.Left * 2
    picout.Height = tempy - picout.Top * 2 - 30
End If

imgimage.Move (picout.Width / 2) - (imgimage.Width / 2), (picout.Height / 2) - (imgimage.Height / 2)
linemove Linmain(0), imgimage.Left - 1, imgimage.Top - 1, imgimage.Left + imgimage.Width + 1, imgimage.Top - 1
linemove Linmain(1), imgimage.Left - 1, imgimage.Top - 1, imgimage.Left - 1, imgimage.Top + imgimage.Height + 1
linemove Linmain(2), imgimage.Left + imgimage.Width, imgimage.Top, imgimage.Left + imgimage.Width, imgimage.Top + imgimage.Height + 1
linemove Linmain(3), imgimage.Left, imgimage.Top + imgimage.Height, imgimage.Left + imgimage.Width, imgimage.Top + imgimage.Height

picin.Left = tempx - picin.Width - 7
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Reg.SaveSetting "HKEY_CURRENT_USER", "Preview", "WindowState", WindowState
    WindowState = 0
    Reg.SaveSetting "HKEY_CURRENT_USER", "Preview", "Width", Width
    Reg.SaveSetting "HKEY_CURRENT_USER", "Preview", "Height", Height
    Reg.SaveSetting "HKEY_CURRENT_USER", "Preview", "Top", Top
    Reg.SaveSetting "HKEY_CURRENT_USER", "Preview", "Left", Left
    End Sub

Public Sub hscrmain_Change()
    imgimage.Left = -hscrmain.Value
End Sub

Private Sub hscrmain_Scroll()
    hscrmain_Change
End Sub

Private Sub imgimage_Click()
Select Case lastop
    Case 0, 1, 11, 12, 5, 6: picmain_Click lastop * 1 'convert integer/long
End Select
End Sub

Private Sub imgimage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgimage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim temp As Boolean
temp = Timer.Enabled
If Button = vbRightButton Then
    Timer.Enabled = False
    ShellContextMenu Me.hWnd, picout, X + imgimage.Left, Y + imgimage.Top, Shift, imagename
    Timer.Enabled = temp
End If
End Sub

Private Sub mnutime_Click(Index As Integer)
    Select Case Index
        Case 0: Timer.Interval = 1
        Case 1: Timer.Interval = 1000
        Case 2: Timer.Interval = 2000
        Case 3: Timer.Interval = 5000
        Case 4: Timer.Interval = 10000
        Case 5: Timer.Interval = 20000
        Case 6: Timer.Interval = 30000
    End Select
End Sub

Public Sub picmain_Click(Index As Integer)
lastop = Index
Dim temp As Long
Select Case Index
    Case 0, 11: seektoimage -1 'Previous
    Case 1, 12: seektoimage 1 'Next
    
    Case 2: bestfit True 'Best Fit
    Case 3: actualsize 'Actual Size
    
    Case 4: switchmode True, vbBlack 'slideshow
    
    Case 5: magnify
    Case 6: minify
    
    Case 7 'delete
        If imagename <> Empty Then
            temp = findfileindex(imagename)
            If File_Delete(imagename) = False Then
                seektoimage 1, temp
                filemain.Refresh
            End If
        End If
    Case 8: ShellFile Me.hWnd, "Open", imagename, vbHide 'Edit
    
    Case 9: Timer.Enabled = True 'play
    Case 10: Timer.Enabled = False 'pause
    Case 13: switchmode False, &HFFF3EF 'stop
End Select
End Sub
Public Sub switchmode(visible As Boolean, backcolor As Long)
        Static winstate As Long
        isinslideshowmode = visible
        Me.backcolor = backcolor
        picin.visible = visible
        picout.backcolor = backcolor
        Timer.Enabled = visible
        
        If visible Then
            oldimage = imagename
            winstate = WindowState
            WindowState = vbMaximized
        Else
            loadimage oldimage
            WindowState = winstate
        End If
        Form_Resize
End Sub
Private Sub picmain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub picmain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
currentpicture = Index
If currentstate <> 1 Then
    currentstate = 1
    drawpicture picmain(Index), currentstate, picmain(Index).Tag
End If
End Sub

Private Sub picmain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If currentpicture <> Index Then
    If currentpicture > -1 Then drawpicture picmain(currentpicture), -1, picmain(currentpicture).Tag
    currentpicture = Index
    currentstate = 0
    drawpicture picmain(Index), currentstate, picmain(Index).Tag
End If
End Sub

Private Sub picmain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
currentpicture = Index
If currentstate <> 2 Then
    currentstate = 2
    drawpicture picmain(Index), currentstate, picmain(Index).Tag
End If

If Button = vbRightButton And Index = 9 Then
    PopupMenu mnudelay
End If
End Sub
Private Sub picout_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub picout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub picout_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    loadimage Data.Files(1)
End Sub

Private Sub Timer_Timer()
    seektoimage Timer.Tag
End Sub

Private Sub Timer1_Timer()
filemain.Refresh
End Sub

Public Sub vscrmain_Change()
    imgimage.Top = -vscrmain.Value
End Sub

Private Sub vscrmain_Scroll()
    vscrmain_Change
End Sub
Public Sub thumbsize(ByRef picwidth As Long, ByRef picheight As Long, ByRef thumbwidth As Long, ByRef thumbheight As Long, Optional forcefit As Boolean = False)
    If forcefit Then
        If picheight < thumbheight Then
            picwidth = picwidth * thumbheight / picheight
            picheight = thumbheight
        End If
    End If
    If picwidth > thumbwidth Then
        picheight = Round(picheight / (picwidth / thumbwidth), 0)
        picwidth = thumbwidth
    End If
    If picheight > thumbheight Then
        picwidth = picwidth / (picheight / thumbheight)
        picheight = picheight / (picheight / thumbheight)
    End If
End Sub
