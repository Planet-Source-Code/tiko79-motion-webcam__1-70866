Attribute VB_Name = "Module1"
Option Explicit
#If Win32 Then
  Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" _
                      (lpszSoundName As Any, ByVal uFlags As Long) As Long
#Else
  Private Declare Function sndPlaySound Lib "MMSYSTEM" ( _
                     lpszSoundName As Any, ByVal uFlags%) As Integer
#End If
Public Const SND_SYNC = &H0
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
Dim bytSound() As Byte
bytSound = LoadResData(vntResourceID, "SONIDOS")
If IsMissing(vntFlags) Then
   vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
End If
If (vntFlags And SND_MEMORY) = 0 Then
   vntFlags = vntFlags Or SND_MEMORY
End If
sndPlaySound bytSound(0), vntFlags
End Sub



