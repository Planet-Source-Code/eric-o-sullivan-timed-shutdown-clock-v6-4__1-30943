Attribute VB_Name = "PlayWav"
'This program was made by me,
'Eric O' Sullivan. CompApp Technologys (tm)
'is my company. If this product is unsatisfactory
'feel free to contact me at
'DiskJunky@hotmail.com
'================================================
'================================================

'API function to play .wav
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Sub PlaySound(File As String)
Dim Response As Integer

Const Sync = 1

If (Dir(File) <> "") And (Trim(LCase(Right(File, 4))) = ".wav") Then
    'if file exists and is a .wav, then play
    Response = sndPlaySound(ByVal CStr(File), Sync)
End If
End Sub


