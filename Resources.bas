Attribute VB_Name = "Resources"
Option Explicit






        '*****************************************************************************
        'CHANGE YOUR DESKTOP WALLPAPER
        '*****************************************************************************
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20
'***************************************

        '*****************************************************************************
        'PLAY A .WAV FILE
        '*****************************************************************************
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName _
As String, ByVal uFlags As Long) As Long

Public Sub PlaySound(strFileName As String)
    sndPlaySound strFileName, 1
End Sub

