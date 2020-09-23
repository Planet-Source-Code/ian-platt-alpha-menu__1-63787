VERSION 5.00
Begin VB.Form FormMenu1Panel 
   Appearance      =   0  'Flat
   BackColor       =   &H00EDAC8B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   0
      Picture         =   "FormMenu1Panel.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin VB.Frame Frame1 
         BackColor       =   &H0096B400&
         Caption         =   "Wallpapers: "
         Height          =   2000
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1200
         Begin VB.OptionButton Option3 
            BackColor       =   &H0096B400&
            Caption         =   "Default"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H0096B400&
            Caption         =   "GUI-II"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0096B400&
            Caption         =   "GUI-I"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "FormMenu1Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wallpaper As String

Private Sub Option1_Click()

Dim lngSuccess As Long
Dim strBitmapImage As String
strBitmapImage = "Wallpaper.bmp"
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)

End Sub

Private Sub Option2_Click()

Dim lngSuccess As Long
Dim strBitmapImage As String
strBitmapImage = "ScrBack1024x768.bmp"
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)

End Sub

Private Sub Option3_Click()
Dim lngSuccess As Long
Dim strBitmapImage As String
strBitmapImage = "Default_Wallpaper.bmp"
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)
End Sub

        '*****************************************************************************
        'ALLOW ESCAPE KEY TO EXIT PROGRAM
        '*****************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
            wallpaper = App.Path & "\Default_Wallpaper.bmp"
            lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wallpaper, 0)
        End
    End Select
End Sub
