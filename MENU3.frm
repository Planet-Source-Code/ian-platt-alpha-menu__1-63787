VERSION 5.00
Begin VB.Form MENU3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.PictureBox BUTTON 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "MENU3.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   0
      Width           =   1500
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit [ESC key]"
         Height          =   195
         Left            =   75
         TabIndex        =   3
         Top             =   105
         Width           =   1005
      End
   End
   Begin VB.PictureBox BUTTONDOWN 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "MENU3.frx":236C
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   1560
      Width           =   1500
   End
   Begin VB.PictureBox BUTTONUP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "MENU3.frx":46D8
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   1080
      Width           =   1500
   End
End
Attribute VB_Name = "MENU3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

       '*****************************************************************************
        'DIM VARIOUS VARIABLES TO GLOBALLY CHANGE BUTTON SIZE, POSITION, FORM SIZE, POSITION ETC
        '*****************************************************************************
Dim BUTTONHEIGHT As Integer
Dim BUTTONWIDTH As Integer

Dim FORMHEIGHT As Integer
Dim FORMWIDTH As Integer

Dim CLICKSOUND As String
Dim FORMTOP As Integer
Dim FORMLEFT As Integer
Dim wallpaper As String
Dim lngSuccess As Long
Dim Togglemenu As Integer
Dim Menu1Slide As Integer


Private Sub Form_Load()
        '*****************************************************************************
        'SET VARIABLES
        '*****************************************************************************
    BUTTONHEIGHT = 450
    BUTTONWIDTH = 1500
    FORMHEIGHT = BUTTONHEIGHT
    FORMWIDTH = BUTTONWIDTH
    CLICKSOUND = App.Path & "\Click.WAV"
    FORMTOP = Menu1.Top
    FORMLEFT = Menu1.Left
    
        '*****************************************************************************
        'INITIALISE FORM AND SIZE
        '*****************************************************************************
    MENU3.Height = FORMHEIGHT
    MENU3.Width = FORMWIDTH

   
End Sub
        '*****************************************************************************
        'CHANGE BUTTON UP PICTURE TO BUTTON DOWN ON MOUSEDOWN/VICE VERSA ON BUTTON UP
        '*****************************************************************************
Private Sub BUTTON_MouseDown(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    Togglemenu = Togglemenu + 1
    Me.BUTTON.Picture = Me.BUTTONDOWN.Picture
    PlaySound CLICKSOUND
    
    End
    
        If Togglemenu = 1 Then
        Me.Label1.Caption = "Exit": Me.Label1.FontBold = True: Me.Label1.ForeColor = vbRed
        End If
    
    
        If Togglemenu = 2 Then
        Me.Label1.Caption = "Exit": Me.Label1.FontBold = False: Me.Label1.ForeColor = vbBlack
        Togglemenu = 0
        End If
End Sub

Private Sub BUTTON_MouseUp(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    Me.BUTTON.Picture = Me.BUTTONUP.Picture
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
