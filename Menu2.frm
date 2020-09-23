VERSION 5.00
Begin VB.Form Menu2 
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
   Begin VB.TextBox PANELEXPAND 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   1560
   End
   Begin VB.PictureBox BUTTONUP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "Menu2.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   1080
      Width           =   1500
   End
   Begin VB.PictureBox BUTTONDOWN 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "Menu2.frx":236C
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin VB.PictureBox BUTTON 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "Menu2.frx":46D8
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programs"
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   105
         Width           =   660
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Menu2"
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

Dim PANELSLIDE As Integer


Private Sub Form_Load()
        '*****************************************************************************
        'SET VARIABLES
        '*****************************************************************************
    
    Me.PANELEXPAND = FormMain.SLIDERSPEED
    PANELSLIDE = Me.PANELEXPAND
    
    
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
    Menu2.Height = FORMHEIGHT
    Menu2.Width = FORMWIDTH

   
End Sub
        '*****************************************************************************
        'CHANGE BUTTON UP PICTURE TO BUTTON DOWN ON MOUSEDOWN/VICE VERSA ON BUTTON UP
        '*****************************************************************************
Private Sub BUTTON_MouseDown(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    Togglemenu = Togglemenu + 1
    Me.BUTTON.Picture = Me.BUTTONDOWN.Picture
    PlaySound CLICKSOUND
    
        If Togglemenu = 1 Then
        Me.Label1.Caption = "Programs": Me.Label1.FontBold = True: Me.Label1.ForeColor = vbRed
        Call menu2panelshow
        End If
    
    
        If Togglemenu = 2 Then
        Me.Label1.Caption = "Programs": Me.Label1.FontBold = False: Me.Label1.ForeColor = vbBlack
        Call menu2panelhide
        Togglemenu = 0
        End If
End Sub

Private Sub BUTTON_MouseUp(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    Me.BUTTON.Picture = Me.BUTTONUP.Picture
End Sub

        '*****************************************************************************
        'SHOW FORMMENUPANEL1 AND DO SLIDES ETC
        '*****************************************************************************
Private Sub menu2panelshow()

FORMMENU2PANEL.Width = Menu2.Width
FORMMENU2PANEL.Height = 2250
FORMMENU2PANEL.Top = Menu2.Top - 2250 + 450
FORMMENU2PANEL.Left = Menu2.Left
FORMMENU2PANEL.Show
Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
FORMMENU2PANEL.Left = FORMMENU2PANEL.Left - (PANELSLIDE * 4)
    If FORMMENU2PANEL.Left = Menu2.Left - 1500 Then
    FORMMENU2PANEL.Left = FORMMENU2PANEL.Left - 50
    Timer1.Enabled = False
    End If

End Sub

Private Sub menu2panelhide()
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
FORMMENU2PANEL.Left = FORMMENU2PANEL.Left + (PANELSLIDE * 4)
    If FORMMENU2PANEL.Left > Menu2.Left Then FORMMENU2PANEL.Hide: Timer2.Enabled = False

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
