VERSION 5.00
Begin VB.Form Menu1 
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1440
   End
   Begin VB.TextBox PANELEXPAND 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
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
      Picture         =   "Menu1.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   0
      Width           =   1500
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wallpaper"
         Height          =   195
         Left            =   75
         TabIndex        =   3
         Top             =   105
         Width           =   720
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
      Picture         =   "Menu1.frx":236C
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
      Picture         =   "Menu1.frx":46D8
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   1080
      Width           =   1500
   End
End
Attribute VB_Name = "Menu1"
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

Dim lngSuccess As Long
Dim Togglemenu As Integer
Dim Menu1Slide As Integer
Dim wallpaper As String
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
    FORMTOP = FormMain.Top
    FORMLEFT = FormMain.Left
    
        '*****************************************************************************
        'INITIALISE FORM SIZE
    Menu1.Height = FORMHEIGHT
    Menu1.Width = FORMWIDTH
    
End Sub
        '*****************************************************************************
        'CHANGE BUTTON UP PICTURE TO BUTTON DOWN ON MOUSEDOWN/VICE VERSA ON BUTTON UP
        '*****************************************************************************
Private Sub BUTTON_MouseDown(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    
    Togglemenu = Togglemenu + 1
    Me.BUTTON.Picture = Me.BUTTONDOWN.Picture
    PlaySound CLICKSOUND
    
        If Togglemenu = 1 Then
        Me.Label1.Caption = "Wallpaper": Me.Label1.FontBold = True: Me.Label1.ForeColor = vbRed
        Call menu1panelshow
        End If
    
    
        If Togglemenu = 2 Then
        Me.Label1.Caption = "Wallpaper": Me.Label1.FontBold = False: Me.Label1.ForeColor = vbBlack
        Call menu1panelhide
        Togglemenu = 0
        End If
End Sub

Private Sub BUTTON_MouseUp(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    Me.BUTTON.Picture = Me.BUTTONUP.Picture
End Sub

        '*****************************************************************************
        'SHOW FORMMENUPANEL1 AND DO SLIDES ETC
        '*****************************************************************************
Private Sub menu1panelshow()

FormMenu1Panel.Width = Menu1.Width
FormMenu1Panel.Height = 2250
FormMenu1Panel.Top = Menu1.Top - 2250 + 450
FormMenu1Panel.Left = Menu1.Left
FormMenu1Panel.Show
Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
FormMenu1Panel.Left = FormMenu1Panel.Left - (PANELSLIDE * 4)
    If FormMenu1Panel.Left = Menu1.Left - 1500 Then
    FormMenu1Panel.Left = FormMenu1Panel.Left - 50
    Timer1.Enabled = False
    End If

End Sub

Private Sub menu1panelhide()
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
FormMenu1Panel.Left = FormMenu1Panel.Left + (PANELSLIDE * 4)
    If FormMenu1Panel.Left > Menu1.Left Then FormMenu1Panel.Hide: Timer2.Enabled = False

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
