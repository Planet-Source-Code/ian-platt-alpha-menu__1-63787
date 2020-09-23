VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SLIDERSPEED 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   2760
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   2760
   End
   Begin VB.PictureBox BUTTONUP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "FormMain.frx":27A2
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   2
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
      Picture         =   "FormMain.frx":4B0E
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   1
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
      Picture         =   "FormMain.frx":6E7A
      ScaleHeight     =   450
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activate menu"
         Height          =   195
         Left            =   80
         TabIndex        =   3
         Top             =   100
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

       '*****************************************************************************
        'DIM VARIOUS VARIABLES TO GLOBALLY CHANGE BUTTON SIZE, POSITION, FORM SIZE, POSITION ETC
        '*****************************************************************************

Dim SLIDESPEED As Integer ' SET THIS TO CONTROL SPEED - BIGGER NUMBER FASTER - KEEP TO MULTIPLES OF 25


Dim BUTTONHEIGHT As Integer
Dim BUTTONWIDTH As Integer

Dim FORMHEIGHT As Integer
Dim FORMWIDTH As Integer
Dim DESKTOPWIDTH As Integer
Dim DESKTOPHEIGHT As Integer
Dim WALLPAPER As String
Dim CLICKSOUND As String
Dim EXPANDSOUND As String
Dim FORMTOP As Integer
Dim FORMLEFT As Integer

Dim lngSuccess As Long
Dim Togglemenu As Integer
Dim Counter As Integer



Private Sub Form_Load()
        '*****************************************************************************
        'SET VARIABLES
        '*****************************************************************************



'================================================================================================
' ***********************  EASY CHANGE SETTINGS FOR SLIDE SPEED, SOUNDS ETC *********************
'================================================================================================
    SLIDESPEED = 25 ' KEEP TO MULTIPLES OF 25 OR WONT WORK!
    
    ' DECLARE SLIDESPEED AVAILABLE TO OTHER FORMS
    Me.SLIDERSPEED = SLIDESPEED
    
    CLICKSOUND = App.Path & "\Click.WAV" ' SET YOUR CHOICE OF WAV SOUND
    EXPANDSOUND = App.Path & "\ELECTRIC.WAV" ' SET YOUR CHOICE OF WAV SOUND
    
'================================================================================================
    
    BUTTONHEIGHT = 450
    BUTTONWIDTH = 1500
    FORMHEIGHT = BUTTONHEIGHT
    FORMWIDTH = BUTTONWIDTH
    DESKTOPWIDTH = Screen.Width '/ 15 ' Returns pixels not TWIPS
    DESKTOPHEIGHT = Screen.Height '/ 15 ' Returns pixels not TWIPS
    WALLPAPER = App.Path & "\Wallpaper.bmp"

    FORMTOP = DESKTOPHEIGHT - BUTTONHEIGHT - 1000
    FORMLEFT = DESKTOPWIDTH - 1500 - 500
    
        '*****************************************************************************
        'INITIALISE FORM POSITION AND SIZE
        '*****************************************************************************
    FormMain.Top = FORMTOP
    FormMain.Left = FORMLEFT
    FormMain.Height = FORMHEIGHT
    FormMain.Width = FORMWIDTH
    'MsgBox "Desktop Height= " & DESKTOPHEIGHT & "Desktop Width = " & DESKTOPWIDTH, vbOKOnly
    
        '*****************************************************************************
        'CHANGE DESKTOP WALLPAPER TO DESIRED .BMP
        '*****************************************************************************
    lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, WALLPAPER, 0)
    
End Sub
        '*****************************************************************************
        'CHANGE BUTTON UP PICTURE TO BUTTON DOWN ON MOUSEDOWN/VICE VERSA ON BUTTON UP
        '*****************************************************************************
Private Sub BUTTON_MouseDown(BUTTON As Integer, Shift As Integer, X As Single, Y As Single)
    Togglemenu = Togglemenu + 1
    Me.BUTTON.Picture = Me.BUTTONDOWN.Picture
    PlaySound CLICKSOUND
    
    '   If Clicked - Expand Menu and Label as Deactivate ready to Collapse menu
        If Togglemenu = 1 Then
        Me.Label1.Caption = "Deactivate"
        Call ExpandMenu
        PlaySound EXPANDSOUND
        End If
    
    '   If Clicked - Collapse Menu and Label as Activate ready to Expand menu
        If Togglemenu = 2 Then
        Me.Label1.Caption = "Activate Menu"
        Togglemenu = 0
        Call CollapseMenu
        PlaySound EXPANDSOUND
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
            WALLPAPER = App.Path & "\Default_Wallpaper.bmp"
            lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, WALLPAPER, 0)
        End
    End Select
End Sub
        '*****************************************************************************
        'ON PROGRAM EXIT UNLOAD ALL, SET WALLAPER TO DEFAULT AND TERMINATE
        '*****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    WALLPAPER = App.Path & "\Default_Wallpaper.bmp"
    lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, WALLPAPER, 0)
    End
End Sub
        '*****************************************************************************
        'BEGIN MENU SETUP
        '*****************************************************************************

Private Sub ExpandMenu()
Menu1.Top = FormMain.Top
Menu1.Left = FormMain.Left
' RESET MENU  CAPTIONS TO DEFAULT IF MENU 1 PANEL LEFT OPEN ON COLLAPSE
Menu1.Label1.FontBold = False: Menu1.Label1.ForeColor = vbBlack
Menu2.Label1.FontBold = False: Menu2.Label1.ForeColor = vbBlack

Timer1.Enabled = True

End Sub

Private Sub CollapseMenu()
FormMenu1Panel.Hide
FORMMENU2PANEL.Hide
Timer6.Enabled = True
End Sub

        '*****************************************************************************
        'EXPAND MENU - COMPLETE MENU1 BEFORE STARTING MENU2 ETC. SET MENU2 TOP TO MENU 1 ETC
        '*****************************************************************************
  
Private Sub Timer1_Timer()
Menu1.Show
Menu1.Top = Menu1.Top - SLIDESPEED
If Menu1.Top = FormMain.Top - 500 Then Timer1.Enabled = False: Menu2.Top = Menu1.Top: Menu2.Left = Menu1.Left: Timer2.Enabled = True

End Sub
Private Sub Timer2_Timer()
Menu2.Show
Menu2.Top = Menu2.Top - SLIDESPEED
If Menu2.Top = Menu1.Top - 500 Then Timer2.Enabled = False: MENU3.Top = Menu2.Top: MENU3.Left = Menu2.Left: Timer5.Enabled = True
End Sub
Private Sub Timer5_Timer()
MENU3.Show
MENU3.Top = MENU3.Top - SLIDESPEED
If MENU3.Top = Menu2.Top - 500 Then Timer5.Enabled = False:
End Sub

        '*****************************************************************************
        'COLLAPSE MENU - COMPLETE last menu BEFORE STARTING previous MENU ETC. REVERSE OF EXPAND MENUS
        '*****************************************************************************
Private Sub Timer6_Timer()
MENU3.Top = MENU3.Top + SLIDESPEED
If MENU3.Top = Menu2.Top Then Timer6.Enabled = False: MENU3.Hide: Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Menu2.Top = Menu2.Top + SLIDESPEED
If Menu2.Top = Menu1.Top Then Timer3.Enabled = False: Menu2.Hide: Timer4.Enabled = True
End Sub
Private Sub Timer4_Timer()
Menu1.Top = Menu1.Top + SLIDESPEED
If Menu1.Top = FormMain.Top Then Timer4.Enabled = False: Menu1.Hide
End Sub

