VERSION 5.00
Object = "{1E5ED54D-4BB2-11D6-8DC1-90B225C3E54F}#1.0#0"; "GRADBUTTON.OCX"
Begin VB.Form frmFSControls 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFSControls.frx":0000
   ScaleHeight     =   675
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCurrentTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF860D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00623100&
      Height          =   255
      Left            =   585
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   0
      ToolTipText     =   "Current Time"
      Top             =   180
      Width           =   1575
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1
      Left            =   0
      Top             =   630
   End
   Begin GradButton.GradientButton cmdExit 
      Height          =   255
      Left            =   2430
      TabIndex        =   1
      ToolTipText     =   "Exit DMP3"
      Top             =   180
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor1  =   16777215
      GradientColor2  =   13027014
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16711935
      Picture         =   "frmFSControls.frx":670A
      Style           =   1
      UseHover        =   0   'False
   End
   Begin GradButton.GradientButton cmdHide 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   2170
      TabIndex        =   2
      ToolTipText     =   "Exit DMP3"
      Top             =   180
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      Caption         =   "H"
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownForeColor   =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientColor1  =   16777215
      GradientColor2  =   13027014
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16711935
      Style           =   1
      UseHover        =   0   'False
   End
   Begin GradButton.GradientButton cmdPause 
      Height          =   255
      Left            =   135
      TabIndex        =   3
      ToolTipText     =   "Pause Media"
      Top             =   180
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor1  =   12632256
      GradientColor2  =   16744448
      GradientType    =   2
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16711935
      Picture         =   "frmFSControls.frx":689C
      Style           =   1
      UseHover        =   0   'False
   End
   Begin GradButton.GradientButton cmdPlay 
      Height          =   255
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "Play Media"
      Top             =   180
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor1  =   12632256
      GradientColor2  =   16744448
      GradientType    =   2
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16711935
      Picture         =   "frmFSControls.frx":69C6
      Style           =   1
      UseHover        =   0   'False
   End
End
Attribute VB_Name = "frmFSControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CShow As Boolean
Const Aliasname As String = "DMP3Media"
Dim Result As String

Private Sub cmdExit_Click()

    Dim MovieSize(1 To 4) As Integer
    
    CShow = False
    frmMedia.Caption = "DMP3 ALPHA - Video"
    CShow = False
    MovieSize(3) = (frmMain.ActualWidth * 15)
    MovieSize(4) = (frmMain.ActualHeight * 15)
    frmMain.WindowState = vbNormal
    MovieSize(1) = frmMain.left
    MovieSize(2) = frmMain.top
    frmMedia.Hide
    frmMedia.left = MovieSize(1)
    frmMedia.top = MovieSize(2)
    frmMedia.Width = MovieSize(3)
    frmMedia.Height = MovieSize(4)
    Result = PutMultimedia(frmMedia.hwnd, Aliasname, Val(0), Val(0), Val(0), Val(0))
    frmMedia.Height = frmMedia.Height + 350
    frmMain.Fullscreen = False
    frmMedia.Show
    Unload Me
   
            
End Sub

Private Sub cmdHide_Click()
    
    CShow = False

End Sub

Private Sub cmdPause_Click()

    If frmMain.Filename = "" Then Exit Sub

    frmMain.PauseMedia
    
    cmdPlay.Visible = True
    cmdPause.Visible = False
    
End Sub

Private Sub cmdPlay_Click()

    If frmMain.Filename = "" Then Exit Sub
    
    frmMain.ResumeMedia
    
    cmdPause.Visible = True
    cmdPlay.Visible = False
    
End Sub

Private Sub Form_Load()

    If Me.Picture <> 0 Then
        Call SetAutoRgn(Me)
    End If
    
    CShow = True

    Me.top = Screen.Height - Me.Height - 10
    Me.left = Screen.Width - Me.Width - 10

End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    
    Dim lFlag As String
    
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.left / Screen.TwipsPerPixelX, _
    myfrm.top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
End Sub

Private Sub tmrUpdate_Timer()

    If CShow = True Then
        AlwaysOnTop Me, True
    Else
        AlwaysOnTop Me, False
        AlwaysOnTop frmMedia, True
    End If
    
    picCurrentTime.Cls
    picCurrentTime.Print frmMain.TimeToString(frmMain.CurrentTime)

End Sub
