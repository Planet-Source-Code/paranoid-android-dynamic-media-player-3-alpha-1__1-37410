VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1E5ED54D-4BB2-11D6-8DC1-90B225C3E54F}#1.0#0"; "GRADBUTTON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{49616C06-65BE-11D4-804C-000021F0FF9D}#1.0#0"; "ALFAFISHEZYID33VB6.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMP3 Alpha 1"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   780
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   52
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScope 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E37200&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00623100&
      Height          =   630
      Left            =   2040
      ScaleHeight     =   600
      ScaleWidth      =   2175
      TabIndex        =   21
      ToolTipText     =   "Not Finished!"
      Top             =   1230
      Width           =   2200
   End
   Begin VB.Timer tmrEndEffects 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3735
      Top             =   1350
   End
   Begin VB.ListBox lstFilenames 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "Main.frx":000C
      Left            =   3240
      List            =   "Main.frx":000E
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTitleBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   5805
      ScaleHeight     =   270
      ScaleWidth      =   75
      TabIndex        =   37
      Top             =   2745
      Width           =   75
   End
   Begin VB.PictureBox picArtistBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3060
      ScaleHeight     =   270
      ScaleWidth      =   75
      TabIndex        =   36
      Top             =   2745
      Width           =   75
   End
   Begin VB.PictureBox picNumberBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   585
      ScaleHeight     =   270
      ScaleWidth      =   75
      TabIndex        =   35
      Top             =   2745
      Width           =   75
   End
   Begin MSComctlLib.ListView lstPlaylist 
      Height          =   2610
      Left            =   105
      TabIndex        =   22
      Top             =   3015
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16774357
      BackColor       =   8404992
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   916
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   4366
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   4842
      EndProperty
   End
   Begin VB.PictureBox picFileContainer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E37200&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   0.397
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   3.836
      TabIndex        =   9
      ToolTipText     =   "Scrolling File Info"
      Top             =   550
      Width           =   2205
   End
   Begin prjDMP3.SliderControl sldVolume 
      Height          =   135
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Volume"
      Top             =   1725
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   238
      Value           =   100
   End
   Begin VB.PictureBox picTotalTime 
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
      Left            =   405
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   16
      ToolTipText     =   "Total Time"
      Top             =   1395
      Width           =   1575
   End
   Begin VB.PictureBox picRemainingTime 
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
      Left            =   405
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   15
      ToolTipText     =   "Remaining Time"
      Top             =   975
      Width           =   1575
   End
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
      Left            =   405
      ScaleHeight     =   255
      ScaleWidth      =   1575
      TabIndex        =   14
      ToolTipText     =   "Current Time"
      Top             =   555
      Width           =   1575
   End
   Begin GradButton.GradientButton cmdPrev 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Previous File"
      Top             =   2130
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
      Picture         =   "Main.frx":0010
      Style           =   1
   End
   Begin prjDMP3.SliderControl sldProgress 
      Height          =   135
      Left            =   405
      TabIndex        =   0
      ToolTipText     =   "Current Progress"
      Top             =   1920
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   238
      Max             =   1
   End
   Begin GradButton.GradientButton cmdPause 
      Height          =   255
      Left            =   745
      TabIndex        =   3
      ToolTipText     =   "Pause Media"
      Top             =   2130
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
      Picture         =   "Main.frx":0116
      Style           =   1
   End
   Begin GradButton.GradientButton cmdStop 
      Height          =   255
      Left            =   1000
      TabIndex        =   4
      ToolTipText     =   "Stop Media"
      Top             =   2130
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
      Picture         =   "Main.frx":0240
      Style           =   1
   End
   Begin GradButton.GradientButton cmdNext 
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      ToolTipText     =   "Next File"
      Top             =   2130
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
      Picture         =   "Main.frx":0322
      Style           =   1
   End
   Begin GradButton.GradientButton cmdOpen 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Open Files"
      Top             =   2130
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
      Picture         =   "Main.frx":0428
      Style           =   1
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   3735
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   4
   End
   Begin GradButton.GradientButton cmdClose 
      Height          =   255
      Left            =   1935
      TabIndex        =   17
      ToolTipText     =   "Close Media"
      Top             =   2130
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
      Picture         =   "Main.frx":0522
      Style           =   1
   End
   Begin VB.Timer tmrMiscVis 
      Interval        =   250
      Left            =   2160
      Top             =   1320
   End
   Begin VB.Timer tmrPercent 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3120
      Top             =   1320
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2655
      Top             =   1305
   End
   Begin GradButton.GradientButton cmdExit 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   4185
      TabIndex        =   18
      ToolTipText     =   "Exit DMP3"
      Top             =   90
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
      DownForeColor   =   16777215
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
      Picture         =   "Main.frx":061C
      Style           =   1
   End
   Begin GradButton.GradientButton cmdMinimize 
      Height          =   255
      Left            =   3930
      TabIndex        =   19
      ToolTipText     =   "Bring DMP3 to the tray"
      Top             =   90
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Alignment       =   8
      Appearance      =   0
      BackColor       =   13132800
      Caption         =   "_"
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
         Name            =   "Arial"
         Size            =   11.25
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
   End
   Begin GradButton.GradientButton cmdFullscreen 
      Height          =   255
      Left            =   2610
      TabIndex        =   20
      ToolTipText     =   "Fullscreen"
      Top             =   2130
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
      HoverPicture    =   "Main.frx":07AE
      MaskColor       =   16711935
      Picture         =   "Main.frx":0944
      Style           =   1
   End
   Begin GradButton.GradientButton cmdVolume 
      Height          =   255
      Left            =   405
      TabIndex        =   24
      ToolTipText     =   "Channels Control"
      Top             =   1650
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   16745997
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
      HoverPicture    =   "Main.frx":0ADA
      MaskColor       =   16711935
      Picture         =   "Main.frx":0DFC
      Style           =   1
   End
   Begin prjDMP3.SliderControl sldLeftVolume 
      Height          =   1380
      Left            =   4770
      TabIndex        =   26
      Top             =   585
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   2434
      Vertical        =   -1  'True
      Invert          =   -1  'True
   End
   Begin prjDMP3.SliderControl sldRightVolume 
      Height          =   1380
      Left            =   5040
      TabIndex        =   27
      Top             =   585
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   2434
      Vertical        =   -1  'True
      Invert          =   -1  'True
   End
   Begin GradButton.GradientButton optLeftOnly 
      Height          =   255
      Left            =   5355
      TabIndex        =   30
      ToolTipText     =   "Left Only"
      Top             =   585
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   16745997
      ButtonType      =   2
      Caption         =   "Left"
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
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
   End
   Begin GradButton.GradientButton optRightOnly 
      Height          =   255
      Left            =   5355
      TabIndex        =   31
      ToolTipText     =   "Right Only"
      Top             =   855
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   16745997
      ButtonType      =   2
      Caption         =   "Right"
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
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
   End
   Begin GradButton.GradientButton optAllOn 
      Height          =   255
      Left            =   5355
      TabIndex        =   32
      ToolTipText     =   "All On"
      Top             =   1125
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   16745997
      ButtonType      =   2
      Caption         =   "All"
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
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
      Value           =   -1  'True
   End
   Begin GradButton.GradientButton optMute 
      Height          =   255
      Left            =   5355
      TabIndex        =   33
      ToolTipText     =   "Mute"
      Top             =   1395
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   16745997
      ButtonType      =   2
      Caption         =   "Mute"
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
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
   End
   Begin prjDMP3.SliderControl sldPlsScroll 
      Height          =   2640
      Left            =   5955
      TabIndex        =   23
      Top             =   3000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   4657
      Value           =   1
      Min             =   1
      Max             =   2
      Vertical        =   -1  'True
      Invert          =   -1  'True
   End
   Begin AlfafishEzyID33.EzyID3 ID3 
      Left            =   2070
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin GradButton.GradientButton cmdPlaylist 
      Height          =   255
      Left            =   2355
      TabIndex        =   40
      ToolTipText     =   "Playlist"
      Top             =   2130
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
      Picture         =   "Main.frx":111E
      Style           =   1
   End
   Begin GradButton.GradientButton cmdPlay 
      Height          =   255
      Left            =   490
      TabIndex        =   41
      ToolTipText     =   "Play Media"
      Top             =   2130
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
      Picture         =   "Main.frx":1234
      Style           =   1
   End
   Begin VB.Timer tmrMoveHeaders 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   1305
   End
   Begin GradButton.GradientButton cmdOpenM3U 
      Height          =   255
      Left            =   810
      TabIndex        =   42
      ToolTipText     =   "ID3 1 & 2 Tag Editor"
      Top             =   5670
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BevelWidth      =   1
      BorderColor     =   16744448
      Caption         =   "Open Playlist"
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
      DownForeColor   =   16744448
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
      Style           =   1
   End
   Begin GradButton.GradientButton cmdSaveM3U 
      Height          =   255
      Left            =   2055
      TabIndex        =   43
      ToolTipText     =   "ID3 1 & 2 Tag Editor"
      Top             =   5670
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BevelWidth      =   1
      BorderColor     =   16744448
      Caption         =   "Save Playlist"
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
      DownForeColor   =   16744448
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
      Style           =   1
   End
   Begin GradButton.GradientButton cmdID3 
      Height          =   255
      Left            =   240
      TabIndex        =   44
      ToolTipText     =   "ID3 1 & 2 Tag Editor"
      Top             =   5670
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BevelWidth      =   1
      BorderColor     =   16744448
      Caption         =   "ID3"
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
      DownForeColor   =   16744448
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
      Style           =   1
   End
   Begin GradButton.GradientButton cmdClear 
      Height          =   255
      Left            =   3300
      TabIndex        =   46
      ToolTipText     =   "ID3 1 & 2 Tag Editor"
      Top             =   5670
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      BevelWidth      =   1
      BorderColor     =   16744448
      Caption         =   "Clear"
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
      DownForeColor   =   16744448
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16774357
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
      HoverForeColor  =   16777215
      MaskColor       =   16711935
      Style           =   1
   End
   Begin GradButton.GradientButton chkRepeatAll 
      Height          =   255
      Left            =   3285
      TabIndex        =   47
      ToolTipText     =   "Repeat All"
      Top             =   2130
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      ButtonType      =   1
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
      Picture         =   "Main.frx":1316
      Style           =   1
   End
   Begin GradButton.GradientButton chkRepeatOne 
      Height          =   255
      Left            =   3030
      TabIndex        =   48
      ToolTipText     =   "Repeat One"
      Top             =   2130
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   0
      BackColor       =   13132800
      ButtonType      =   1
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
      Picture         =   "Main.frx":1468
      Style           =   1
   End
   Begin VB.Label lblFileNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   3960
      TabIndex        =   45
      Top             =   5700
      Width           =   1995
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   3195
      TabIndex        =   39
      Top             =   2745
      Width           =   2580
   End
   Begin VB.Label lblArtist 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   675
      TabIndex        =   38
      Top             =   2745
      Width           =   2310
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   135
      TabIndex        =   34
      Top             =   2745
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   7
      X2              =   410.333
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Image imgPlsVisCC 
      Height          =   225
      Left            =   765
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgPlsNotVisCC 
      Height          =   225
      Left            =   495
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5040
      TabIndex        =   29
      Top             =   405
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4770
      TabIndex        =   28
      Top             =   405
      Width           =   75
   End
   Begin VB.Image imgPlsNotVis 
      Height          =   225
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgPlsVis 
      Height          =   225
      Left            =   270
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   405
      TabIndex        =   13
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaning Time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   405
      TabIndex        =   12
      Top             =   810
      Width           =   1020
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E37200&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "Audio Mode (Stereo, etc.)"
      Top             =   915
      Width           =   885
   End
   Begin VB.Label lblKhz 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E37200&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 khz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   2805
      TabIndex        =   8
      ToolTipText     =   "Kilohertz"
      Top             =   915
      Width           =   570
   End
   Begin VB.Label lblKbps 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E37200&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 kbps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF4D5&
      Height          =   210
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Bitrate"
      Top             =   915
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   405
      TabIndex        =   6
      Top             =   390
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DefaultText As String
Const Aliasname As String = "DMP3MEDIA"

Public typeDevice As String
Public Result As String
Public ActualWidth As Single, ActualHeight As Single, TempWidth As Single, TempHeight As Single
Public TotalFrames As String, FPS As String, TotalTime As Single
Public SongTitle As String, SongArtist As String
Private Mode(0 To 3) As String
Private KBPS As Integer, KHZ As Integer
Public Percent As Single, CurrentPosition As Single, CurrentTime As Single
Public ScrollText As String
Public RealScrollText As String
Public Status As String
Public Filename As String, SelectedFilename As String, SelectedIndex As Integer, NewIndex As Integer
Public TempVolume As Long, TempRightVolume As Long, TempLeftVolume As Long
Private TempPlaylistEntry As String
Private PlaylistIndex As Integer
Public Fullscreen As Boolean

Private Sub cmdClear_Click()

    lstPlaylist.ListItems.Clear
    lstFilenames.Clear
    cmdClose_Click
    
End Sub

Private Sub cmdClose_Click()

    If Filename = "" Then Exit Sub

    StopMedia
    CloseMedia
    CloseAll
    
    If frmMedia.Visible = True Then Unload frmMedia
    
    KBPS = 0
    KHZ = 0
    
    Filename = ""
    
    lblKbps.Caption = KBPS & " kbps"
    lblKhz.Caption = KHZ & " khz"
    picTotalTime.Cls
    picTotalTime.Print "00:00:00.00"
    TotalTime = 0
    
    lblMode.Caption = "None"
    
    ScrollText = DefaultText
    
    Dim X
        
    For X = 1 To (picFileContainer.ScaleWidth * 9.5) - Len(ScrollText)
        ScrollText = ScrollText & " "
    Next

End Sub

Private Sub cmdExit_Click()

    StopMedia
    CloseMedia
    CloseAll
    
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
    Next

End Sub

Private Sub cmdFullscreen_Click()

    Dim MovieSize(1 To 4) As Integer

    If LCase$(Right$(Filename, 4)) = ".avi" Or LCase$(Right$(Filename, 4)) = ".mpg" Or LCase$(Right$(Filename, 5)) = ".mpeg" Or LCase$(Right$(Filename, 4)) = ".mpe" Or LCase$(Right$(Filename, 4)) = ".m1v" Or LCase$(Right$(Filename, 4)) = ".mp2" Or LCase$(Right$(Filename, 5)) = ".mpv2" Or LCase$(Right$(Filename, 4)) = ".mpa" Then
        frmMedia.Caption = ""
        MovieSize(1) = 0
        MovieSize(2) = 0
        MovieSize(3) = Screen.Width
        MovieSize(4) = Screen.Height
        AlwaysOnTop frmMedia, True
        frmMedia.Hide
        frmMedia.left = MovieSize(1)
        frmMedia.top = MovieSize(2)
        frmMedia.Width = MovieSize(3)
        frmMedia.Height = MovieSize(4)
        Result = PutMultimedia(frmMedia.hwnd, Aliasname, Val(0), Val(0), Val(0), Val(0))
        frmMedia.Show
        Me.WindowState = vbMinimized
        Fullscreen = True
    End If

End Sub

Private Sub cmdID3_Click()

    If SelectedFilename = "" Then Exit Sub
    
    If SelectedFilename = Filename Then
        MsgBox "This file is in use", vbInformation + vbOKOnly, "DMP3 Alpha 1"
        Exit Sub
    End If
    
    Me.Enabled = False
    
    Load frmId3
    frmId3.Show

End Sub

Private Sub cmdMinimize_Click()

    Dim frm As Form
    
    For Each frm In Forms
        frm.WindowState = vbMinimized
    Next
    
End Sub

Private Sub cmdNext_Click()

    If Filename = "" Then Exit Sub
    If lstFilenames.ListCount = 0 Then Exit Sub
    
    StopMedia
    CloseMedia
    CloseAll
    
    AcquireNextFile
    
End Sub

Private Sub cmdOpen_Click()
   
    Dim U As Integer

    On Error Resume Next 'Dont want any errors

    With FileDialog
    
    'To be able to select multiple files we need to set the flags
    .Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNLongNames + &O4
    
    .MaxFileSize = 20000

    'Clear the filename out of the box
    .Filename = ""
    
    'We want to add files not playlists...
    .DialogTitle = "Add Files"
    
    'Change the filters
    .Filter = "All Media Types|*.ivf;*.aif;*.aifc;*.aiff;*.asf;*.asx;*.wax;*.wma;*.wmv;*.wvx;*.wmp;*.wmx;*.avi;*.wav;*.mpeg;*.mpg;*.miv;*.mp2;*.mp3;*.mpa;*.mpe;*.mpv2;*.mid;*.midi;*.rmi;*.au;*.snd|Intel Video (ivf)|*.ivf|Macintosh AIFF (aiff, aif, aifc)|*.aif;*.aiff;*.aifc|Windows Media (asf, asx, wax, wma, wmv, wvx, wmp, wmx)|*.asf;*.asx;*.wax;*.wma;*.wmv;*.wvx;*.wmp;*.wmx|Windows Formats (avi, wav)|*.avi;*.wav|MPEG (mpeg, mpg, m1v, mp2, mp3, mpa, mpe, mpv2|*.mpeg;*.mpg;*.m1v;*.mp2;*.mp3;*.mpa;*.mpe;*.mpv2|MIDI (mid, midi, rmi)|*.mid;*.midi;*.rmi|UNIX (au, snd)|*.au;*.snd"
    .ShowOpen

    'If you select nothing then exit sub
    If .Filename = "" Then Exit Sub

    Dim I As Integer
    
    If TempVolume = 0 Then TempVolume = 100

    Result = SetVolume(Aliasname, "Both", TempVolume)
    
    sldVolume.Value = TempVolume

    'Parse all of the files in the list
    
    For I = 1 To CountFilesInList(.Filename)
        lstFilenames.AddItem GetFileFromList(.Filename, I)
        lstPlaylist.ListItems.Add lstPlaylist.ListItems.Count + 1, , lstPlaylist.ListItems.Count + 1
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 1, , ""
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 2, , ""
        ParseFiles (lstPlaylist.ListItems.Count)
    Next I
            
    OpenMedia GetFileFromList(.Filename, 1), True
    
    End With
 
    
    On Error GoTo 0

End Sub

Private Sub cmdOpenM3U_Click()

    On Error Resume Next
    
    FileDialog.Filename = ""
    
    FileDialog.Filter = "DMP3 Playlists (dpl)|*.dpl|M3U Playlist (m3u)|*.m3u|PLS Playlist (pls)|*.pls"
    
    FileDialog.CancelError = True
    
    FileDialog.DialogTitle = "Open Playlist"
    
    FileDialog.ShowOpen
    
    If FileDialog.Filename = "" Then Exit Sub
    
    If TempVolume = 0 Then TempVolume = 100

    Result = SetVolume(Aliasname, "Both", TempVolume)
    
    sldVolume.Value = TempVolume
    
    lstPlaylist.ListItems.Clear
    lstFilenames.Clear
    
    If FileDialog.FilterIndex = 1 Then
    
        Call PullFromDPL(FileDialog.Filename)
        
    ElseIf FileDialog.FilterIndex = 2 Then
    
        Call PullFromM3U(FileDialog.Filename)
    
    ElseIf FileDialog.FilterIndex = 3 Then
    
        Call PullFromPLS(FileDialog.Filename)
    
    End If
    
    OpenMedia lstFilenames.List(0), True
    
    On Error GoTo 0

End Sub

Private Sub cmdPause_Click()

    If Filename = "" Then Exit Sub

    If left(Status, 5) = "Pause" Then
        ResumeMedia
    Else
        PauseMedia
    End If

End Sub

Private Sub cmdPlay_Click()

    If Filename = "" Then Exit Sub
    
    If left(Status, 5) = "Pause" Then
        ResumeMedia
    Else
        PlayMedia
    End If
    
End Sub

Private Sub cmdPlaylist_Click()

    If Me.Picture = imgPlsVis.Picture Then
        Me.Picture = imgPlsNotVis.Picture
        Call SetAutoRgn(Me)
    ElseIf Me.Picture = imgPlsNotVis.Picture Then
        Me.Picture = imgPlsVis.Picture
        Call SetAutoRgn(Me)
    ElseIf Me.Picture = imgPlsNotVisCC.Picture Then
        Me.Picture = imgPlsVisCC.Picture
        Call SetAutoRgn(Me)
    ElseIf Me.Picture = imgPlsVisCC.Picture Then
        Me.Picture = imgPlsNotVisCC.Picture
        Call SetAutoRgn(Me)
    End If

End Sub

Private Sub cmdPrev_Click()
    
    If Filename = "" Then Exit Sub
    If lstFilenames.ListCount = 0 Then Exit Sub

    StopMedia
    CloseMedia
    CloseAll
    
    AcquirePrevFile

End Sub

Private Sub cmdSaveM3U_Click()

On Error Resume Next

    Dim iFilenum As Integer
    Dim iCnt As Integer
  
    FileDialog.Filename = ""
    
    FileDialog.Filter = "DMP3 Playlists (dpl)|*.dpl|M3U Playlist (m3u)|*.m3u|PLS Playlist (pls)|*.pls"
    
    FileDialog.CancelError = True
    
    FileDialog.DialogTitle = "Save Playlist"
    
    FileDialog.ShowSave
    
    If FileDialog.Filename = "" Then Exit Sub
    
    If FileDialog.FilterIndex = 1 Then
    
        Call WriteToDPL(FileDialog.Filename)
    
    ElseIf FileDialog.FilterIndex = 2 Then
            
        Call WriteToM3U(FileDialog.Filename)
    
    ElseIf FileDialog.FilterIndex = 3 Then
    
        Call WriteToPLS(FileDialog.Filename)
        
    End If
    
On Error GoTo 0
    
End Sub

Private Sub cmdStop_Click()

    If Filename = "" Then Exit Sub
    
    StopMedia

End Sub

Private Sub cmdVolume_Click()

    If Me.Picture = imgPlsNotVis.Picture Then
        Me.Picture = imgPlsNotVisCC.Picture
        Call SetAutoRgn(Me)
    ElseIf Me.Picture = imgPlsNotVisCC.Picture Then
        Me.Picture = imgPlsNotVis.Picture
        Call SetAutoRgn(Me)
    ElseIf Me.Picture = imgPlsVis.Picture Then
        Me.Picture = imgPlsVisCC.Picture
        Call SetAutoRgn(Me)
    ElseIf Me.Picture = imgPlsVisCC.Picture Then
        Me.Picture = imgPlsVis.Picture
        Call SetAutoRgn(Me)
    End If

End Sub

Private Sub Form_Load()

    Dim X As Integer
    Dim Text As String
    
    Dim DarkColor As OLE_COLOR
    Dim LightColor As OLE_COLOR
    Dim PlaylistColor As OLE_COLOR
    Dim TempTitle As String
    Dim SliderBGColor As OLE_COLOR
    Dim SliderFGColor As OLE_COLOR
    Dim SliderIBGColor As OLE_COLOR
    Dim SliderIFGColor As OLE_COLOR
    Dim TextForeColor As OLE_COLOR

On Error Resume Next

    Open App.Path & "\Gui\ColorConfig.Con" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Text
        If left(Text, 6) = "Title=" Then
            TempTitle = Right(Text, Len(Text) - 6)
        End If
        If left(Text, 10) = "DarkColor=" Then
            DarkColor = Right(Text, Len(Text) - 10)
        End If
        If left(Text, 11) = "LightColor=" Then
            LightColor = Right(Text, Len(Text) - 11)
        End If
        If left(Text, 14) = "PlaylistColor=" Then
            PlaylistColor = Right(Text, Len(Text) - 14)
        End If
        If left(Text, 14) = "SliderBGColor=" Then
            SliderBGColor = Right(Text, Len(Text) - 14)
        End If
        If left(Text, 14) = "SliderFGColor=" Then
            SliderFGColor = Right(Text, Len(Text) - 14)
        End If
        If left(Text, 15) = "SliderIBGColor=" Then
            SliderIBGColor = Right(Text, Len(Text) - 15)
        End If
        If left(Text, 15) = "SliderIFGColor=" Then
            SliderIFGColor = Right(Text, Len(Text) - 15)
        End If
            If left(Text, 14) = "TextForeColor=" Then
            TextForeColor = Right(Text, Len(Text) - 14)
        End If
    Loop
    Close #1

    cmdPrev.BackColor = DarkColor
    cmdPlay.BackColor = DarkColor
    cmdPause.BackColor = DarkColor
    cmdStop.BackColor = DarkColor
    cmdNext.BackColor = DarkColor
    cmdOpen.BackColor = DarkColor
    cmdClose.BackColor = DarkColor
    cmdVolume.BackColor = LightColor
    cmdFullscreen.BackColor = DarkColor
    cmdExit.BackColor = DarkColor
    cmdMinimize.BackColor = DarkColor
    cmdPlaylist.BackColor = DarkColor
    cmdID3.BackColor = DarkColor
    chkRepeatOne.BackColor = DarkColor
    chkRepeatAll.BackColor = DarkColor
    cmdOpenM3U.BackColor = DarkColor
    cmdSaveM3U.BackColor = DarkColor
    cmdClear.BackColor = DarkColor
    picTotalTime.BackColor = LightColor
    picCurrentTime.BackColor = LightColor
    optLeftOnly.BackColor = LightColor
    optRightOnly.BackColor = LightColor
    optAllOn.BackColor = LightColor
    optMute.BackColor = LightColor
    picRemainingTime.BackColor = LightColor
    picFileContainer.BackColor = DarkColor
    lblMode.BackColor = DarkColor
    lblKhz.BackColor = DarkColor
    lblKbps.BackColor = DarkColor
    picScope.BackColor = DarkColor
    lstPlaylist.BackColor = PlaylistColor
    sldVolume.BackColor = SliderBGColor
    sldProgress.BackColor = SliderBGColor
    sldPlsScroll.BackColor = SliderBGColor
    sldRightVolume.BackColor = SliderBGColor
    sldLeftVolume.BackColor = SliderBGColor
    sldVolume.SliderColor = SliderFGColor
    sldProgress.SliderColor = SliderFGColor
    sldPlsScroll.SliderColor = SliderFGColor
    sldRightVolume.SliderColor = SliderFGColor
    sldLeftVolume.SliderColor = SliderFGColor
    sldVolume.InvertedSliderColor = SliderIFGColor
    sldProgress.InvertedSliderColor = SliderIFGColor
    sldPlsScroll.InvertedSliderColor = SliderIFGColor
    sldRightVolume.InvertedSliderColor = SliderIFGColor
    sldLeftVolume.InvertedSliderColor = SliderIFGColor
    sldVolume.InvertedBackColor = SliderIBGColor
    sldProgress.InvertedBackColor = SliderIBGColor
    sldPlsScroll.InvertedBackColor = SliderIBGColor
    sldRightVolume.InvertedBackColor = SliderIBGColor
    sldLeftVolume.InvertedBackColor = SliderIBGColor
    picCurrentTime.ForeColor = TextForeColor
    picRemainingTime.ForeColor = TextForeColor
    picTotalTime.ForeColor = TextForeColor
    picFileContainer.ForeColor = TextForeColor
    lblKbps.ForeColor = TextForeColor
    lblKhz.ForeColor = TextForeColor
    lblMode.ForeColor = TextForeColor
    lblNumber.ForeColor = TextForeColor
    lblArtist.ForeColor = TextForeColor
    lblTitle.ForeColor = TextForeColor
    lblFileNumber.ForeColor = TextForeColor
    cmdOpenM3U.ForeColor = TextForeColor
    cmdSaveM3U.ForeColor = TextForeColor
    cmdClear.ForeColor = TextForeColor
    cmdID3.ForeColor = TextForeColor
    optLeftOnly.ForeColor = TextForeColor
    optRightOnly.ForeColor = TextForeColor
    optAllOn.ForeColor = TextForeColor
    optMute.ForeColor = TextForeColor
    optLeftOnly.DownForeColor = TextForeColor
    optRightOnly.DownForeColor = TextForeColor
    optAllOn.DownForeColor = TextForeColor
    optMute.DownForeColor = TextForeColor
    cmdOpenM3U.DownForeColor = TextForeColor
    cmdSaveM3U.DownForeColor = TextForeColor
    cmdClear.DownForeColor = TextForeColor
    cmdID3.DownForeColor = TextForeColor
    DefaultText = TempTitle

    imgPlsNotVis.Picture = LoadPicture(App.Path & "\Gui\GUINOPLS.GUI")
    imgPlsVis.Picture = LoadPicture(App.Path & "\Gui\GUIPLS.GUI")
    imgPlsNotVisCC.Picture = LoadPicture(App.Path & "\Gui\GUINOPLS_CC.GUI")
    imgPlsVisCC.Picture = LoadPicture(App.Path & "\Gui\GUIPLS_CC.GUI")
    picNumberBar.Picture = LoadPicture(App.Path & "\Gui\HeaderSplit.GUI")
    picArtistBar.Picture = LoadPicture(App.Path & "\Gui\HeaderSplit.GUI")
    picTitleBar.Picture = LoadPicture(App.Path & "\Gui\HeaderSplit.GUI")
        
    Me.Picture = imgPlsNotVis.Picture

    If Me.Picture <> 0 Then
        Call SetAutoRgn(Me)
    End If
    
    CloseMedia
    
    Status = ""
    
    picCurrentTime.Print "00:00:00.00"
    picRemainingTime.Print "00:00:00.00"
    picTotalTime.Print "00:00:00.00"
    
    If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
        SetDefaultDevice "MPEGVideo", "mciqtz.drv"
    End If

    If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
        SetDefaultDevice "avivideo", "mciavi.drv"
    End If
    
    Mode(0) = "Stereo"
    Mode(1) = "Joint Stereo"
    Mode(2) = "Dual Channel"
    Mode(3) = "Mono"
    
    Load frmMedia
    
    RealScrollText = DefaultText
    ScrollText = DefaultText
        
    For X = 1 To (picFileContainer.ScaleWidth * 9) - Len(ScrollText)
        ScrollText = ScrollText & " "
    Next
    
    TempVolume = 100
    TempRightVolume = 100
    TempLeftVolume = 100
    
    NewIndex = 1
    
    On Error GoTo 0

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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
    
End Sub

Public Function OpenMedia(sFilename As String, AutoPlay As Boolean) As Boolean

    On Error Resume Next
    
    StopMedia
    CloseMedia
        
    If LCase$(Right$(sFilename, 4)) = ".avi" Then
        typeDevice = "AviVideo"
    Else
        typeDevice = "MPEGVideo"
    End If
    
    Result = OpenMultimedia(frmMedia.hwnd, Aliasname, sFilename, typeDevice)

    If Result = "Success" Then
    
        Filename = sFilename
        SelectedFilename = sFilename
    
        ActualWidth = GetSize(Aliasname, "cx")
        ActualHeight = GetSize(Aliasname, "cy")
    
        tmrPercent.Enabled = True
        
        TotalFrames = GetTotalframes(Aliasname)
        FPS = GetFramesPerSecond(Aliasname)
        TotalTime = GetTotalTimeByMS(Aliasname) / 1000
        sldProgress.Max = TotalFrames / (FPS * 2)
        sldProgress.Value = 0
        ID3.Filename = sFilename
        ID3.Read
        
        SongTitle = ID3.Song
        SongArtist = ID3.Artist
        
        If Replace(SongTitle, " ", "") = "" Then
            SongTitle = GetFileTitle(sFilename)
        End If
    
        If Replace(SongArtist, " ", "") = "" Then
            SongArtist = "Unknown"
        End If
        
        KBPS = ID3.Bitrate
        
        KHZ = ID3.Frequency / 1000
        
        RealScrollText = SongArtist & " - " & SongTitle
        ScrollText = SongArtist & " - " & SongTitle
        
        lblKbps.Caption = KBPS & " kbps"
        lblKhz.Caption = KHZ & " khz"
        
        lblMode.Caption = Mode(ID3.Mode)
        
        picTotalTime.Cls
        picTotalTime.Print TimeToString(TotalTime)
        
        Dim X
        
        For X = 1 To (picFileContainer.ScaleWidth * 9.5) - Len(ScrollText)
            ScrollText = ScrollText & " "
        Next
        
        If LCase$(Right$(sFilename, 4)) = ".avi" Or LCase$(Right$(sFilename, 4)) = ".mpg" Or LCase$(Right$(sFilename, 5)) = ".mpeg" Or LCase$(Right$(sFilename, 4)) = ".mpe" Or LCase$(Right$(sFilename, 4)) = ".m1v" Or LCase$(Right$(sFilename, 4)) = ".mp2" Or LCase$(Right$(sFilename, 5)) = ".mpv2" Or LCase$(Right$(sFilename, 4)) = ".mpa" Then
            
            If AutoPlay = True Then
                ResizeVideo (0)
                Load frmMedia
                frmMedia.Show
            End If
            
        If LCase$(Right$(sFilename, 4)) = ".avi" Then lblKbps.Caption = "? kbps"
            
        Else
            
            frmMedia.Hide
            
        End If
        
        If AutoPlay = True Then
            PlayMedia
        End If
        
    Else
    
        MsgBox "Cannot play back the file.  The format is not" & vbCrLf & "supported or the selected file is corrupted.", vbExclamation + vbOKOnly, "Error"
        Filename = ""
        
    End If
    
    Status = "Open File: " & Result
    
    On Error GoTo 0

End Function

Public Sub PlayMedia()

    Result = PlayMultimedia(Aliasname, "", "")
    tmrTime.Enabled = True
    tmrPercent.Enabled = True
    
    TempVolume = sldVolume.Value
    Result = SetVolume(Aliasname, "both", TempVolume)

    TempLeftVolume = 100 - sldLeftVolume.Value
    Result = SetVolume(Aliasname, "left", TempLeftVolume)

    TempRightVolume = 100 - sldLeftVolume.Value
    Result = SetVolume(Aliasname, "right", TempRightVolume)

    Status = "Play File: " & Result
    
    If Result = "Success" Then
        
        tmrEndEffects.Enabled = True
        
    End If
    
End Sub

Public Sub PauseMedia()

    Result = PauseMultimedia(Aliasname)
    tmrTime.Enabled = False
    tmrPercent.Enabled = False
    Status = "Pause File: " & Result
    
End Sub

Public Sub StopMedia()

    Result = StopMultimedia(Aliasname)
    tmrTime.Enabled = False
    tmrPercent.Enabled = False
    picCurrentTime.Cls
    picRemainingTime.Cls
    sldProgress.Value = 0
    picCurrentTime.Print "00:00:00.00"
    picRemainingTime.Print "00:00:00.00"
    Status = "Stop File: " & Result
    
End Sub

Public Sub CloseMedia()

    Result = CloseMultimedia(Aliasname)
    tmrEndEffects.Enabled = False
    Status = "Close File: " & Result

End Sub

Public Function TimeToString(CurrTime As Single) As String

  Dim sMinutes As String
  Dim sSeconds As String
  Dim sMilliseconds As String
  Dim sHours As String
  Dim iMinutes As Integer
  Dim iSeconds As Integer
  Dim iHours As Integer

    iHours = Int(CurrTime / 3600)
    iMinutes = Int((CurrTime - iHours * 3600) / 60)
    iSeconds = Int(CurrTime - iHours * 3600 - iMinutes * 60)
    sHours = Format$(Str(iHours), "00")
    TimeToString = sHours & ":"
    sMinutes = Format$(Str(Int(iMinutes)), "00")
    sSeconds = Format$(Str(Int(iSeconds)), "00")
    sMilliseconds = Format$(Str(Int((CurrTime - iHours * 3600 - iMinutes * 60 - iSeconds) * 100)), "00")
    TimeToString = TimeToString & sMinutes & ":" & sSeconds & "." & sMilliseconds

End Function

Private Sub lstPlaylist_Click()

    If lstFilenames.ListCount = 0 Then Exit Sub
    
    SelectedFilename = lstFilenames.List(lstPlaylist.SelectedItem - 1)
    SelectedIndex = lstPlaylist.SelectedItem
    
End Sub

Private Sub lstPlaylist_DblClick()

    If lstFilenames.ListCount = 0 Then Exit Sub
    
    PlaylistIndex = lstPlaylist.SelectedItem
    Filename = lstFilenames.List(lstPlaylist.SelectedItem - 1)
    OpenMedia Filename, True

End Sub

Private Sub lstPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim X As Integer
    
    On Error Resume Next
    
    If KeyCode = vbKeyDelete Then
    
        If lstFilenames.ListCount = 0 Then Exit Sub
        
        Dim I As Integer
        
        If SelectedFilename = Filename Then
            cmdClose_Click
        End If
                
        lstFilenames.RemoveItem CInt(lstPlaylist.SelectedItem.Text) - 1
        lstPlaylist.ListItems.Remove CInt(lstPlaylist.SelectedItem.Text)
       
        For X = 1 To lstPlaylist.ListItems.Count
            lstPlaylist.ListItems(X).Text = X
        Next X
                        
        sldPlsScroll.Max = lstPlaylist.ListItems.Count - 1
        
        If lstFilenames.ListCount = 0 Then
            sldPlsScroll.Max = 2
        End If
        
        lstPlaylist.ListItems.Item(1).Selected = True
        
        SelectedFilename = lstFilenames.List(0)
        
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub optAllOn_Click()

    optMute.Value = False
    optRightOnly.Value = False
    optLeftOnly.Value = False
    
    Result = ChannelsControl(Aliasname, "All", "on")
        
    sldVolume.Value = 100
    TempVolume = 100
    sldLeftVolume.Value = 0
    TempLeftVolume = 100
    sldRightVolume.Value = 0
    TempRightVolume = 100
    
    If Result = "Success" Then

        SetVolume Aliasname, "All", 100

    End If

End Sub

Private Sub optLeftOnly_Click()

    optMute.Value = False
    optRightOnly.Value = False
    optAllOn.Value = False
    
    Result = ChannelsControl(Aliasname, "Left", "on")
    Result = ChannelsControl(Aliasname, "Right", "off")

    sldLeftVolume.Value = 0
    TempLeftVolume = 100
    sldRightVolume.Value = 100
    TempRightVolume = 0
    sldVolume.Value = 50
    TempVolume = 50
    
End Sub

Private Sub optMute_Click()

    optAllOn.Value = False
    optRightOnly.Value = False
    optLeftOnly.Value = False
    
    Result = ChannelsControl(Aliasname, "All", "off")

End Sub

Private Sub optRightOnly_Click()

    optMute.Value = False
    optAllOn.Value = False
    optLeftOnly.Value = False
    
    Result = ChannelsControl(Aliasname, "Right", "on")
    Result = ChannelsControl(Aliasname, "Left", "off")

    sldLeftVolume.Value = 100
    TempLeftVolume = 0
    sldRightVolume.Value = 0
    TempRightVolume = 100
    sldVolume.Value = 50
    TempVolume = 50
    
End Sub

Private Sub sldLeftVolume_MouseDown()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempLeftVolume = 100 - sldLeftVolume.Value
        Result = SetVolume(Aliasname, "left", TempLeftVolume)
    Else
        TempLeftVolume = 100 - sldLeftVolume.Value
    End If
    
End Sub

Private Sub sldLeftVolume_MouseMove()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempLeftVolume = 100 - sldLeftVolume.Value
        Result = SetVolume(Aliasname, "left", TempLeftVolume)
    Else
        TempLeftVolume = 100 - sldLeftVolume.Value
    End If
     
End Sub

Private Sub sldLeftVolume_MouseUp()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempLeftVolume = 100 - sldLeftVolume.Value
        Result = SetVolume(Aliasname, "left", TempLeftVolume)
    Else
        TempLeftVolume = 100 - sldLeftVolume.Value
    End If
    
End Sub

Private Sub sldPlsScroll_MouseMove()

    If lstFilenames.ListCount = 1 Then Exit Sub

    If lstPlaylist.ListItems.Count = 0 Then Exit Sub

    lstPlaylist.ListItems.Item(sldPlsScroll.Value).EnsureVisible
    
End Sub

Private Sub sldProgress_MouseDown()
    
    If Filename = "" Then Exit Sub
    
    tmrPercent.Enabled = False

End Sub

Private Sub sldProgress_MouseUp()

Dim Pos As Long

    If Filename = "" Then Exit Sub

    Pos = sldProgress.Value * (FPS * 2)
    Result = MoveMultimedia(Aliasname, Pos)

    If Result = "Success" Then
    
        tmrPercent.Enabled = True
        
    End If
    
End Sub

Private Sub sldRightVolume_MouseDown()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempRightVolume = 100 - sldRightVolume.Value
        Result = SetVolume(Aliasname, "right", TempRightVolume)
    Else
        TempRightVolume = 100 - sldRightVolume.Value
    End If
    
End Sub

Private Sub sldRightVolume_MouseMove()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempRightVolume = 100 - sldRightVolume.Value
        Result = SetVolume(Aliasname, "right", TempRightVolume)
    Else
        TempRightVolume = 100 - sldRightVolume.Value
    End If
    
End Sub

Private Sub sldRightVolume_MouseUp()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempRightVolume = 100 - sldRightVolume.Value
        Result = SetVolume(Aliasname, "right", TempRightVolume)
    Else
        TempRightVolume = 100 - sldRightVolume.Value
    End If
    
End Sub

Private Sub sldVolume_MouseDown()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempVolume = sldVolume.Value
        Result = SetVolume(Aliasname, "both", TempVolume)
    Else
        TempVolume = sldVolume.Value
    End If

End Sub

Private Sub sldVolume_MouseMove()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempVolume = sldVolume.Value
        Result = SetVolume(Aliasname, "Both", TempVolume)
    Else
        TempVolume = sldVolume.Value
    End If
    
End Sub

Private Sub sldVolume_MouseUp()

    If left(Status, 5) <> "Close" And Status <> "" Then
        TempVolume = sldVolume.Value
        Result = SetVolume(Aliasname, "Both", TempVolume)
    Else
        TempVolume = sldVolume.Value
    End If
    
End Sub

Private Sub tmrEndEffects_Timer()

    If AreMultimediaAtEnd(Aliasname, Val(0)) = True Then
        
        If chkRepeatOne.Value = True And chkRepeatAll.Value = False Then
            PlayMedia
        End If
        If chkRepeatAll.Value = True Then
            If PlaylistIndex = lstPlaylist.ListItems.Count Then
                AcquireNextFile
            End If
        End If
        If chkRepeatAll.Value = False And chkRepeatOne.Value = False And PlaylistIndex = lstPlaylist.ListItems.Count Then
            Exit Sub
        End If
        If chkRepeatAll.Value = False And chkRepeatOne.Value = False And PlaylistIndex <> lstPlaylist.ListItems.Count Then
            AcquireNextFile
        End If
        
    End If
    
End Sub

Private Sub tmrMiscVis_Timer()

    picFileContainer.Cls
    ScrollText = Right(ScrollText, Len(ScrollText) - 1) & left(ScrollText, 1)
    picFileContainer.Print Right(ScrollText, Len(ScrollText) - 1) & left(ScrollText, 1)
    
    lblFileNumber.Caption = lstFilenames.ListCount & " Files"
    
End Sub

Private Sub tmrPercent_Timer()

    Percent = GetPercent(Aliasname)
    If Not Percent = -1 Then sldProgress.Value = Percent * sldProgress.Max \ 100
    CurrentPosition = GetCurrentMultimediaPos(Aliasname)
    CurrentTime = Format(Val(CurrentPosition) / Val(FPS), "00.000")
    
End Sub

Private Sub tmrTime_Timer()
    
    picCurrentTime.Cls
    picRemainingTime.Cls
    
    picCurrentTime.Print TimeToString(Val(CurrentPosition) / Val(FPS))
    picRemainingTime.Print TimeToString(Val(TotalTime) - (Val(CurrentPosition) / Val(FPS)))

End Sub

Public Sub ResizeVideo(AddedSize As Integer)
    
    TempWidth = ActualWidth * 15
        
    TempHeight = ActualHeight * 15
    
    frmMedia.Width = TempWidth + AddedSize
            
    frmMedia.Height = TempHeight + AddedSize
        
    Result = PutMultimedia(frmMedia.hwnd, Aliasname, Val(0), Val(0), Val(0), Val(0))
                
    frmMedia.Height = frmMedia.Height + 350
    
End Sub

Public Sub ResumeMedia()

    Result = ResumeMultimedia(Aliasname)
    tmrTime.Enabled = True
    tmrPercent.Enabled = True
    Status = "Resume File: " & Result

End Sub

Public Sub AcquireNextFile()
    
    If PlaylistIndex = lstPlaylist.ListItems.Count Then
        PlaylistIndex = 1
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    Else
        PlaylistIndex = PlaylistIndex + 1
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    End If
    
End Sub

Public Sub AcquirePrevFile()

    If PlaylistIndex = 1 Then
        PlaylistIndex = lstPlaylist.ListItems.Count
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    Else
        PlaylistIndex = PlaylistIndex - 1
        OpenMedia lstFilenames.List(PlaylistIndex - 1), True
    End If
    
End Sub

Public Function FileTime(Filename As String) As Single
    
    Dim TempTime As Single
        
    If LCase$(Right$(Filename, 4)) = ".avi" Then
        typeDevice = "AviVideo"
    Else
        typeDevice = "MPEGVideo"
    End If
    
    Result = OpenMultimedia(frmMedia.hwnd, "Time", Filename, typeDevice)
    
    TempTime = (GetTotalTimeByMS("Time") / 1000)
    
    If FileExists(Filename) = False Then TempTime = 0

    Result = StopMultimedia("Time")
    Result = CloseMultimedia("Time")

    If TempTime < 0 Then TempTime = 0

    FileTime = TempTime
    
End Function

Private Sub ParseFiles(ListIndex As Integer)
    
    Dim Title As String, Artist As String
    
    ID3.Filename = lstFilenames.List(ListIndex - 1)
    ID3.Read
    
    Title = ID3.Song
    Artist = ID3.Artist
    
    If Replace(Title, " ", "") = "" Then
        Title = GetFileTitle(lstFilenames.List(ListIndex - 1))
    End If
    
    If Replace(Artist, " ", "") = "" Then
        Artist = "Unknown"
    End If
    
    lstPlaylist.ListItems(ListIndex).ListSubItems(1).Text = Artist
    lstPlaylist.ListItems(ListIndex).ListSubItems(2).Text = Title

    sldPlsScroll.Max = lstFilenames.ListCount
    
End Sub

Public Function FileExists(FullFilename As String) As Boolean

    'To make sure that the files actually exist
    
    On Error GoTo MakeF
    Open FullFilename For Input As #1
    Close #1 'Closes the file so we dont get an error
    FileExists = True
    'Very simple, if theres an error then the file does not exist
    
Exit Function

MakeF:
    FileExists = False

Exit Function

End Function

Public Function GetFileTitle(ByVal sFilename As String) As String

Dim lPos As Long, SPos As Long
Dim TempTitle As String

    lPos = InStrRev(sFilename, "\")
    
    If lPos > 0 Then

        If lPos < Len(sFilename) Then
            TempTitle = Mid$(sFilename, lPos + 1)
            SPos = InStrRev(TempTitle, ".")
            GetFileTitle = left$(TempTitle, SPos - 1)
        Else
            GetFileTitle = ""
        End If
        
    Else
    
        GetFileTitle = sFilename
        
    End If
    
End Function

Sub WriteToDPL(Path As String)

On Error Resume Next
    
    Open Path For Output As #1
    Print #1, "<dpl>"
    Print #1, ""

    Dim I As Integer
    
    For I = 0 To lstFilenames.ListCount - 1
        Print #1, "<Entry>"
        Print #1, "<short>" & Chr(34) & GetFileTitle(lstFilenames.List(I)) & Chr(34) & "</short>"
        Print #1, "<url>" & Chr(34) & lstFilenames.List(I) & Chr(34) & "</url>"
        Print #1, "</Entry>"
        Print #1, ""
    Next I
    Print #1, "</dpl>"
    Print #1, "<DPL Script - v1.0>"
    Close #1
    Exit Sub
    
On Error GoTo 0

End Sub

Sub PullFromDPL(Path As String)

    Dim SStart, Getname As Long
    Dim Part2 As Long
    Dim FullFilename As String
    
    On Error Resume Next
    
    SStart = 1
    Getname = 1
    
    Open Path For Input As #1
    TempPlaylistEntry = Input(LOF(1), 1)
    Close #1


    Do Until Getname = "0"
        Getname = InStr(SStart, TempPlaylistEntry, "<url>" & Chr(34))
        Part2 = InStr(Getname + 13, TempPlaylistEntry, Chr(34))
        FullFilename = Mid(TempPlaylistEntry, Getname + 6, Part2 - Getname - 6)
        If Getname <> 0 Then lstFilenames.AddItem FullFilename
        SStart = Getname + 1
        lstPlaylist.ListItems.Add lstPlaylist.ListItems.Count + 1, , lstPlaylist.ListItems.Count + 1
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 1, , ""
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 2, , ""
        ParseFiles (lstPlaylist.ListItems.Count)
        sldPlsScroll.Max = lstFilenames.ListCount
    Loop
    
    lstPlaylist.ListItems.Remove (lstPlaylist.ListItems.Count)
    
    Exit Sub
    
    On Error GoTo 0
    
End Sub

Sub WriteToM3U(Path As String)

On Error Resume Next

    Dim I As Integer

    Open Path For Output As #1
    Print #1, "#EXTM3U"
    For I = 0 To lstFilenames.ListCount - 1
        Print #1, "#EXTINF:" & lstFilenames.List(I)
    Next I
    Close #1
    
On Error GoTo 0
    
End Sub

Sub PullFromM3U(Path As String)

On Error Resume Next

    Dim Filename As String

    Open Path For Input As #1
    Do While Not EOF(1)
        Line Input #1, Filename
        If Right(Filename, 7) = "#EXTM3U" Then
            Line Input #1, Filename
        End If
        Filename = Right$(Filename, Len(Filename) - 8)
        lstFilenames.AddItem Filename
        lstPlaylist.ListItems.Add lstPlaylist.ListItems.Count + 1, , lstPlaylist.ListItems.Count + 1
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 1, , ""
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 2, , ""
        ParseFiles (lstPlaylist.ListItems.Count)
        sldPlsScroll.Max = lstFilenames.ListCount
    Loop
    Close #1
    
On Error GoTo 0
    
End Sub

Sub WriteToPLS(Path As String)

On Error Resume Next

    Dim I As Integer
    

    Open Path For Output As #1
    Print #1, "[Playlist]"
    Print #1, "NumberOfEntries=" & lstFilenames.ListCount
    Print #1, "CachedInfo=1"
    Print #1, "Version=2"
    For I = 1 To lstFilenames.ListCount
        Print #1, "file" & I & "=" & lstFilenames.List(I - 1)
    Next I
    Close #1
    
On Error GoTo 0
    
End Sub

Sub PullFromPLS(Path As String)

On Error Resume Next

    Dim Number As String
    Dim Filename As String
    Dim NumberOfEntries As Integer
    
    Number = 1

    Open Path For Input As #1
    Do While Not EOF(1)
        Line Input #1, Filename
        If Right(Filename, 10) = "[Playlist]" Then
           Line Input #1, Filename
        End If
        
        If left(Filename, 16) = "NumberOfEntries=" Then
            NumberOfEntries = Right(Filename, Len(Filename) - 16)
            Line Input #1, Filename
        End If
        
        If left(Filename, 11) = "CachedInfo=" Then
            Line Input #1, Filename
        End If
        
        If left(Filename, 8) = "Version=" Then
            Line Input #1, Filename
        End If
        
        If left(Filename, 6) = "length" Then
            Line Input #1, Filename
        End If
        
        If left(Filename, 5) = "title" Then
            Line Input #1, Filename
        End If
        
        If left(Filename, 6) = "author" Then
            Line Input #1, Filename
        End If
        
        If Right(Filename, 10) <> "[Playlist]" And left(Filename, 16) <> "NumberOfEntries=" And left(Filename, 11) <> "CachedInfo=" And left(Filename, 8) <> "Version=" And left(Filename, 6) <> "length" And left(Filename, 5) <> "title" And left(Filename, 6) <> "author" Then
            lstFilenames.AddItem Right(Filename, Len(Filename) - Len(Number) - 5)
        End If
        
        lstPlaylist.ListItems.Add lstPlaylist.ListItems.Count + 1, , lstPlaylist.ListItems.Count + 1
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 1, , ""
        lstPlaylist.ListItems(lstPlaylist.ListItems.Count).ListSubItems.Add 2, , ""
        ParseFiles (lstPlaylist.ListItems.Count)
        sldPlsScroll.Max = lstFilenames.ListCount

        Number = Number + 1
    Loop
    Close #1
                
    On Error GoTo 0
    
End Sub
