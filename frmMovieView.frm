VERSION 5.00
Begin VB.Form frmMedia 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMP3 ALPHA - Video"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmMovieView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrKeyPress 
      Interval        =   1
      Left            =   45
      Top             =   45
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub tmrKeyPress_Timer()

If frmMain.Fullscreen = False Then Exit Sub

    If GetAsyncKeyState(vbKeyEscape) Then
        
        If frmFSControls.CShow = False Then
            frmFSControls.CShow = True
        End If
        
        Load frmFSControls
        AlwaysOnTop frmFSControls, True
        frmFSControls.Show
        
    End If
    
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
